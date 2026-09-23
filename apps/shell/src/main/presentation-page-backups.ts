import { createHash, randomUUID } from 'node:crypto'
import { constants } from 'node:fs'
import { lstat, mkdir, open, readdir, rename, rm } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { inflateRawSync } from 'node:zlib'
import { PresentationStore, assertPresentationId } from '@wiswork/project-store'
import { openPptx } from '@wiswork/pptx-engine'

const LIMIT = 8 * 1024 * 1024
const CHUNK = 128 * 1024
const locks = new Map<string, Promise<void>>()
const hash = (value: string | Uint8Array) => createHash('sha256').update(value).digest('hex')
function fail(code: string): never {
  throw new Error(code)
}
const check = (signal: AbortSignal) => {
  if (signal.aborted) fail('aborted')
}
const integer = (value: unknown, min: number, max: number): value is number =>
  typeof value === 'number' && Number.isSafeInteger(value) && value >= min && value <= max
const digest = (value: unknown): value is string =>
  typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
const hostId = (value: unknown): value is string =>
  typeof value === 'string' &&
  value.length > 0 &&
  value.length <= 256 &&
  !Array.from(value).some(
    (char) => char.charCodeAt(0) < 32 || (char.charCodeAt(0) >= 127 && char.charCodeAt(0) <= 159),
  )
interface Metadata {
  backupId: string
  projectId: string
  documentId: string
  requestId: string
  parentRequestId: string
  pageId: string
  hostSlideId: string
  slideIds: string[]
  sha256: string
  sizeBytes: number
  parentInputDigest: string
  inputDigest: string
  status: 'uploading' | 'ready'
}
const beginFields = [
  'backupId',
  'requestId',
  'pageId',
  'hostSlideId',
  'slideIds',
  'sha256',
  'sizeBytes',
]
const metadataKeys = [
  ...beginFields,
  'projectId',
  'documentId',
  'parentRequestId',
  'parentInputDigest',
  'inputDigest',
  'status',
]
function validBegin(value: Record<string, unknown>) {
  for (const key of ['backupId', 'requestId', 'pageId']) assertPresentationId(value[key])
  if (
    !hostId(value.hostSlideId) ||
    !Array.isArray(value.slideIds) ||
    !value.slideIds.length ||
    value.slideIds.length > 512 ||
    !value.slideIds.every(hostId) ||
    new Set(value.slideIds).size !== value.slideIds.length ||
    !value.slideIds.includes(value.hostSlideId) ||
    !digest(value.sha256) ||
    !integer(value.sizeBytes, 1, LIMIT)
  )
    fail('invalid_request')
}
async function directory(path: string) {
  try {
    await mkdir(path, { mode: 0o700 })
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'EEXIST') throw error
  }
  const info = await lstat(path)
  if (!info.isDirectory() || info.isSymbolicLink()) fail('invalid_state')
}
async function syncDirectory(path: string) {
  // Windows does not expose directory fsync through Node; file fsync is still required.
  if (process.platform === 'win32') return
  const handle = await open(path, constants.O_RDONLY | constants.O_NOFOLLOW)
  try {
    await handle.sync()
  } finally {
    await handle.close()
  }
}
async function bytes(path: string, limit: number) {
  const handle = await open(path, constants.O_RDONLY | constants.O_NOFOLLOW)
  try {
    const stat = await handle.stat()
    if (!stat.isFile() || stat.size > limit) fail('invalid_state')
    return await handle.readFile()
  } finally {
    await handle.close()
  }
}
async function atomic(path: string, value: string | Buffer) {
  const temporary = `${path}.${randomUUID()}.tmp`
  const handle = await open(
    temporary,
    constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | constants.O_NOFOLLOW,
    0o600,
  )
  try {
    await handle.writeFile(value)
    await handle.sync()
  } finally {
    await handle.close()
  }
  try {
    await rename(temporary, path)
    await syncDirectory(dirname(path))
  } finally {
    await rm(temporary, { force: true })
  }
}
async function metadata(dir: string): Promise<Metadata> {
  await directory(dir)
  try {
    const value = JSON.parse(
      (await bytes(join(dir, 'metadata.json'), 256 * 1024)).toString(),
    ) as Record<string, unknown>
    if (
      !value ||
      typeof value !== 'object' ||
      Object.keys(value).length !== metadataKeys.length ||
      metadataKeys.some((key) => !Object.hasOwn(value, key))
    )
      fail('invalid_state')
    validBegin(value)
    assertPresentationId(value.projectId)
    assertPresentationId(value.parentRequestId)
    if (
      typeof value.documentId !== 'string' ||
      !value.documentId.length ||
      value.documentId.length > 2048 ||
      !digest(value.parentInputDigest) ||
      !digest(value.inputDigest) ||
      !['ready', 'uploading'].includes(value.status as string)
    )
      fail('invalid_state')
    return value as unknown as Metadata
  } catch {
    return fail('invalid_state')
  }
}
function classicExtra(raw: Buffer, start: number, length: number) {
  const end = start + length
  if (end > raw.length) fail('unsupported_file')
  for (let cursor = start; cursor < end;) {
    if (cursor + 4 > end || raw.readUInt16LE(cursor) === 1) fail('unsupported_file')
    cursor += 4 + raw.readUInt16LE(cursor + 2)
    if (cursor > end) fail('unsupported_file')
  }
}
// Bound each inflation before the existing package reader allocates XML/media.
// Reject ZIP64, encryption, duplicate/unsafe paths and inconsistent local headers.
async function singlePage(raw: Buffer) {
  try {
    let end = -1
    for (let i = raw.length - 22; i >= Math.max(0, raw.length - 65557); i--)
      if (raw.readUInt32LE(i) === 0x06054b50 && i + 22 + raw.readUInt16LE(i + 20) === raw.length) {
        end = i
        break
      }
    if (end < 0 || raw.readUInt16LE(end + 4) !== 0 || raw.readUInt16LE(end + 6) !== 0)
      fail('unsupported_file')
    const count = raw.readUInt16LE(end + 10),
      start = raw.readUInt32LE(end + 16)
    if (
      !count ||
      count > 2048 ||
      count !== raw.readUInt16LE(end + 8) ||
      start + raw.readUInt32LE(end + 12) !== end
    )
      fail('unsupported_file')
    let cursor = start,
      total = 0
    const names = new Set<string>()
    for (let i = 0; i < count; i++) {
      if (cursor + 46 > end || raw.readUInt32LE(cursor) !== 0x02014b50) fail('unsupported_file')
      const flags = raw.readUInt16LE(cursor + 8),
        method = raw.readUInt16LE(cursor + 10),
        compressed = raw.readUInt32LE(cursor + 20),
        expanded = raw.readUInt32LE(cursor + 24),
        local = raw.readUInt32LE(cursor + 42),
        nameLength = raw.readUInt16LE(cursor + 28)
      const next =
        cursor + 46 + nameLength + raw.readUInt16LE(cursor + 30) + raw.readUInt16LE(cursor + 32)
      if (
        next > end ||
        flags & 1 ||
        ![0, 8].includes(method) ||
        raw.readUInt16LE(cursor + 34) !== 0 ||
        expanded > 10 * 1024 * 1024 ||
        total + expanded > 32 * 1024 * 1024 ||
        local + 30 > start ||
        raw.readUInt32LE(local) !== 0x04034b50 ||
        raw.readUInt16LE(local + 8) !== method ||
        raw.readUInt16LE(local + 6) !== flags ||
        raw.readUInt16LE(local + 26) !== nameLength
      )
        fail('unsupported_file')
      const nameBytes = raw.subarray(cursor + 46, cursor + 46 + nameLength),
        name = nameBytes.toString('utf8')
      if (
        !name ||
        name.includes('\\') ||
        name.startsWith('/') ||
        name.split('/').some((part) => part === '..' || part === '.') ||
        name.includes('\0') ||
        names.has(name) ||
        !raw.subarray(local + 30, local + 30 + nameLength).equals(nameBytes)
      )
        fail('unsupported_file')
      names.add(name)
      classicExtra(raw, cursor + 46 + nameLength, raw.readUInt16LE(cursor + 30))
      classicExtra(raw, local + 30 + nameLength, raw.readUInt16LE(local + 28))
      const content = local + 30 + nameLength + raw.readUInt16LE(local + 28)
      if (content + compressed > start) fail('unsupported_file')
      const input = raw.subarray(content, content + compressed),
        inflated =
          method === 0 ? input : inflateRawSync(input, { maxOutputLength: 10 * 1024 * 1024 })
      if (
        inflated.length !== expanded ||
        (name.endsWith('.xml') && /<!DOCTYPE|<!ENTITY/i.test(inflated.toString('utf8')))
      )
        fail('unsupported_file')
      total += expanded
      cursor = next
    }
    if (cursor !== end || !names.has('[Content_Types].xml') || !names.has('ppt/presentation.xml'))
      fail('unsupported_file')
    const opened = await openPptx(raw),
      slides = opened.archive.readPresentation().slidePaths
    if (slides.length !== 1 || !opened.archive.has(slides[0]!) || opened.deck.slides.length !== 1)
      fail('unsupported_file')
  } catch {
    fail('unsupported_file')
  }
}
export function createPresentationPageBackupService(options: { userDataPath: string }) {
  const root = join(resolve(options.userDataPath), 'presentation-page-backups'),
    store = new PresentationStore(options.userDataPath)
  function binding(projectId: string, documentId: string, requestId: string, pageId: string) {
    const child = store.production(projectId, documentId, requestId)
    if (!child) fail('not_found')
    if (!child.revision || child.revision.pageId !== pageId) fail('invalid_request')
    const parent = store.production(projectId, documentId, child.revision.parentRequestId)
    if (
      !parent ||
      child.pages.some((page) => page.state !== 'compiled') ||
      parent.pages.some((page) => page.state !== 'compiled')
    )
      fail('page_not_ready')
    if (
      parent.inputDigest !== child.revision.parentInputDigest ||
      !parent.pages.some((page) => page.pageId === pageId)
    )
      fail('invalid_state')
    return {
      parentRequestId: parent.requestId,
      parentInputDigest: parent.inputDigest,
      inputDigest: child.inputDigest,
    }
  }
  return async (request: Record<string, unknown>, signal: AbortSignal): Promise<unknown> => {
    check(signal)
    const fields: Record<string, string[]> = {
      page_backup_begin: beginFields,
      page_backup_chunk: ['backupId', 'offset', 'base64'],
      page_backup_finish: ['backupId'],
      page_backup_status: ['backupId'],
      page_backup_read: ['backupId', 'offset', 'length'],
    }
    const op = request.operation
    if (typeof op !== 'string' || !Object.hasOwn(fields, op)) fail('invalid_request')
    const allowed = ['operation', 'documentId', 'projectId', ...fields[op]!]
    if (
      Object.keys(request).some((key) => !allowed.includes(key)) ||
      allowed.some((key) => !Object.hasOwn(request, key)) ||
      typeof request.documentId !== 'string' ||
      !request.documentId.length ||
      request.documentId.length > 2048 ||
      Buffer.byteLength(JSON.stringify(request)) > 256 * 1024
    )
      fail('invalid_request')
    assertPresentationId(request.projectId)
    assertPresentationId(request.backupId)
    if (op === 'page_backup_begin') validBegin(request)
    // Snapshot caller-owned arrays before awaiting the project lock.
    const body = structuredClone(request),
      projectId = body.projectId as string,
      documentId = body.documentId as string,
      backupId = body.backupId as string
    const project = join(root, hash(projectId)),
      dir = join(project, hash(backupId)),
      previous = locks.get(project) ?? Promise.resolve()
    let release!: () => void
    const tail = new Promise<void>((resolve) => {
      release = resolve
    })
    locks.set(project, tail)
    await previous
    try {
      check(signal)
      // Check project/document ownership even before looking up backup paths.
      store.production(projectId, documentId)
      await directory(root)
      await directory(project)
      const entries = await readdir(project)
      if (entries.length > 8 || entries.some((entry) => !digest(entry))) fail('invalid_state')
      const exists = entries.includes(hash(backupId))
      if (!exists && op !== 'page_backup_begin') fail('not_found')
      if (!exists) {
        const lineage = binding(
          projectId,
          documentId,
          body.requestId as string,
          body.pageId as string,
        )
        if (entries.length >= 8) fail('quota_exceeded')
        const m: Metadata = {
          backupId,
          projectId,
          documentId,
          requestId: body.requestId as string,
          pageId: body.pageId as string,
          hostSlideId: body.hostSlideId as string,
          slideIds: body.slideIds as string[],
          sha256: body.sha256 as string,
          sizeBytes: body.sizeBytes as number,
          ...lineage,
          status: 'uploading',
        }
        const staging = join(root, `.tmp-${randomUUID()}`)
        await directory(staging)
        try {
          await atomic(join(staging, 'raw.pptx'), Buffer.alloc(0))
          await atomic(join(staging, 'metadata.json'), JSON.stringify(m))
          check(signal)
          await rename(staging, dir)
          await syncDirectory(project)
        } finally {
          await rm(staging, { recursive: true, force: true })
        }
        return { ...m, receivedBytes: 0 }
      }
      let m = await metadata(dir)
      if (m.backupId !== backupId || m.projectId !== projectId) fail('invalid_state')
      if (m.documentId !== documentId) fail('document_mismatch')
      const lineage = binding(projectId, documentId, m.requestId, m.pageId)
      if (
        lineage.parentRequestId !== m.parentRequestId ||
        lineage.parentInputDigest !== m.parentInputDigest ||
        lineage.inputDigest !== m.inputDigest
      )
        fail('invalid_state')
      const path = join(dir, 'raw.pptx'),
        raw = await bytes(path, LIMIT),
        received = raw.length
      if (received > m.sizeBytes || (m.status === 'ready' && received !== m.sizeBytes))
        fail('invalid_state')
      if (m.status === 'ready' && hash(raw) !== m.sha256) fail('digest_mismatch')
      if (op === 'page_backup_begin') {
        if (
          beginFields.some(
            (key) => JSON.stringify(body[key]) !== JSON.stringify(m[key as keyof Metadata]),
          )
        )
          fail('request_conflict')
        return { ...m, receivedBytes: received }
      }
      if (op === 'page_backup_chunk') {
        if (
          !integer(body.offset, 0, m.sizeBytes) ||
          typeof body.base64 !== 'string' ||
          body.base64.length > Math.ceil(CHUNK / 3) * 4
        )
          fail('invalid_request')
        const chunk = Buffer.from(body.base64, 'base64'),
          offset = body.offset
        if (
          !chunk.length ||
          chunk.length > CHUNK ||
          chunk.toString('base64') !== body.base64 ||
          offset + chunk.length > m.sizeBytes
        )
          fail('invalid_request')
        if (offset > received) fail('request_conflict')
        const overlap = Math.min(chunk.length, received - offset)
        if (!raw.subarray(offset, offset + overlap).equals(chunk.subarray(0, overlap)))
          fail('request_conflict')
        check(signal)
        if (overlap < chunk.length) {
          const handle = await open(
            path,
            constants.O_RDWR | constants.O_APPEND | constants.O_NOFOLLOW,
          )
          try {
            const stat = await handle.stat()
            if (!stat.isFile() || stat.size !== received) fail('invalid_state')
            await handle.writeFile(chunk.subarray(overlap))
            await handle.sync()
          } finally {
            await handle.close()
          }
        }
        return { ...m, receivedBytes: Math.max(received, offset + chunk.length) }
      }
      if (op === 'page_backup_finish') {
        if (received !== m.sizeBytes) fail('invalid_state')
        if (hash(raw) !== m.sha256) fail('digest_mismatch')
        await singlePage(raw)
        check(signal)
        if (m.status !== 'ready') {
          m = { ...m, status: 'ready' }
          await atomic(join(dir, 'metadata.json'), JSON.stringify(m))
        }
      }
      if (op === 'page_backup_read') {
        if (!integer(body.offset, 0, m.sizeBytes) || !integer(body.length, 1, CHUNK))
          fail('invalid_request')
        if (m.status !== 'ready') fail('page_not_ready')
        check(signal)
        return {
          backupId,
          offset: body.offset,
          sizeBytes: m.sizeBytes,
          sha256: m.sha256,
          base64: raw.subarray(body.offset, body.offset + body.length).toString('base64'),
        }
      }
      check(signal)
      return { ...m, receivedBytes: received }
    } catch (error) {
      if (
        error instanceof Error &&
        [
          'invalid_request',
          'invalid_state',
          'document_mismatch',
          'request_conflict',
          'not_found',
          'aborted',
          'output_too_large',
          'quota_exceeded',
          'digest_mismatch',
          'unsupported_file',
          'page_not_ready',
        ].includes(error.message)
      )
        throw error
      return fail('invalid_state')
    } finally {
      release()
      if (locks.get(project) === tail) locks.delete(project)
    }
  }
}
