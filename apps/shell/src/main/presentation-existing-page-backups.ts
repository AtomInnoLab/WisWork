import { createHash, randomUUID } from 'node:crypto'
import { constants } from 'node:fs'
import { lstat, mkdir, open, readdir, rename, rm } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { assertPresentationId } from '@wiswork/project-store'
import { validateSinglePageBackupPackage } from './presentation-page-backups'

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
  documentId: string
  hostSlideId: string
  slideIds: string[]
  sha256: string
  sizeBytes: number
  status: 'uploading' | 'ready'
}
const beginFields = [
  'backupId',
  'hostSlideId',
  'slideIds',
  'sha256',
  'sizeBytes',
]
const metadataKeys = [
  ...beginFields,
  'documentId',
  'status',
]
const sameScope = (request: Record<string, unknown>, stored: Metadata) =>
  beginFields.every((key) => JSON.stringify(request[key]) === JSON.stringify(stored[key as keyof Metadata]))
function validBegin(value: Record<string, unknown>) {
  assertPresentationId(value.backupId)
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
    if (
      typeof value.documentId !== 'string' ||
      !value.documentId.length ||
      value.documentId.length > 2048 ||
      !['ready', 'uploading'].includes(value.status as string)
    )
      fail('invalid_state')
    return value as unknown as Metadata
  } catch {
    return fail('invalid_state')
  }
}
export function createPresentationExistingPageBackupService(options: { userDataPath: string }) {
  const root = join(resolve(options.userDataPath), 'presentation-existing-page-backups')
  return async (request: Record<string, unknown>, signal: AbortSignal): Promise<unknown> => {
    check(signal)
    const fields: Record<string, string[]> = {
      existing_page_backup_begin: beginFields,
      existing_page_backup_chunk: ['backupId', 'offset', 'base64'],
      existing_page_backup_finish: ['backupId'],
      existing_page_backup_status: ['backupId'],
      existing_page_backup_read: ['backupId', 'offset', 'length'],
      existing_page_backup_release: beginFields,
      existing_page_backup_list: [],
    }
    const op = request.operation
    if (typeof op !== 'string' || !Object.hasOwn(fields, op)) fail('invalid_request')
    const allowed = ['operation', 'documentId', ...fields[op]!]
    if (
      Object.keys(request).some((key) => !allowed.includes(key)) ||
      allowed.some((key) => !Object.hasOwn(request, key)) ||
      typeof request.documentId !== 'string' ||
      !request.documentId.length ||
      request.documentId.length > 2048 ||
      Buffer.byteLength(JSON.stringify(request)) > 256 * 1024
    )
      fail('invalid_request')
    if (op !== 'existing_page_backup_list') assertPresentationId(request.backupId)
    if (op === 'existing_page_backup_begin' || op === 'existing_page_backup_release') validBegin(request)
    // Snapshot caller-owned arrays before awaiting the document lock.
    const body = structuredClone(request),
      documentId = body.documentId as string,
      backupId = body.backupId as string,
      backupHash = op === 'existing_page_backup_list' ? '' : hash(backupId)
    const document = join(root, hash(documentId)),
      dir = join(document, backupHash),
      receiptDocument = join(root, '.released', hash(documentId)),
      receiptPath = join(receiptDocument, `${backupHash}.json`),
      previous = locks.get(document) ?? Promise.resolve()
    let release!: () => void
    const tail = new Promise<void>((resolve) => {
      release = resolve
    })
    locks.set(document, tail)
    await previous
    try {
      check(signal)
      await directory(root)
      await directory(document)
      await directory(join(root, '.released'))
      await directory(receiptDocument)
      const entries = await readdir(document)
      if (entries.length > 8 || entries.some((entry) => !digest(entry))) fail('invalid_state')
      if (op === 'existing_page_backup_list') {
        const backups = []
        for (const entry of entries) {
          check(signal)
          const stored = await metadata(join(document, entry))
          if (stored.documentId !== documentId || hash(stored.backupId) !== entry) fail('invalid_state')
          backups.push(stored)
        }
        return { documentId, backups }
      }
      const exists = entries.includes(hash(backupId))
      let hasReceipt = false
      try {
        const receiptInfo = await lstat(receiptPath)
        if (!receiptInfo.isFile() || receiptInfo.isSymbolicLink()) fail('invalid_state')
        hasReceipt = true
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
      }
      if (hasReceipt) {
        let stored: Metadata
        try {
          stored = JSON.parse((await bytes(receiptPath, 4096)).toString()) as Metadata
        } catch {
          fail('invalid_state')
        }
        if (stored.backupId !== backupId || stored.documentId !== documentId || stored.status !== 'ready') fail('invalid_state')
        if (op === 'existing_page_backup_begin') fail('request_conflict')
        if (op === 'existing_page_backup_release') {
          if (!sameScope(body, stored)) fail('request_conflict')
          if (exists) {
            // The durable receipt was written only after verifying the ready backup.
            // Retry must finish a deletion interrupted after only some files were removed.
            check(signal)
            await rm(dir, { recursive: true })
            await syncDirectory(document)
          }
          return { ...stored, status: 'released' }
        }
      }
      if (!exists && op !== 'existing_page_backup_begin') fail('not_found')
      if (!exists) {
        if (entries.length >= 8) fail('quota_exceeded')
        const m: Metadata = {
          backupId,
          documentId,
          hostSlideId: body.hostSlideId as string,
          slideIds: body.slideIds as string[],
          sha256: body.sha256 as string,
          sizeBytes: body.sizeBytes as number,
          status: 'uploading',
        }
        const staging = join(root, `.tmp-${randomUUID()}`)
        await directory(staging)
        try {
          await atomic(join(staging, 'raw.pptx'), Buffer.alloc(0))
          await atomic(join(staging, 'metadata.json'), JSON.stringify(m))
          check(signal)
          await rename(staging, dir)
          await syncDirectory(document)
        } finally {
          await rm(staging, { recursive: true, force: true })
        }
        return { ...m, receivedBytes: 0 }
      }
      let m = await metadata(dir)
      if (m.backupId !== backupId) fail('invalid_state')
      if (m.documentId !== documentId) fail('document_mismatch')
      const path = join(dir, 'raw.pptx'),
        raw = await bytes(path, LIMIT),
        received = raw.length
      if (received > m.sizeBytes || (m.status === 'ready' && received !== m.sizeBytes))
        fail('invalid_state')
      if (m.status === 'ready' && hash(raw) !== m.sha256) fail('digest_mismatch')
      if (op === 'existing_page_backup_release') {
        if (!sameScope(body, m)) fail('request_conflict')
        if (m.status !== 'ready') fail('page_not_ready')
        check(signal)
        await atomic(receiptPath, JSON.stringify(m))
        await rm(dir, { recursive: true })
        await syncDirectory(document)
        return { ...m, status: 'released' }
      }
      if (op === 'existing_page_backup_begin') {
        if (!sameScope(body, m))
          fail('request_conflict')
        return { ...m, receivedBytes: received }
      }
      if (op === 'existing_page_backup_chunk') {
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
      if (op === 'existing_page_backup_finish') {
        if (received !== m.sizeBytes) fail('invalid_state')
        if (hash(raw) !== m.sha256) fail('digest_mismatch')
        await validateSinglePageBackupPackage(raw)
        check(signal)
        if (m.status !== 'ready') {
          m = { ...m, status: 'ready' }
          await atomic(join(dir, 'metadata.json'), JSON.stringify(m))
        }
      }
      if (op === 'existing_page_backup_read') {
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
      if (locks.get(document) === tail) locks.delete(document)
    }
  }
}
