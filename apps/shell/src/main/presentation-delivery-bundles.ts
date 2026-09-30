import { registerPresentationProjectWork } from './presentation-project-work'
import { parsePresentationResearchRecord } from '@wiswork/project-store/presentation-research'
import { PresentationResearchStore } from '@wiswork/project-store/presentation-research-store'
import { createHash, randomUUID } from 'node:crypto'
import { constants, renameSync, rmSync } from 'node:fs'
import { lstat, mkdir, open, readdir, rm } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { inflateRawSync } from 'node:zlib'
import JSZip from 'jszip'
import { PresentationStore, assertPresentationId } from '@wiswork/project-store'
import {
  parsePresentationDeliveryBundleManifest,
  parsePresentationDeliveryBundleReceipt,
  type PresentationDeliveryBundleReceipt,
} from '@wiswork/project-store/presentation-delivery-bundle'
import { parsePresentationDeliveryReport } from '@wiswork/pptx-engine/presentation-delivery-report'
import { canonicalPresentationValue } from '@wiswork/project-store/presentation-canonical'
const LIMIT = 20 * 1024 * 1024,
  CHUNK = 128 * 1024
const locks = new Map<string, Promise<void>>()
const hash = (v: string | Uint8Array) => createHash('sha256').update(v).digest('hex')
function fail(code: string): never {
  throw new Error(code)
}
const integer = (v: unknown, min: number, max: number): v is number =>
  typeof v === 'number' && Number.isSafeInteger(v) && v >= min && v <= max
const digest = (v: unknown): v is string => typeof v === 'string' && /^[a-f0-9]{64}$/.test(v)
const check = (signal: AbortSignal) => {
  if (signal.aborted) fail('aborted')
}
type Guard = () => void
async function directory(path: string, create = true, guard?: Guard) {
  guard?.()
  if (create)
    await mkdir(path, { mode: 0o700 }).catch((e) => {
      if ((e as NodeJS.ErrnoException).code !== 'EEXIST') throw e
    })
  guard?.()
  const s = await lstat(path)
  guard?.()
  if (s.isSymbolicLink() || !s.isDirectory()) fail('invalid_state')
}
const uuidPattern = '[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}'
const stagePattern = new RegExp('^\\.tmp-' + uuidPattern + '$')
const atomicPattern = new RegExp('^(bundle\\.zip|metadata\\.json)\\.' + uuidPattern + '\\.tmp$')
// Readers validate recognized leftovers without assuming another invocation's ownership.
async function canonicalEntries(path: string, guard?: Guard) {
  const entries = await readdir(path)
  guard?.()
  const canonical: string[] = []
  for (const name of entries) {
    if (stagePattern.test(name)) {
      const stat = await lstat(join(path, name))
      guard?.()
      if (stat.isSymbolicLink() || !stat.isDirectory()) fail('invalid_state')
    } else if (!digest(name)) fail('invalid_state')
    else canonical.push(name)
  }
  if (canonical.length > 32) fail('invalid_state')
  return canonical
}
async function validateTemporaryFiles(path: string, guard?: Guard) {
  const entries = await readdir(path)
  guard?.()
  for (const name of entries.filter((name) => atomicPattern.test(name))) {
    const stat = await lstat(join(path, name))
    guard?.()
    if (stat.isSymbolicLink() || !stat.isFile()) fail('invalid_state')
  }
}
async function bytes(path: string, limit: number, guard?: Guard) {
  guard?.()
  const f = await open(path, constants.O_RDONLY | constants.O_NOFOLLOW | constants.O_NONBLOCK)
  try {
    guard?.()
    const stat = await f.stat()
    guard?.()
    if (!stat.isFile() || stat.size > limit) fail('invalid_state')
    const raw = await f.readFile()
    guard?.()
    if (raw.length > limit) fail('invalid_state')
    return raw
  } finally {
    await f.close()
    guard?.()
  }
}
async function atomic(path: string, value: string | Buffer, guard?: Guard) {
  const tmp = path + '.' + randomUUID() + '.tmp'
  let created = false
  try {
    guard?.()
    const f = await open(tmp, 'wx', 0o600)
    created = true
    try {
      guard?.()
      await f.writeFile(value)
      guard?.()
      await f.sync()
    } finally {
      await f.close()
    }
    guard?.()
    renameSync(tmp, path)
    created = false
    if (process.platform !== 'win32') {
      guard?.()
      const parent = await open(dirname(path), constants.O_RDONLY | constants.O_NOFOLLOW)
      try {
        guard?.()
        await parent.sync()
      } finally {
        await parent.close()
      }
    }
  } finally {
    if (created) await rm(tmp, { force: true })
  }
}
async function save(dir: string, r: PresentationDeliveryBundleReceipt, guard?: Guard) {
  parsePresentationDeliveryBundleReceipt(r)
  await atomic(
    join(dir, 'metadata.json'),
    JSON.stringify({ receipt: r, checksum: hash(JSON.stringify(r)) }),
    guard,
  )
  guard?.()
}
async function metadata(dir: string, guard?: Guard) {
  await directory(dir, false, guard)
  await validateTemporaryFiles(dir, guard)
  const v = JSON.parse((await bytes(join(dir, 'metadata.json'), 256 * 1024, guard)).toString())
  if (
    Object.keys(v).sort().join(',') !== 'checksum,receipt' ||
    v.checksum !== hash(JSON.stringify(v.receipt))
  )
    fail('invalid_state')
  const receipt = parsePresentationDeliveryBundleReceipt(v.receipt)
  const raw = await bytes(join(dir, 'bundle.zip'), LIMIT, guard)
  if (
    raw.length < receipt.receivedBytes ||
    raw.length > receipt.sizeBytes ||
    (receipt.state === 'ready' &&
      (raw.length !== receipt.receivedBytes || hash(raw) !== receipt.sha256))
  )
    fail('invalid_state')
  return receipt
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
async function validateBundleZip(raw: Buffer) {
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
      count > 20 ||
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
        flags & (1 | 64) ||
        ![0, 0x8000].includes((raw.readUInt32LE(cursor + 38) >>> 16) & 0xf000) ||
        ![0, 8].includes(method) ||
        raw.readUInt16LE(cursor + 34) !== 0 ||
        expanded > 20 * 1024 * 1024 ||
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
          method === 0 ? input : inflateRawSync(input, { maxOutputLength: 20 * 1024 * 1024 })
      if (
        inflated.length !== expanded ||
        (name.endsWith('.xml') && /<!DOCTYPE|<!ENTITY/i.test(inflated.toString('utf8')))
      )
        fail('unsupported_file')
      total += expanded
      cursor = next
    }
    if (cursor !== end) fail('unsupported_file')
    return await JSZip.loadAsync(raw, { checkCRC32: true })
  } catch {
    fail('unsupported_file')
  }
}
export function createPresentationDeliveryBundleService(options: {
  userDataPath: string
  acquireProjectLock?(projectId: string): Promise<() => void>
  captureProjectLease?(
    scope: Readonly<{ documentId: string; projectId: string }>,
    mode: 'read' | 'write',
    signal: AbortSignal,
  ): { assertCurrent(): void }
}) {
  options = { ...options, userDataPath: resolve(options.userDataPath) }
  const root = join(resolve(options.userDataPath), 'presentation-delivery-bundles')
  const store = new PresentationStore(options.userDataPath)
  const service = async (
    request: Record<string, unknown>,
    signal: AbortSignal,
  ): Promise<unknown> => {
    check(signal)
    request = structuredClone(request)
    const fields: Record<string, string[]> = {
      delivery_bundle_begin: ['bundleId', 'sha256', 'sizeBytes', 'manifest'],
      delivery_bundle_chunk: ['bundleId', 'offset', 'base64'],
      delivery_bundle_finish: ['bundleId'],
      delivery_bundle_metadata: ['bundleId'],
      delivery_bundle_read: ['bundleId', 'offset', 'length'],
      delivery_bundle_list: [],
      delivery_bundle_delete: ['bundleId'],
    }
    const op = request.operation
    if (typeof op !== 'string' || !Object.hasOwn(fields, op)) fail('invalid_request')
    const allowed = ['operation', 'documentId', 'projectId', 'requestId', ...fields[op]!]
    if (
      Object.keys(request).some((k) => !allowed.includes(k)) ||
      allowed.some((k) => !Object.hasOwn(request, k)) ||
      typeof request.documentId !== 'string' ||
      !request.documentId.trim() ||
      request.documentId.length > 4096 ||
      Buffer.byteLength(JSON.stringify(request)) > 256 * 1024
    )
      fail('invalid_request')
    assertPresentationId(request.projectId)
    assertPresentationId(request.requestId)
    if (op !== 'delivery_bundle_list' && !digest(request.bundleId)) fail('invalid_request')
    const documentId = request.documentId as string,
      projectId = request.projectId as string,
      requestId = request.requestId as string
    const accepted = store.production(projectId, documentId, requestId)
    if (!accepted) fail('not_found')
    const mode = [
      'delivery_bundle_list',
      'delivery_bundle_metadata',
      'delivery_bundle_read',
    ].includes(op)
      ? 'read'
      : 'write'
    const work = registerPresentationProjectWork({
      scope: {
        root: options.userDataPath,
        projectId: accepted.projectId,
        documentId: accepted.documentId,
      },
      signal,
    })
    signal = work.signal
    let releaseProject: (() => void) | undefined
    try {
      const lease = options.captureProjectLease?.(
        Object.freeze({ projectId: accepted.projectId, documentId: accepted.documentId }),
        mode,
        signal,
      )
      const guard = () => {
        check(signal)
        lease?.assertCurrent()
        check(signal)
      }
      guard()
      releaseProject = await options.acquireProjectLock?.(projectId)
      guard()
      const project = join(root, hash(documentId), hash(projectId))
      const previous = locks.get(project) ?? Promise.resolve()
      let release!: () => void
      const tail = new Promise<void>((r) => {
        release = r
      })
      locks.set(project, tail)
      await previous
      try {
        guard()
        const production = store.production(projectId, documentId, requestId)
        if (!production) fail('not_found')
        if (
          production.inputDigest !== accepted.inputDigest ||
          production.planDigest !== accepted.planDigest ||
          production.plan.revision !== accepted.plan.revision
        )
          fail('revision_conflict')
        guard()
        for (const path of [root, join(root, hash(documentId)), project]) {
          try {
            await directory(path, mode === 'write', guard)
          } catch (error) {
            if (mode === 'read' && (error as NodeJS.ErrnoException).code === 'ENOENT') {
              guard()
              if (op === 'delivery_bundle_list') return { bundles: [] }
              fail('not_found')
            }
            throw error
          }
        }
        const entries = await canonicalEntries(project, guard)
        const binding = (r: PresentationDeliveryBundleReceipt) => {
          if (
            r.documentId !== documentId ||
            r.projectId !== projectId ||
            r.requestId !== requestId ||
            r.manifest.planRevision !== production.plan.revision ||
            r.manifest.inputDigest !== production.inputDigest ||
            r.manifest.planDigest !== production.planDigest
          )
            fail('invalid_state')
        }
        if (op === 'delivery_bundle_list') {
          const bundles = []
          for (const entry of entries) {
            const r = await metadata(join(project, entry), guard)
            if (
              hash(r.bundleId) !== entry ||
              r.documentId !== documentId ||
              r.projectId !== projectId
            )
              fail('invalid_state')
            if (r.requestId === requestId) {
              binding(r)
              bundles.push(r)
            }
          }
          guard()
          return { bundles }
        }
        const bundleId = request.bundleId as string,
          dir = join(project, hash(bundleId)),
          exists = entries.includes(hash(bundleId))
        if (op === 'delivery_bundle_begin') {
          const manifest = parsePresentationDeliveryBundleManifest(request.manifest)
          if (request.sha256 !== bundleId || !integer(request.sizeBytes, 1, LIMIT))
            fail('invalid_request')
          const receipt: PresentationDeliveryBundleReceipt = {
            version: 1,
            documentId,
            projectId,
            requestId,
            bundleId,
            sha256: bundleId,
            sizeBytes: request.sizeBytes,
            receivedBytes: 0,
            state: 'uploading',
            createdAt: new Date().toISOString(),
            manifest,
          }
          binding(receipt)
          parsePresentationDeliveryBundleReceipt(receipt)
          if (exists) {
            const old = await metadata(dir, guard)
            binding(old)
            if (
              old.bundleId !== bundleId ||
              old.sizeBytes !== receipt.sizeBytes ||
              canonicalPresentationValue(old.manifest) !== canonicalPresentationValue(manifest)
            )
              fail('attachment_conflict')
            guard()
            return old
          }
          let reserved = 0
          for (const entry of entries) {
            const r = await metadata(join(project, entry), guard)
            if (
              r.documentId !== documentId ||
              r.projectId !== projectId ||
              hash(r.bundleId) !== entry
            )
              fail('invalid_state')
            reserved += r.sizeBytes
          }
          if (entries.length >= 32 || reserved + receipt.sizeBytes > 100 * 1024 * 1024)
            fail('quota_exceeded')
          const staging = join(project, '.tmp-' + randomUUID())
          let created = false
          try {
            guard()
            await mkdir(staging, { mode: 0o700 })
            created = true
            await directory(staging, false, guard)
            await atomic(join(staging, 'bundle.zip'), Buffer.alloc(0), guard)
            await save(staging, receipt, guard)
            guard()
            renameSync(staging, dir)
            created = false
          } finally {
            if (created) await rm(staging, { recursive: true, force: true })
          }
          guard()
          return receipt
        }
        if (!exists) fail('not_found')
        let r = await metadata(dir, guard)
        binding(r)
        if (r.bundleId !== bundleId) fail('invalid_state')
        if (op === 'delivery_bundle_delete') {
          guard()
          rmSync(dir, { recursive: true })
          return { bundleId, deleted: true }
        }
        const raw = await bytes(join(dir, 'bundle.zip'), LIMIT, guard)
        if (
          raw.length < r.receivedBytes ||
          raw.length > r.sizeBytes ||
          (r.state === 'ready' && (raw.length !== r.receivedBytes || hash(raw) !== r.sha256))
        )
          fail('invalid_state')
        if (mode === 'write' && r.state === 'uploading' && raw.length !== r.receivedBytes) {
          r = { ...r, receivedBytes: raw.length }
          await save(dir, r, guard)
        }
        if (op === 'delivery_bundle_metadata') {
          guard()
          return r
        }
        if (op === 'delivery_bundle_read') {
          if (r.state !== 'ready') fail('invalid_state')
          if (!integer(request.offset, 0, raw.length) || !integer(request.length, 1, CHUNK))
            fail('invalid_request')
          check(signal)
          return {
            bundleId,
            offset: request.offset,
            totalBytes: raw.length,
            base64: raw
              .subarray(request.offset, request.offset + request.length)
              .toString('base64'),
          }
        }
        if (op === 'delivery_bundle_chunk') {
          if (
            !integer(request.offset, 0, r.sizeBytes) ||
            typeof request.base64 !== 'string' ||
            request.base64.length > Math.ceil(CHUNK / 3) * 4 ||
            !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(request.base64)
          )
            fail('invalid_request')
          const chunk = Buffer.from(request.base64, 'base64')
          if (
            !chunk.length ||
            chunk.length > CHUNK ||
            request.offset > raw.length ||
            request.offset + chunk.length > r.sizeBytes
          )
            fail('invalid_request')
          const overlap = Math.min(chunk.length, raw.length - request.offset)
          if (
            !raw
              .subarray(request.offset, request.offset + overlap)
              .equals(chunk.subarray(0, overlap))
          )
            fail('attachment_conflict')
          if (overlap === chunk.length) {
            guard()
            return r
          }
          if (r.state !== 'uploading') fail('invalid_state')
          check(signal)
          // ponytail: copying each chunk is bounded by the 20MiB bundle cap; use a journaled append if larger bundles are supported.
          const next = Buffer.concat([raw, chunk.subarray(overlap)])
          await atomic(join(dir, 'bundle.zip'), next, guard)
          const updated = { ...r, receivedBytes: next.length }
          await save(dir, updated, guard)
          return updated
        }
        if (raw.length !== r.sizeBytes) fail('invalid_state')
        if (hash(raw) !== r.sha256) fail('digest_mismatch')
        if (r.state === 'ready') {
          guard()
          return r
        }
        const zip = await validateBundleZip(raw)
        guard()
        const names = Object.keys(zip.files).sort()
        if (
          names.join(',') !==
          [...r.manifest.files.map((f) => f.name), 'manifest.json'].sort().join(',')
        )
          fail('unsupported_file')
        const files = new Map<string, Buffer>()
        for (const file of r.manifest.files) {
          const data = await zip.file(file.name)!.async('nodebuffer')
          guard()
          if (data.length !== file.sizeBytes || hash(data) !== file.sha256) fail('digest_mismatch')
          files.set(file.name, data)
        }
        const manifest = parsePresentationDeliveryBundleManifest(
          JSON.parse(await zip.file('manifest.json')!.async('string')),
        )
        guard()
        if (canonicalPresentationValue(manifest) !== canonicalPresentationValue(r.manifest))
          fail('attachment_conflict')
        const evidence = parsePresentationDeliveryReport(
          JSON.parse(files.get('evidence.json')!.toString()),
        )
        if (
          evidence.documentId !== documentId ||
          evidence.projectId !== projectId ||
          evidence.requestId !== requestId ||
          evidence.planRevision !== production.plan.revision ||
          evidence.inputDigest !== production.inputDigest ||
          evidence.planDigest !== production.planDigest ||
          canonicalPresentationValue(evidence.plan) !==
            canonicalPresentationValue(production.plan.plan) ||
          evidence.pages.map((p) => p.pageId).join(',') !==
            production.pages.map((p) => p.pageId).join(',')
        )
          fail('invalid_state')
        if (
          canonicalPresentationValue(JSON.parse(files.get('claims.json')!.toString())) !==
            canonicalPresentationValue(evidence.plan.claims) ||
          canonicalPresentationValue(JSON.parse(files.get('sources.json')!.toString())) !==
            canonicalPresentationValue(evidence.plan.sources)
        )
          fail('invalid_state')
        if (evidence.plan.research && !files.has('research.json')) fail('invalid_state')
        if (files.has('research.json')) {
          const research = parsePresentationResearchRecord(
            JSON.parse(files.get('research.json')!.toString()),
          )
          if (
            research.documentId !== documentId ||
            research.projectId !== projectId ||
            research.state !== 'completed'
          )
            fail('invalid_state')
          if (
            evidence.plan.research &&
            canonicalPresentationValue(research) !==
              canonicalPresentationValue(evidence.research?.record)
          )
            fail('invalid_state')
          const original = await new PresentationResearchStore(options.userDataPath).read(
            documentId,
            projectId,
            research.id,
          )
          guard()
          if (canonicalPresentationValue(research) !== canonicalPresentationValue(original))
            fail('invalid_state')
        }
        if (
          !files
            .get('presentation.pptx')!
            .subarray(0, 4)
            .equals(Buffer.from([80, 75, 3, 4])) ||
          (files.has('presentation.pdf') &&
            !files.get('presentation.pdf')!.subarray(0, 5).equals(Buffer.from('%PDF-')))
        )
          fail('unsupported_file')
        for (let page = 1; page <= 8; page++) {
          const image = files.get(`page-${page}.png`)
          if (!image) continue
          if (
            image.length < 24 ||
            !image.subarray(0, 8).equals(Buffer.from('89504e470d0a1a0a', 'hex')) ||
            image.toString('ascii', 12, 16) !== 'IHDR' ||
            image.readUInt32BE(16) < 1 ||
            image.readUInt32BE(20) < 1 ||
            image.readUInt32BE(16) * image.readUInt32BE(20) > 16_000_000
          )
            fail('unsupported_file')
        }
        check(signal)
        const ready = {
          ...r,
          state: 'ready' as const,
          completedAt: [new Date().toISOString(), r.createdAt].sort().at(-1)!,
        }
        await save(dir, ready, guard)
        return ready
      } catch (e) {
        let code = ''
        try {
          const message = e instanceof Error ? e.message : undefined
          if (typeof message === 'string') code = message
        } catch {}
        if (
          [
            'invalid_request',
            'invalid_state',
            'not_found',
            'attachment_conflict',
            'quota_exceeded',
            'unsupported_file',
            'digest_mismatch',
            'aborted',
            'revision_conflict',
            'project_deleting',
            'project_deleted',
            'project_not_found',
            'document_mismatch',
          ].includes(code)
        )
          throw e
        return fail('invalid_state')
      } finally {
        release()
        if (locks.get(project) === tail) locks.delete(project)
      }
    } finally {
      releaseProject?.()
      work.finish()
    }
  }
  return Object.assign(service, {
    async assertResearchCleanupAvailable(documentId: string, projectId: string) {
      const project = join(root, hash(documentId), hash(projectId))
      for (const path of [root, join(root, hash(documentId)), project]) {
        try {
          await directory(path, false)
        } catch (error) {
          if ((error as NodeJS.ErrnoException).code === 'ENOENT') return
          throw new Error('invalid_state', { cause: error })
        }
      }
      const entries = await canonicalEntries(project)
      for (const entry of entries) {
        const receipt = await metadata(join(project, entry))
        if (
          hash(receipt.bundleId) !== entry ||
          receipt.documentId !== documentId ||
          receipt.projectId !== projectId
        )
          fail('invalid_state')
        if (
          receipt.state !== 'ready' &&
          receipt.manifest.files.some((file) => file.name === 'research.json')
        )
          fail('busy')
      }
    },
  })
}
