import { createHash, randomUUID } from 'node:crypto'
import { constants } from 'node:fs'
import { lstat, mkdir, open, readdir, rename, rm } from 'node:fs/promises'
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
async function directory(path: string, create = true) {
  if (create)
    await mkdir(path, { mode: 0o700 }).catch((e) => {
      if ((e as NodeJS.ErrnoException).code !== 'EEXIST') throw e
    })
  const s = await lstat(path)
  if (s.isSymbolicLink() || !s.isDirectory()) fail('invalid_state')
}
const uuidPattern = '[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}'
const stagePattern = new RegExp('^\\.tmp-' + uuidPattern + '$')
const atomicPattern = new RegExp('^(bundle\\.zip|metadata\\.json)\\.' + uuidPattern + '\\.tmp$')
// Called only while holding the owning project's lock; another project may still be uploading.
async function cleanupTemporary(path: string, stages: boolean) {
  for (const name of await readdir(path)) {
    if (!(stages ? stagePattern : atomicPattern).test(name)) continue
    const temporary = join(path, name)
    const stat = await lstat(temporary)
    if (stat.isSymbolicLink() || (stages ? !stat.isDirectory() : !stat.isFile()))
      fail('invalid_state')
    await rm(temporary, { recursive: stages, force: true })
  }
}
async function bytes(path: string, limit: number) {
  const f = await open(path, constants.O_RDONLY | constants.O_NOFOLLOW | constants.O_NONBLOCK)
  try {
    const stat = await f.stat()
    if (!stat.isFile() || stat.size > limit) fail('invalid_state')
    const raw = await f.readFile()
    if (raw.length > limit) fail('invalid_state')
    return raw
  } finally {
    await f.close()
  }
}
async function atomic(path: string, value: string | Buffer) {
  const tmp = path + '.' + randomUUID() + '.tmp'
  try {
    const f = await open(tmp, 'wx', 0o600)
    try {
      await f.writeFile(value)
      await f.sync()
    } finally {
      await f.close()
    }
    await rename(tmp, path)
    if (process.platform !== 'win32') {
      const parent = await open(dirname(path), constants.O_RDONLY | constants.O_NOFOLLOW)
      try {
        await parent.sync()
      } finally {
        await parent.close()
      }
    }
  } finally {
    await rm(tmp, { force: true })
  }
}
async function save(dir: string, r: PresentationDeliveryBundleReceipt) {
  parsePresentationDeliveryBundleReceipt(r)
  await atomic(
    join(dir, 'metadata.json'),
    JSON.stringify({ receipt: r, checksum: hash(JSON.stringify(r)) }),
  )
}
async function metadata(dir: string) {
  await directory(dir, false)
  await cleanupTemporary(dir, false)
  const v = JSON.parse((await bytes(join(dir, 'metadata.json'), 256 * 1024)).toString())
  if (
    Object.keys(v).sort().join(',') !== 'checksum,receipt' ||
    v.checksum !== hash(JSON.stringify(v.receipt))
  )
    fail('invalid_state')
  const receipt = parsePresentationDeliveryBundleReceipt(v.receipt)
  const raw = await bytes(join(dir, 'bundle.zip'), LIMIT)
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
      count > 10 ||
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
export function createPresentationDeliveryBundleService(options: { userDataPath: string }) {
  const root = join(resolve(options.userDataPath), 'presentation-delivery-bundles')
  const store = new PresentationStore(options.userDataPath)
  return async (request: Record<string, unknown>, signal: AbortSignal): Promise<unknown> => {
    check(signal)
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
    const documentId = request.documentId,
      projectId = request.projectId as string,
      requestId = request.requestId as string
    const project = join(root, hash(documentId), hash(projectId))
    const previous = locks.get(project) ?? Promise.resolve()
    let release!: () => void
    const tail = new Promise<void>((r) => {
      release = r
    })
    locks.set(project, tail)
    await previous
    try {
      check(signal)
      const production = store.production(projectId, documentId, requestId)
      if (!production) fail('not_found')
      await directory(root)
      await directory(join(root, hash(documentId)))
      await directory(project)
      await cleanupTemporary(project, true)
      const entries = await readdir(project)
      if (entries.length > 32 || entries.some((e) => !digest(e))) fail('invalid_state')
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
          const r = await metadata(join(project, entry))
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
          const old = await metadata(dir)
          binding(old)
          if (
            old.bundleId !== bundleId ||
            old.sizeBytes !== receipt.sizeBytes ||
            canonicalPresentationValue(old.manifest) !== canonicalPresentationValue(manifest)
          )
            fail('attachment_conflict')
          return old
        }
        let reserved = 0
        for (const entry of entries) {
          const r = await metadata(join(project, entry))
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
        try {
          await directory(staging)
          await atomic(join(staging, 'bundle.zip'), Buffer.alloc(0))
          await save(staging, receipt)
          check(signal)
          await rename(staging, dir)
        } finally {
          await rm(staging, { recursive: true, force: true })
        }
        return receipt
      }
      if (!exists) fail('not_found')
      let r = await metadata(dir)
      binding(r)
      if (r.bundleId !== bundleId) fail('invalid_state')
      if (op === 'delivery_bundle_delete') {
        check(signal)
        await rm(dir, { recursive: true })
        return { bundleId, deleted: true }
      }
      const raw = await bytes(join(dir, 'bundle.zip'), LIMIT)
      if (
        raw.length < r.receivedBytes ||
        raw.length > r.sizeBytes ||
        (r.state === 'ready' && (raw.length !== r.receivedBytes || hash(raw) !== r.sha256))
      )
        fail('invalid_state')
      if (r.state === 'uploading' && raw.length !== r.receivedBytes) {
        r = { ...r, receivedBytes: raw.length }
        await save(dir, r)
      }
      if (op === 'delivery_bundle_metadata') return r
      if (op === 'delivery_bundle_read') {
        if (r.state !== 'ready') fail('invalid_state')
        if (!integer(request.offset, 0, raw.length) || !integer(request.length, 1, CHUNK))
          fail('invalid_request')
        check(signal)
        return {
          bundleId,
          offset: request.offset,
          totalBytes: raw.length,
          base64: raw.subarray(request.offset, request.offset + request.length).toString('base64'),
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
          !raw.subarray(request.offset, request.offset + overlap).equals(chunk.subarray(0, overlap))
        )
          fail('attachment_conflict')
        if (overlap === chunk.length) return r
        if (r.state !== 'uploading') fail('invalid_state')
        check(signal)
        // ponytail: copying each chunk is bounded by the 20MiB bundle cap; use a journaled append if larger bundles are supported.
        const next = Buffer.concat([raw, chunk.subarray(overlap)])
        await atomic(join(dir, 'bundle.zip'), next)
        const updated = { ...r, receivedBytes: next.length }
        await save(dir, updated)
        return updated
      }
      if (raw.length !== r.sizeBytes) fail('invalid_state')
      if (hash(raw) !== r.sha256) fail('digest_mismatch')
      if (r.state === 'ready') return r
      const zip = await validateBundleZip(raw)
      const names = Object.keys(zip.files).sort()
      if (
        names.join(',') !==
        [...r.manifest.files.map((f) => f.name), 'manifest.json'].sort().join(',')
      )
        fail('unsupported_file')
      const files = new Map<string, Buffer>()
      for (const file of r.manifest.files) {
        const data = await zip.file(file.name)!.async('nodebuffer')
        if (data.length !== file.sizeBytes || hash(data) !== file.sha256) fail('digest_mismatch')
        files.set(file.name, data)
      }
      const manifest = parsePresentationDeliveryBundleManifest(
        JSON.parse(await zip.file('manifest.json')!.async('string')),
      )
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
      if (
        !files
          .get('presentation.pptx')!
          .subarray(0, 4)
          .equals(Buffer.from([80, 75, 3, 4])) ||
        (files.has('presentation.pdf') &&
          !files.get('presentation.pdf')!.subarray(0, 5).equals(Buffer.from('%PDF-')))
      )
        fail('unsupported_file')
      check(signal)
      const ready = {
        ...r,
        state: 'ready' as const,
        completedAt: [new Date().toISOString(), r.createdAt].sort().at(-1)!,
      }
      await save(dir, ready)
      return ready
    } catch (e) {
      const code = e instanceof Error ? e.message : ''
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
        ].includes(code)
      )
        throw e
      return fail('invalid_state')
    } finally {
      release()
      if (locks.get(project) === tail) locks.delete(project)
    }
  }
}
