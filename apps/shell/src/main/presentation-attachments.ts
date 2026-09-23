import { createHash, randomUUID } from 'node:crypto'
import { constants } from 'node:fs'
import { lstat, mkdir, open, readdir, rename, rm } from 'node:fs/promises'
import { extname, join, resolve } from 'node:path'
import { inflateRawSync } from 'node:zlib'
import { parseFileToText } from '@wiswork/file-parse'
import {
  inspectPresentationImage,
  normalizePresentationImage,
  PRESENTATION_IMAGE_INPUT_LIMIT,
  PRESENTATION_IMAGE_CACHE_LIMIT,
} from './presentation-image'

const FILE_LIMIT = 50 * 1024 * 1024
const CHUNK_LIMIT = 128 * 1024
const TEXT_LIMIT = 1_000_000
const locks = new Map<string, Promise<void>>()
const hash = (bytes: string | Uint8Array) => createHash('sha256').update(bytes).digest('hex')
function fail(code: string): never {
  throw new Error(code)
}
const checkAbort = (signal: AbortSignal) => {
  if (signal.aborted) fail('aborted')
}
const isId = (id: unknown): id is string => typeof id === 'string' && /^[a-f0-9]{64}$/.test(id)
const integer = (n: unknown, min: number, max: number): n is number =>
  typeof n === 'number' && Number.isSafeInteger(n) && n >= min && n <= max
function filename(name: unknown): name is string {
  return (
    typeof name === 'string' &&
    name.length > 0 &&
    name.length <= 180 &&
    !/[\\/]/.test(name) &&
    !Array.from(name).some((char) => char.charCodeAt(0) < 32 || char.charCodeAt(0) === 127) &&
    name !== '.' &&
    name !== '..'
  )
}
function imageFile(name: string) {
  return ['.png', '.jpg', '.jpeg'].includes(extname(name).toLowerCase())
}
function fileLimit(name: string) {
  return imageFile(name) ? PRESENTATION_IMAGE_INPUT_LIMIT : FILE_LIMIT
}
function supported(name: string) {
  return (
    imageFile(name) ||
    ['.pdf', '.docx', '.txt', '.md', '.csv', '.json'].includes(extname(name).toLowerCase())
  )
}
interface Metadata {
  attachmentId: string
  name: string
  sizeBytes: number
  sha256: string
  status: 'uploading' | 'ready' | 'failed'
  kind?: 'text' | 'image'
  mime?: 'image/png'
  width?: number
  height?: number
  assetSha256?: string
  error?: string
  totalChars?: number
  textDigest?: string
}
async function directory(path: string) {
  try {
    await mkdir(path, { mode: 0o700 })
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code !== 'EEXIST') throw e
  }
  const info = await lstat(path)
  if (!info.isDirectory() || info.isSymbolicLink()) fail('invalid_state')
}
async function bytes(path: string, limit: number): Promise<Buffer> {
  const file = await open(path, constants.O_RDONLY | constants.O_NOFOLLOW)
  try {
    const stat = await file.stat()
    if (!stat.isFile() || stat.size > limit) fail('invalid_state')
    return await file.readFile()
  } finally {
    await file.close()
  }
}
async function rawSize(path: string): Promise<number> {
  const handle = await open(path, constants.O_RDONLY | constants.O_NOFOLLOW)
  try {
    const stat = await handle.stat()
    if (!stat.isFile() || stat.size > FILE_LIMIT) fail('invalid_state')
    return stat.size
  } finally {
    await handle.close()
  }
}
// Acknowledgements follow fsync. A crash can leave a chunk prefix on disk;
// replay validates that prefix and appends only the missing suffix.
async function appendChunk(path: string, offset: number, chunk: Buffer, received: number) {
  const handle = await open(path, constants.O_RDWR | constants.O_APPEND | constants.O_NOFOLLOW)
  try {
    const stat = await handle.stat()
    if (!stat.isFile() || stat.size !== received) fail('invalid_state')
    const overlap = Math.min(chunk.length, received - offset)
    if (overlap > 0) {
      const prefix = Buffer.alloc(overlap)
      const result = await handle.read(prefix, 0, overlap, offset)
      if (result.bytesRead !== overlap || !prefix.equals(chunk.subarray(0, overlap)))
        fail('attachment_conflict')
    }
    if (overlap < chunk.length) {
      await handle.writeFile(chunk.subarray(overlap))
      await handle.sync()
    }
  } finally {
    await handle.close()
  }
}
async function atomic(path: string, value: string | Buffer) {
  const temp = `${path}.${randomUUID()}.tmp`
  const handle = await open(
    temp,
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
    await rename(temp, path)
  } finally {
    await rm(temp, { force: true })
  }
}
async function metadata(dir: string, id: string): Promise<Metadata> {
  await directory(dir)
  let m: Metadata
  try {
    m = JSON.parse((await bytes(join(dir, 'metadata.json'), 4096)).toString('utf8')) as Metadata
  } catch {
    return fail('invalid_state')
  }
  if (
    !m ||
    m.attachmentId !== id ||
    m.sha256 !== id ||
    !filename(m.name) ||
    !supported(m.name) ||
    !integer(m.sizeBytes, 0, fileLimit(m.name)) ||
    !['uploading', 'ready', 'failed'].includes(m.status)
  )
    fail('invalid_state')
  if (m.status === 'ready') {
    if (imageFile(m.name)) {
      if (
        m.kind !== 'image' ||
        m.mime !== 'image/png' ||
        !integer(m.width, 1, 8192) ||
        !integer(m.height, 1, 8192) ||
        m.width * m.height > 16_000_000 ||
        !isId(m.assetSha256) ||
        m.totalChars !== undefined ||
        m.textDigest !== undefined
      )
        fail('invalid_state')
    } else if (m.kind !== 'text' || !integer(m.totalChars, 0, TEXT_LIMIT) || !isId(m.textDigest))
      fail('invalid_state')
  }
  if (m.status === 'failed' && m.error !== 'parse_failed') fail('invalid_state')
  return m
}
const publicMetadata = (m: Metadata, receivedBytes: number) => ({
  attachmentId: m.attachmentId,
  name: m.name,
  sizeBytes: m.sizeBytes,
  sha256: m.sha256,
  receivedBytes,
  status: m.status,
  ...(m.kind ? { kind: m.kind } : {}),
  ...(m.error ? { error: m.error } : {}),
  ...(m.totalChars !== undefined ? { totalChars: m.totalChars } : {}),
  ...(m.kind === 'image'
    ? { mime: m.mime, width: m.width, height: m.height, assetSha256: m.assetSha256 }
    : {}),
})

async function cachedImage(dir: string, m: Metadata): Promise<Buffer> {
  if (m.status !== 'ready' || m.kind !== 'image') fail('invalid_state')
  const value = await bytes(join(dir, 'image.png'), PRESENTATION_IMAGE_CACHE_LIMIT)
  const info = inspectPresentationImage(value)
  if (
    hash(value) !== m.assetSha256 ||
    info.mime !== 'image/png' ||
    info.width !== m.width ||
    info.height !== m.height
  )
    fail('invalid_state')
  return value
}

// Preflight every ZIP member with bounded inflation, before the DOCX parser allocates XML.
// Classic ZIP only: ZIP64/encryption/unknown compression fail closed. PDF parsing is
// still in-process; its 50 MiB input and 1M output caps are not a CPU/memory sandbox.
function checkDocx(data: Buffer) {
  let end = -1
  for (let i = data.length - 22; i >= Math.max(0, data.length - 65557); i--) {
    if (data.readUInt32LE(i) === 0x06054b50 && i + 22 + data.readUInt16LE(i + 20) === data.length) {
      end = i
      break
    }
  }
  if (end < 0 || data.readUInt16LE(end + 4) !== 0 || data.readUInt16LE(end + 6) !== 0)
    fail('parse_failed')
  const count = data.readUInt16LE(end + 10)
  let cursor = data.readUInt32LE(end + 16),
    total = 0
  if (
    count > 2048 ||
    count !== data.readUInt16LE(end + 8) ||
    cursor + data.readUInt32LE(end + 12) !== end
  )
    fail('parse_failed')
  for (let i = 0; i < count; i++) {
    if (cursor + 46 > end || data.readUInt32LE(cursor) !== 0x02014b50) fail('parse_failed')
    const flags = data.readUInt16LE(cursor + 8),
      method = data.readUInt16LE(cursor + 10)
    const compressed = data.readUInt32LE(cursor + 20),
      expanded = data.readUInt32LE(cursor + 24)
    const local = data.readUInt32LE(cursor + 42)
    if (
      flags & 1 ||
      ![0, 8].includes(method) ||
      expanded > 10 * 1024 * 1024 ||
      total + expanded > 32 * 1024 * 1024 ||
      local + 30 > cursor ||
      data.readUInt32LE(local) !== 0x04034b50
    )
      fail('parse_failed')
    const start = local + 30 + data.readUInt16LE(local + 26) + data.readUInt16LE(local + 28)
    if (start + compressed > cursor) fail('parse_failed')
    const input = data.subarray(start, start + compressed)
    const inflated =
      method === 0 ? input : inflateRawSync(input, { maxOutputLength: 10 * 1024 * 1024 })
    if (inflated.length !== expanded) fail('parse_failed')
    total += inflated.length
    cursor +=
      46 +
      data.readUInt16LE(cursor + 28) +
      data.readUInt16LE(cursor + 30) +
      data.readUInt16LE(cursor + 32)
  }
  if (cursor !== end) fail('parse_failed')
}

export function createPresentationAttachmentService(options: {
  userDataPath: string
  parse?: typeof parseFileToText
  normalizeImage?: typeof normalizePresentationImage
}) {
  const root = join(resolve(options.userDataPath), 'presentation-attachments')
  const parse = options.parse ?? parseFileToText
  const normalizeImage = options.normalizeImage ?? normalizePresentationImage
  return async (body: Record<string, unknown>, signal: AbortSignal): Promise<unknown> => {
    checkAbort(signal)
    const fields: Record<string, string[]> = {
      attachment_begin: ['attachmentId', 'name', 'sizeBytes', 'sha256'],
      attachment_chunk: ['attachmentId', 'offset', 'base64'],
      attachment_finish: ['attachmentId'],
      attachment_list: [],
      attachment_list_assets: [],
      attachment_asset: ['attachmentId'],
      attachment_read: ['attachmentId', 'offset', 'maxChars'],
    }
    const op = body.operation
    if (
      typeof op !== 'string' ||
      !Object.hasOwn(fields, op) ||
      typeof body.documentId !== 'string' ||
      !body.documentId.length ||
      body.documentId.length > 2048
    )
      fail('invalid_request')
    const allowed = ['operation', 'documentId', ...fields[op]!]
    if (
      Object.keys(body).some((k) => !allowed.includes(k)) ||
      allowed.some((k) => !Object.hasOwn(body, k))
    )
      fail('invalid_request')
    if (!['attachment_list', 'attachment_list_assets'].includes(op) && !isId(body.attachmentId))
      fail('invalid_request')
    if (op === 'attachment_begin') {
      if (
        !filename(body.name) ||
        body.sha256 !== body.attachmentId ||
        !integer(body.sizeBytes, 0, fileLimit(body.name))
      )
        fail('invalid_request')
      if (!supported(body.name)) fail('unsupported_file')
    }
    const doc = join(root, hash(body.documentId))
    const previous = locks.get(doc) ?? Promise.resolve()
    let release!: () => void
    const tail = new Promise<void>((r) => {
      release = r
    })
    locks.set(doc, tail)
    await previous
    try {
      checkAbort(signal)
      await directory(root)
      await directory(doc)
      const entries = await readdir(doc)
      if (entries.length > 32 || entries.some((e) => !isId(e))) fail('invalid_state')
      const id = body.attachmentId as string
      const dir = id ? join(doc, id) : doc
      const exists = entries.includes(id)
      if (op === 'attachment_list' || op === 'attachment_list_assets') {
        const attachments = []
        for (const entry of entries) {
          const m = await metadata(join(doc, entry), entry)
          if (op === 'attachment_list' && imageFile(m.name)) continue
          const received = await rawSize(join(doc, entry, `raw${extname(m.name).toLowerCase()}`))
          if (received > m.sizeBytes) fail('invalid_state')
          attachments.push(publicMetadata(m, received))
        }
        checkAbort(signal)
        return { attachments }
      }
      if (!exists && op !== 'attachment_begin') fail('not_found')
      if (op === 'attachment_begin' && !exists) {
        let declared = 0
        for (const entry of entries) declared += (await metadata(join(doc, entry), entry)).sizeBytes
        if (entries.length >= 32 || declared + (body.sizeBytes as number) > 100 * 1024 * 1024)
          fail('quota_exceeded')
        checkAbort(signal)
        const m: Metadata = {
          attachmentId: id,
          sha256: id,
          name: body.name as string,
          sizeBytes: body.sizeBytes as number,
          status: 'uploading',
        }
        const staging = join(root, `.tmp-${randomUUID()}`)
        await directory(staging)
        try {
          await atomic(join(staging, `raw${extname(m.name).toLowerCase()}`), Buffer.alloc(0))
          await atomic(join(staging, 'metadata.json'), JSON.stringify(m))
          await rename(staging, dir)
        } finally {
          await rm(staging, { recursive: true, force: true })
        }
        return publicMetadata(m, 0)
      }
      let m = await metadata(dir, id)
      const rawPath = join(dir, `raw${extname(m.name).toLowerCase()}`)
      const received = await rawSize(rawPath)
      if (received > m.sizeBytes) fail('invalid_state')
      if (op === 'attachment_begin') {
        if (m.name !== body.name || m.sizeBytes !== body.sizeBytes) fail('attachment_conflict')
        return publicMetadata(m, received)
      }
      if (op === 'attachment_chunk') {
        if (
          !integer(body.offset, 0, m.sizeBytes) ||
          typeof body.base64 !== 'string' ||
          body.base64.length > Math.ceil(CHUNK_LIMIT / 3) * 4 ||
          !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(body.base64)
        )
          fail('invalid_request')
        const chunk = Buffer.from(body.base64, 'base64'),
          offset = body.offset
        if (
          !chunk.length ||
          chunk.length > CHUNK_LIMIT ||
          chunk.toString('base64') !== body.base64 ||
          offset + chunk.length > m.sizeBytes
        )
          fail('invalid_request')
        if (offset > received || (m.status !== 'uploading' && offset + chunk.length > received))
          fail('attachment_conflict')
        checkAbort(signal)
        await appendChunk(rawPath, offset, chunk, received)
        return publicMetadata(m, Math.max(received, offset + chunk.length))
      }
      if (op === 'attachment_finish') {
        if (received !== m.sizeBytes) fail('invalid_state')
        const raw = await bytes(rawPath, FILE_LIMIT)
        if (hash(raw) !== id) fail('digest_mismatch')
        if (m.status !== 'ready') {
          try {
            if (imageFile(m.name)) {
              const info = inspectPresentationImage(raw)
              const expected = extname(m.name).toLowerCase() === '.png' ? 'image/png' : 'image/jpeg'
              if (info.mime !== expected) fail('parse_failed')
              checkAbort(signal)
              const image = await normalizeImage(raw)
              checkAbort(signal)
              const normalized = inspectPresentationImage(image.bytes)
              if (
                image.bytes.length > PRESENTATION_IMAGE_CACHE_LIMIT ||
                normalized.mime !== 'image/png' ||
                image.width !== info.width ||
                image.height !== info.height ||
                normalized.width !== image.width ||
                normalized.height !== image.height
              )
                fail('parse_failed')
              await atomic(join(dir, 'image.png'), Buffer.from(image.bytes))
              checkAbort(signal)
              m = {
                attachmentId: id,
                sha256: id,
                name: m.name,
                sizeBytes: m.sizeBytes,
                status: 'ready',
                kind: 'image',
                mime: 'image/png',
                width: image.width,
                height: image.height,
                assetSha256: hash(image.bytes),
              }
            } else {
              if (extname(m.name).toLowerCase() === '.docx') checkDocx(raw)
              checkAbort(signal)
              const parsed = await parse(rawPath)
              checkAbort(signal)
              if (
                !parsed.ok ||
                parsed.kind !== 'text' ||
                typeof parsed.text !== 'string' ||
                parsed.text.length > TEXT_LIMIT
              )
                fail('parse_failed')
              await atomic(join(dir, 'text.txt'), parsed.text)
              m = {
                attachmentId: id,
                sha256: id,
                name: m.name,
                sizeBytes: m.sizeBytes,
                status: 'ready',
                kind: 'text',
                totalChars: parsed.text.length,
                textDigest: hash(parsed.text),
              }
            }
          } catch {
            checkAbort(signal)
            m = {
              attachmentId: id,
              sha256: id,
              name: m.name,
              sizeBytes: m.sizeBytes,
              status: 'failed',
              error: 'parse_failed',
            }
          }
          await atomic(join(dir, 'metadata.json'), JSON.stringify(m))
        }
        if (m.kind === 'image') await cachedImage(dir, m)
        checkAbort(signal)
        return publicMetadata(m, received)
      }
      if (op === 'attachment_asset') {
        const image = await cachedImage(dir, m)
        checkAbort(signal)
        return {
          id,
          mime: 'image/png',
          base64: image.toString('base64'),
          width: m.width,
          height: m.height,
          source: `attachment:${id}`,
        }
      }
      if (!integer(body.offset, 0, TEXT_LIMIT) || !integer(body.maxChars, 1, 24000))
        fail('invalid_request')
      if (m.status !== 'ready' || m.kind !== 'text') fail('invalid_state')
      const text = (await bytes(join(dir, 'text.txt'), TEXT_LIMIT * 4)).toString('utf8')
      if (text.length !== m.totalChars || hash(text) !== m.textDigest || body.offset > text.length)
        fail('invalid_state')
      checkAbort(signal)
      return {
        attachmentId: id,
        name: m.name,
        offset: body.offset,
        totalChars: text.length,
        text: text.slice(body.offset, body.offset + body.maxChars),
        sourceUri: `attachment:${id}`,
      }
    } catch (e) {
      const code = e instanceof Error ? e.message : ''
      if (
        [
          'invalid_request',
          'unsupported_file',
          'not_found',
          'invalid_state',
          'attachment_conflict',
          'quota_exceeded',
          'digest_mismatch',
          'parse_failed',
          'aborted',
        ].includes(code)
      )
        throw e
      return fail('invalid_state')
    } finally {
      release()
      if (locks.get(doc) === tail) locks.delete(doc)
    }
  }
}
