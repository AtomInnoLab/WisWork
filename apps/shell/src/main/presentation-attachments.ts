import { PresentationAcquisitionStore } from '@wiswork/project-store/presentation-acquisition-store'
import {
  presentationAcquisitionErrors,
  type PresentationAcquisitionRecord,
  type PresentationAcquisitionError,
} from '@wiswork/project-store/presentation-acquisition'
import { MAX_PRESENTATION_SOURCE_TEXT_CHARS } from '@wiswork/pptx-engine/presentation-source-limits'
import { createHash, randomUUID } from 'node:crypto'
import { constants } from 'node:fs'
import { lstat, mkdir, open, readdir, rename, rm } from 'node:fs/promises'
import { extname, join, resolve } from 'node:path'
import { inflateRawSync } from 'node:zlib'
import {
  decodeHtmlBytes,
  htmlToText,
  paragraphSections,
  parseFileToText,
} from '@wiswork/file-parse'
import { fetchRemoteImage, fetchWithSsrfGuard, isSafeRemoteUrl } from '@wiswork/electron-utils'
import {
  inspectPresentationImage,
  normalizePresentationImage,
  normalizePresentationImageFirstFrame,
  PRESENTATION_IMAGE_INPUT_LIMIT,
  PRESENTATION_IMAGE_CACHE_LIMIT,
} from './presentation-image'

const FILE_LIMIT = 50 * 1024 * 1024
const DOCUMENT_LIMIT = 100 * 1024 * 1024
const FILE_RESERVATION_FLOOR = 64 * 1024
const CHUNK_LIMIT = 128 * 1024
const TEXT_LIMIT = MAX_PRESENTATION_SOURCE_TEXT_CHARS
const WEBPAGE_TEXT_LIMIT = 1_000_000
const WEBPAGE_LIMIT = 5 * 1024 * 1024
const STAGING_MAX_AGE_MS = 24 * 60 * 60 * 1000
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
  return ['.png', '.jpg', '.jpeg', '.gif', '.webp'].includes(extname(name).toLowerCase())
}
function imageMime(name: string) {
  const ext = extname(name).toLowerCase()
  return ext === '.png'
    ? 'image/png'
    : ext === '.gif'
      ? 'image/gif'
      : ext === '.webp'
        ? 'image/webp'
        : 'image/jpeg'
}
function fileLimit(name: string) {
  return imageFile(name) ? PRESENTATION_IMAGE_INPUT_LIMIT : FILE_LIMIT
}
function supported(name: string) {
  return (
    imageFile(name) ||
    ['.pdf', '.docx', '.txt', '.md', '.csv', '.json', '.html', '.htm'].includes(
      extname(name).toLowerCase(),
    )
  )
}
function sourceValid(value: unknown): value is string {
  if (typeof value !== 'string' || value.length > 2048) return false
  try {
    const url = new URL(value)
    return (
      ['http:', 'https:'].includes(url.protocol) &&
      !url.username &&
      !url.password &&
      !url.search &&
      !url.hash
    )
  } catch {
    return false
  }
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
  animationHandling?: 'first_frame'
  source?: string
  sourceUrlHash?: string
  retrievedAt?: number
  sourceAliases?: { source: string; sourceUrlHash: string }[]
  licenseDeclaration?: {
    kind: 'owned' | 'licensed' | 'public_domain'
    evidenceAttachmentId: string
    assertedAt: number
  }
  error?: string
  totalChars?: number
  textDigest?: string
  sectionsDigest?: string
  sectionCount?: number
  pagesWithoutExtractedText?: number[]
  pagesWithSparseExtractedText?: number[]
}
async function directory(path: string, create = true) {
  if (create) {
    try {
      await mkdir(path, { mode: 0o700 })
    } catch (e) {
      if ((e as NodeJS.ErrnoException).code !== 'EEXIST') throw e
    }
  }
  const info = await lstat(path).catch((error: NodeJS.ErrnoException) => {
    if (!create && error.code === 'ENOENT') fail('not_found')
    throw error
  })
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
async function cleanupOldStaging(root: string) {
  const entries = await readdir(root, { withFileTypes: true })
  for (const entry of entries) {
    if (!/^\.tmp-[a-f0-9-]{36}$/.test(entry.name) || !entry.isDirectory()) continue
    const path = join(root, entry.name)
    const info = await lstat(path).catch(() => undefined)
    if (
      !info?.isDirectory() ||
      info.isSymbolicLink() ||
      Date.now() - info.mtimeMs < STAGING_MAX_AGE_MS
    )
      continue
    await rm(path, { recursive: true, force: true }).catch(() => undefined)
  }
}
async function metadata(dir: string, id: string): Promise<Metadata> {
  await directory(dir)
  let m: Metadata
  try {
    m = JSON.parse(
      (await bytes(join(dir, 'metadata.json'), 32 * 1024)).toString('utf8'),
    ) as Metadata
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
  if (
    (m.source !== undefined && !sourceValid(m.source)) ||
    (m.sourceUrlHash !== undefined && !isId(m.sourceUrlHash)) ||
    (m.source === undefined) !== (m.sourceUrlHash === undefined) ||
    (m.retrievedAt !== undefined &&
      (!integer(m.retrievedAt, 1, Number.MAX_SAFE_INTEGER) || m.source === undefined)) ||
    (m.sourceAliases !== undefined &&
      (!Array.isArray(m.sourceAliases) ||
        m.sourceAliases.length > 31 ||
        !m.source ||
        !m.sourceUrlHash ||
        new Set(m.sourceAliases.map((alias) => alias?.sourceUrlHash)).size !==
          m.sourceAliases.length ||
        m.sourceAliases.some(
          (alias) =>
            !alias ||
            !sourceValid(alias.source) ||
            !isId(alias.sourceUrlHash) ||
            alias.sourceUrlHash === m.sourceUrlHash,
        ))) ||
    (m.status === 'uploading' && m.source !== undefined) ||
    (m.status === 'failed' && m.source !== undefined && m.error !== 'animated_image_unsupported') ||
    (!imageFile(m.name) && extname(m.name).toLowerCase() !== '.html' && m.source !== undefined)
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
        (m.animationHandling !== undefined && m.animationHandling !== 'first_frame') ||
        (m.licenseDeclaration !== undefined &&
          (!m.licenseDeclaration ||
            typeof m.licenseDeclaration !== 'object' ||
            Array.isArray(m.licenseDeclaration) ||
            Object.keys(m.licenseDeclaration).sort().join(',') !==
              'assertedAt,evidenceAttachmentId,kind' ||
            !['owned', 'licensed', 'public_domain'].includes(m.licenseDeclaration.kind) ||
            !isId(m.licenseDeclaration.evidenceAttachmentId) ||
            m.licenseDeclaration.evidenceAttachmentId === id ||
            !integer(m.licenseDeclaration.assertedAt, 1, Number.MAX_SAFE_INTEGER))) ||
        m.totalChars !== undefined ||
        m.textDigest !== undefined
      )
        fail('invalid_state')
    } else if (
      m.kind !== 'text' ||
      !integer(m.totalChars, 0, TEXT_LIMIT) ||
      !isId(m.textDigest) ||
      m.licenseDeclaration !== undefined ||
      m.animationHandling !== undefined
    )
      fail('invalid_state')
  }
  if (
    (m.sectionsDigest !== undefined || m.sectionCount !== undefined) &&
    (m.status !== 'ready' ||
      m.kind !== 'text' ||
      !['.pdf', '.docx', '.html', '.htm'].includes(extname(m.name).toLowerCase()) ||
      !isId(m.sectionsDigest) ||
      !integer(m.sectionCount, 1, 4096))
  )
    fail('invalid_state')
  if (
    m.pagesWithoutExtractedText !== undefined &&
    (m.status !== 'ready' ||
      m.kind !== 'text' ||
      extname(m.name).toLowerCase() !== '.pdf' ||
      !m.sectionCount ||
      !Array.isArray(m.pagesWithoutExtractedText) ||
      m.pagesWithoutExtractedText.length < 1 ||
      m.pagesWithoutExtractedText.length >= m.sectionCount ||
      m.pagesWithoutExtractedText.some(
        (page, index) =>
          !integer(page, 1, m.sectionCount!) ||
          (index > 0 && page <= m.pagesWithoutExtractedText![index - 1]!),
      ))
  )
    fail('invalid_state')
  if (
    m.pagesWithSparseExtractedText !== undefined &&
    (m.status !== 'ready' ||
      m.kind !== 'text' ||
      extname(m.name).toLowerCase() !== '.pdf' ||
      !m.sectionCount ||
      !Array.isArray(m.pagesWithSparseExtractedText) ||
      m.pagesWithSparseExtractedText.length < 1 ||
      m.pagesWithSparseExtractedText.length > m.sectionCount ||
      m.pagesWithSparseExtractedText.some(
        (page, index) =>
          !integer(page, 1, m.sectionCount!) ||
          (index > 0 && page <= m.pagesWithSparseExtractedText![index - 1]!) ||
          m.pagesWithoutExtractedText?.includes(page),
      ))
  )
    fail('invalid_state')
  if (m.status !== 'ready' && m.licenseDeclaration !== undefined) fail('invalid_state')
  if (m.status !== 'ready' && m.animationHandling !== undefined) fail('invalid_state')
  if (
    m.status === 'failed' &&
    !['parse_failed', 'animated_image_unsupported'].includes(m.error ?? '')
  )
    fail('invalid_state')
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
  ...(m.sectionCount !== undefined ? { sectionCount: m.sectionCount } : {}),
  ...(m.pagesWithoutExtractedText
    ? { pagesWithoutExtractedText: m.pagesWithoutExtractedText }
    : {}),
  ...(m.pagesWithSparseExtractedText
    ? { pagesWithSparseExtractedText: m.pagesWithSparseExtractedText }
    : {}),
  ...(m.source ? { source: m.source } : {}),
  ...(m.sourceUrlHash ? { sourceUrlHash: m.sourceUrlHash } : {}),
  ...(m.retrievedAt ? { retrievedAt: m.retrievedAt } : {}),
  ...(m.sourceAliases?.length
    ? { sources: [m.source!, ...m.sourceAliases.map((alias) => alias.source)] }
    : {}),
  ...(m.kind === 'image'
    ? {
        mime: m.mime,
        width: m.width,
        height: m.height,
        assetSha256: m.assetSha256,
        ...(m.animationHandling ? { animationHandling: m.animationHandling } : {}),
        ...(m.licenseDeclaration ? { licenseDeclaration: m.licenseDeclaration } : {}),
      }
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

function sectionUnit(name: string): { unit: '页' | '段'; separator: number } {
  return extname(name).toLowerCase() === '.pdf'
    ? { unit: '页', separator: 2 }
    : { unit: '段', separator: 1 }
}

async function textSections(
  dir: string,
  m: Metadata,
  text: string,
): Promise<{ locator: string; start: number; end: number }[] | undefined> {
  if (!m.sectionsDigest) return undefined
  const raw = await bytes(join(dir, 'sections.json'), 512 * 1024)
  if (hash(raw) !== m.sectionsDigest) fail('invalid_state')
  let sections: unknown
  try {
    sections = JSON.parse(raw.toString('utf8'))
  } catch {
    fail('invalid_state')
  }
  if (!Array.isArray(sections) || sections.length !== m.sectionCount) fail('invalid_state')
  const { unit, separator } = sectionUnit(m.name)
  let previous = -separator
  for (const [index, section] of sections.entries()) {
    if (
      !section ||
      typeof section !== 'object' ||
      Array.isArray(section) ||
      Object.keys(section).sort().join(',') !== 'end,locator,start' ||
      section.locator !== `第 ${index + 1} ${unit}` ||
      !integer(section.start, 0, text.length) ||
      !integer(section.end, section.start, text.length) ||
      section.start !== previous + separator
    )
      fail('invalid_state')
    previous = section.end
  }
  if (previous !== text.length) fail('invalid_state')
  return sections as { locator: string; start: number; end: number }[]
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
  normalizeFirstFrame?: typeof normalizePresentationImageFirstFrame
  fetchImage?: (url: string, signal: AbortSignal) => Promise<Response | null>
  fetchPage?: (url: string, signal: AbortSignal) => Promise<Response | null>
}) {
  const history = new PresentationAcquisitionStore(options.userDataPath)
  const root = join(resolve(options.userDataPath), 'presentation-attachments')
  const parse = options.parse ?? parseFileToText
  const normalizeImage = options.normalizeImage ?? normalizePresentationImage
  const normalizeFirstFrame = options.normalizeFirstFrame ?? normalizePresentationImageFirstFrame
  const fetchImage =
    options.fetchImage ??
    ((url: string, signal: AbortSignal) =>
      fetchRemoteImage(url, {
        fetchImpl: (input, init) => fetch(input, { ...init, signal }),
      }))
  const fetchPage =
    options.fetchPage ??
    ((url: string, signal: AbortSignal) =>
      fetchWithSsrfGuard(url, {
        headers: { Accept: 'text/html,application/xhtml+xml;q=0.9' },
        fetchImpl: (input, init) => fetch(input, { ...init, signal }),
      }))
  let stagingCleanup: Promise<void> | undefined
  return async (body: Record<string, unknown>, signal: AbortSignal): Promise<unknown> => {
    const fields: Record<string, string[]> = {
      attachment_begin: ['attachmentId', 'name', 'sizeBytes', 'sha256'],
      attachment_chunk: ['attachmentId', 'offset', 'base64'],
      attachment_finish: ['attachmentId'],
      attachment_extract_first_frame: ['attachmentId'],
      attachment_delete: ['attachmentId'],
      attachment_import_url: ['url'],
      attachment_import_webpage: ['url'],
      attachment_attest_license: ['attachmentId', 'license', 'evidenceAttachmentId'],
      attachment_revoke_license: ['attachmentId'],
      attachment_acquisition_history: [],
      attachment_list: [],
      attachment_list_assets: [],
      attachment_metadata: ['attachmentId'],
      attachment_match_excerpt: ['attachmentId', 'excerpt'],
      attachment_asset: ['attachmentId'],
      attachment_original: ['attachmentId', 'offset', 'length'],
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
    const allowed = [
      'operation',
      'documentId',
      ...fields[op]!,
      ...(op === 'attachment_list_assets' ? ['after'] : []),
      ...(op === 'attachment_import_url' ? ['stageAnimated'] : []),
      ...(op === 'attachment_match_excerpt' ? ['locator'] : []),
    ]
    if (
      Object.keys(body).some((k) => !allowed.includes(k)) ||
      allowed.some(
        (k) => !['after', 'stageAnimated', 'locator'].includes(k) && !Object.hasOwn(body, k),
      )
    )
      fail('invalid_request')
    if (
      ![
        'attachment_list',
        'attachment_acquisition_history',
        'attachment_list_assets',
        'attachment_import_url',
        'attachment_import_webpage',
      ].includes(op) &&
      !isId(body.attachmentId)
    )
      fail('invalid_request')
    if (
      (op === 'attachment_import_url' || op === 'attachment_import_webpage') &&
      (typeof body.url !== 'string' || body.url.length > 2048)
    )
      fail('invalid_request')
    if (body.stageAnimated !== undefined && body.stageAnimated !== true) fail('invalid_request')
    if (body.after !== undefined && !isId(body.after)) fail('invalid_request')
    if (
      op === 'attachment_match_excerpt' &&
      (typeof body.excerpt !== 'string' || body.excerpt.length > 12000)
    )
      fail('invalid_request')
    if (
      body.locator !== undefined &&
      (typeof body.locator !== 'string' || !/^第 [1-9]\d{0,5} (页|段)$/.test(body.locator))
    )
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
    if (op === 'attachment_acquisition_history') {
      checkAbort(signal)
      const result = await history.read(body.documentId)
      checkAbort(signal)
      return result
    }
    const remote = op === 'attachment_import_url' || op === 'attachment_import_webpage'
    let source: URL | undefined
    if (remote) {
      try {
        source = new URL(body.url as string)
      } catch {
        fail('invalid_request')
      }
      if (!['http:', 'https:'].includes(source.protocol) || source.username || source.password)
        fail('invalid_request')
    }
    const doc = join(root, hash(body.documentId))
    const previous = locks.get(doc) ?? Promise.resolve()
    let release!: () => void
    const tail = new Promise<void>((r) => {
      release = r
    })
    locks.set(doc, tail)
    await previous
    let acquisition: PresentationAcquisitionRecord | undefined
    const complete = async (item: Metadata, size: number) => {
      if (acquisition) {
        const result =
          item.status === 'ready'
            ? {
                state: 'ready' as const,
                attachmentId: item.attachmentId,
                sha256: item.sha256,
                sizeBytes: item.sizeBytes,
                ...(acquisition.kind === 'image' ? { assetSha256: item.assetSha256 } : {}),
              }
            : {
                state: 'rejected' as const,
                error: 'animated_image_unsupported' as const,
                attachmentId: item.attachmentId,
              }
        // The published attachment remains valid if the independent history write fails.
        await history
          .finish(body.documentId as string, acquisition.id, result)
          .catch(() => undefined)
      }
      return publicMetadata(item, size)
    }
    try {
      if (source) {
        const sourceUrlHash = hash(source.toString())
        source.search = ''
        source.hash = ''
        acquisition = await history.begin(body.documentId, {
          kind: op === 'attachment_import_url' ? 'image' : 'webpage',
          source: source.toString(),
          sourceUrlHash,
        })
      }
      checkAbort(signal)
      await directory(
        root,
        ![
          'attachment_read',
          'attachment_original',
          'attachment_metadata',
          'attachment_match_excerpt',
        ].includes(op),
      )
      if (!stagingCleanup) stagingCleanup = cleanupOldStaging(root)
      await stagingCleanup
      await directory(
        doc,
        ![
          'attachment_read',
          'attachment_original',
          'attachment_metadata',
          'attachment_match_excerpt',
        ].includes(op),
      )
      const entries = await readdir(doc)
      if (entries.some((e) => !isId(e))) fail('invalid_state')
      const id = body.attachmentId as string
      const dir = id ? join(doc, id) : doc
      const exists = entries.includes(id)
      if (op === 'attachment_import_webpage') {
        let url: URL
        try {
          url = new URL(body.url as string)
        } catch {
          fail('invalid_request')
        }
        if (!['http:', 'https:'].includes(url.protocol) || url.username || url.password)
          fail('invalid_request')
        const urlHash = hash(url.toString())
        for (const entry of entries) {
          const item = await metadata(join(doc, entry), entry)
          if (item.status !== 'ready' || item.kind !== 'text' || item.sourceUrlHash !== urlHash)
            continue
          const stored = await bytes(join(doc, entry, 'raw.html'), WEBPAGE_LIMIT)
          if (hash(stored) !== entry) fail('digest_mismatch')
          return await complete(item, item.sizeBytes)
        }
        if (!(await isSafeRemoteUrl(url.toString()))) fail('remote_webpage_unavailable')
        const combined = AbortSignal.any([signal, AbortSignal.timeout(15_000)])
        let response: Response | null
        try {
          response = await fetchPage(url.toString(), combined)
        } catch {
          checkAbort(combined)
          fail('remote_webpage_unavailable')
        }
        if (
          !response?.ok ||
          !response.body ||
          !/^text\/html(?:\s*;|$)/i.test(response.headers.get('content-type') ?? '')
        )
          fail('remote_webpage_unavailable')
        if (Number(response.headers.get('content-length')) > WEBPAGE_LIMIT) fail('quota_exceeded')
        const reader = response.body.getReader()
        const chunks: Uint8Array[] = []
        let total = 0
        try {
          while (true) {
            checkAbort(combined)
            const next = await reader.read().catch(() => {
              checkAbort(combined)
              fail('remote_webpage_unavailable')
            })
            if (next.done) break
            total += next.value.byteLength
            if (total > WEBPAGE_LIMIT) fail('quota_exceeded')
            chunks.push(next.value)
          }
        } finally {
          await reader.cancel().catch(() => undefined)
        }
        checkAbort(combined)
        const raw = Buffer.concat(
          chunks.map((chunk) => Buffer.from(chunk)),
          total,
        )
        let text: string
        try {
          text = htmlToText(decodeHtmlBytes(raw, response.headers.get('content-type') ?? undefined))
        } catch {
          fail('parse_failed')
        }
        if (!text || text.length > WEBPAGE_TEXT_LIMIT) fail('parse_failed')
        const sections = paragraphSections(text)
        if (sections.length > 4096) fail('parse_failed')
        const sectionsRaw = JSON.stringify(sections)
        if (Buffer.byteLength(sectionsRaw) > 512 * 1024) fail('parse_failed')
        const attachmentId = hash(raw)
        if (entries.includes(attachmentId)) fail('remote_webpage_source_conflict')
        let declared = 0
        for (const entry of entries)
          declared += Math.max(
            (await metadata(join(doc, entry), entry)).sizeBytes,
            FILE_RESERVATION_FLOOR,
          )
        if (
          entries.length >= 32 ||
          declared + Math.max(raw.length, FILE_RESERVATION_FLOOR) > DOCUMENT_LIMIT
        )
          fail('quota_exceeded')
        url.search = ''
        url.hash = ''
        const item: Metadata = {
          attachmentId,
          sha256: attachmentId,
          name: 'remote.html',
          sizeBytes: raw.length,
          status: 'ready',
          kind: 'text',
          totalChars: text.length,
          textDigest: hash(text),
          sectionCount: sections.length,
          sectionsDigest: hash(sectionsRaw),
          source: url.toString(),
          sourceUrlHash: urlHash,
          retrievedAt: Date.now(),
        }
        const staging = join(root, `.tmp-${randomUUID()}`)
        await directory(staging)
        try {
          await atomic(join(staging, 'raw.html'), raw)
          await atomic(join(staging, 'text.txt'), text)
          await atomic(join(staging, 'sections.json'), sectionsRaw)
          await atomic(join(staging, 'metadata.json'), JSON.stringify(item))
          await rename(staging, join(doc, attachmentId))
        } finally {
          await rm(staging, { recursive: true, force: true })
        }
        return await complete(item, raw.length)
      }
      if (op === 'attachment_import_url') {
        let url: URL
        try {
          url = new URL(body.url as string)
        } catch {
          fail('invalid_request')
        }
        if (!['http:', 'https:'].includes(url.protocol) || url.username || url.password)
          fail('invalid_request')
        const urlHash = hash(url.toString())
        for (const entry of entries) {
          const item = await metadata(join(doc, entry), entry)
          if (
            (item.status !== 'ready' &&
              !(item.status === 'failed' && item.error === 'animated_image_unsupported')) ||
            (item.sourceUrlHash !== urlHash &&
              !item.sourceAliases?.some((alias) => alias.sourceUrlHash === urlHash))
          )
            continue
          if (item.status === 'failed' && body.stageAnimated !== true)
            fail('animated_image_unsupported')
          if (item.status === 'ready') await cachedImage(join(doc, entry), item)
          else {
            const original = await bytes(
              join(doc, entry, `raw${extname(item.name).toLowerCase()}`),
              PRESENTATION_IMAGE_INPUT_LIMIT,
            )
            if (original.length !== item.sizeBytes || hash(original) !== entry)
              fail('digest_mismatch')
          }
          return await complete(item, item.sizeBytes)
        }
        if (!(await isSafeRemoteUrl(url.toString()))) fail('remote_image_unavailable')
        const timeout = AbortSignal.timeout(15_000)
        const combined = AbortSignal.any([signal, timeout])
        const checkDownloadAbort = () => {
          if (signal.aborted) fail('aborted')
          if (timeout.aborted) fail('remote_image_unavailable')
        }
        let response: Response | null
        try {
          response = await fetchImage(url.toString(), combined)
        } catch {
          checkDownloadAbort()
          fail('remote_image_unavailable')
        }
        checkDownloadAbort()
        if (!response?.ok || !response.body) fail('remote_image_unavailable')
        if (Number(response.headers.get('content-length')) > PRESENTATION_IMAGE_INPUT_LIMIT)
          fail('quota_exceeded')
        const reader = response.body.getReader()
        const chunks: Uint8Array[] = []
        let total = 0
        try {
          while (true) {
            checkDownloadAbort()
            let next: ReadableStreamReadResult<Uint8Array>
            try {
              next = await reader.read()
            } catch {
              checkDownloadAbort()
              fail('remote_image_unavailable')
            }
            if (next.done) break
            total += next.value.byteLength
            if (total > PRESENTATION_IMAGE_INPUT_LIMIT) fail('quota_exceeded')
            chunks.push(next.value)
          }
        } finally {
          await reader.cancel().catch(() => undefined)
        }
        checkDownloadAbort()
        const raw = Buffer.concat(
          chunks.map((chunk) => Buffer.from(chunk)),
          total,
        )
        let info: ReturnType<typeof inspectPresentationImage>
        let animatedInput = false
        try {
          info = inspectPresentationImage(raw)
        } catch (error) {
          if (!(error instanceof Error) || error.message !== 'animated_image_unsupported')
            throw error
          info = inspectPresentationImage(raw, true)
          if (!info.animated) fail('parse_failed')
          if (body.stageAnimated !== true) fail('animated_image_unsupported')
          animatedInput = true
        }
        const image = animatedInput ? undefined : await normalizeImage(raw)
        checkDownloadAbort()
        if (image) {
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
        }
        const attachmentId = hash(raw)
        if (entries.includes(attachmentId)) {
          const existingDir = join(doc, attachmentId)
          const existing = await metadata(existingDir, attachmentId)
          if (
            (existing.status !== 'ready' &&
              !(
                animatedInput &&
                existing.status === 'failed' &&
                existing.error === 'animated_image_unsupported'
              )) ||
            (existing.status === 'ready' && existing.kind !== 'image') ||
            !existing.source ||
            !existing.sourceUrlHash ||
            !(
              await bytes(
                join(existingDir, `raw${extname(existing.name)}`),
                PRESENTATION_IMAGE_INPUT_LIMIT,
              )
            ).equals(raw)
          )
            fail('remote_image_source_conflict')
          if (existing.status === 'ready') await cachedImage(existingDir, existing)
          if ((existing.sourceAliases?.length ?? 0) >= 31) fail('quota_exceeded')
          url.search = ''
          url.hash = ''
          const updated = {
            ...existing,
            sourceAliases: [
              ...(existing.sourceAliases ?? []),
              { source: url.toString(), sourceUrlHash: urlHash },
            ],
          }
          await atomic(join(existingDir, 'metadata.json'), JSON.stringify(updated))
          return await complete(updated, updated.sizeBytes)
        }
        let declared = 0
        for (const entry of entries)
          declared += Math.max(
            (await metadata(join(doc, entry), entry)).sizeBytes,
            FILE_RESERVATION_FLOOR,
          )
        if (declared + Math.max(raw.length, FILE_RESERVATION_FLOOR) > DOCUMENT_LIMIT)
          fail('quota_exceeded')
        const name = `remote-${attachmentId}.${{ 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp' }[info.mime]}`
        url.search = ''
        url.hash = ''
        const item: Metadata = {
          attachmentId,
          sha256: attachmentId,
          name,
          sizeBytes: raw.length,
          ...(image
            ? {
                status: 'ready' as const,
                kind: 'image' as const,
                mime: 'image/png' as const,
                width: image.width,
                height: image.height,
                assetSha256: hash(image.bytes),
              }
            : { status: 'failed' as const, error: 'animated_image_unsupported' }),
          source: url.toString(),
          sourceUrlHash: urlHash,
        }
        const staging = join(root, `.tmp-${randomUUID()}`)
        await directory(staging)
        try {
          await atomic(join(staging, `raw${extname(name)}`), raw)
          if (image) await atomic(join(staging, 'image.png'), Buffer.from(image.bytes))
          await atomic(join(staging, 'metadata.json'), JSON.stringify(item))
          await rename(staging, join(doc, attachmentId))
        } finally {
          await rm(staging, { recursive: true, force: true })
        }
        return await complete(item, raw.length)
      }
      if (op === 'attachment_list' || op === 'attachment_list_assets') {
        const attachments = []
        const sorted = entries.sort()
        const page =
          op === 'attachment_list_assets'
            ? sorted
                .filter((entry) => entry > ((body.after as string | undefined) ?? ''))
                .slice(0, 32)
            : sorted
        for (const entry of page) {
          const m = await metadata(join(doc, entry), entry)
          if (op === 'attachment_list' && imageFile(m.name)) continue
          const received = await rawSize(join(doc, entry, `raw${extname(m.name).toLowerCase()}`))
          if (received > m.sizeBytes) fail('invalid_state')
          attachments.push(publicMetadata(m, received))
        }
        checkAbort(signal)
        return op === 'attachment_list_assets' &&
          page.length > 0 &&
          sorted.some((entry) => entry > page[page.length - 1]!)
          ? { attachments, nextAfter: page[page.length - 1] }
          : { attachments }
      }
      if (!exists && op !== 'attachment_begin') fail('not_found')
      if (op === 'attachment_delete') {
        await metadata(dir, id)
        for (const entry of entries) {
          if (entry === id) continue
          const other = await metadata(join(doc, entry), entry)
          if (other.licenseDeclaration?.evidenceAttachmentId === id) fail('attachment_in_use')
        }
        checkAbort(signal)
        await rm(dir, { recursive: true })
        return { attachmentId: id, deleted: true }
      }
      if (op === 'attachment_begin' && !exists) {
        let declared = 0,
          sourceCount = 0
        for (const entry of entries) {
          const item = await metadata(join(doc, entry), entry)
          declared += Math.max(item.sizeBytes, FILE_RESERVATION_FLOOR)
          if (!imageFile(item.name)) sourceCount++
        }
        if (
          (!imageFile(body.name as string) && sourceCount >= 32) ||
          declared + Math.max(body.sizeBytes as number, FILE_RESERVATION_FLOOR) > DOCUMENT_LIMIT
        )
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
      if (op === 'attachment_extract_first_frame') {
        if (m.status === 'ready' && m.animationHandling === 'first_frame') {
          await cachedImage(dir, m)
          return publicMetadata(m, m.sizeBytes)
        }
        if (m.status !== 'failed' || m.error !== 'animated_image_unsupported' || !imageFile(m.name))
          fail('invalid_state')
        const raw = await bytes(
          join(dir, `raw${extname(m.name).toLowerCase()}`),
          PRESENTATION_IMAGE_INPUT_LIMIT,
        )
        if (raw.length !== m.sizeBytes || hash(raw) !== id) fail('digest_mismatch')
        const info = inspectPresentationImage(raw, true)
        if (!info.animated || info.mime !== imageMime(m.name)) fail('invalid_state')
        checkAbort(signal)
        const image = await normalizeFirstFrame(raw)
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
          animationHandling: 'first_frame',
          ...(m.source
            ? {
                source: m.source,
                sourceUrlHash: m.sourceUrlHash,
                ...(m.sourceAliases ? { sourceAliases: m.sourceAliases } : {}),
              }
            : {}),
        }
        await atomic(join(dir, 'metadata.json'), JSON.stringify(m))
        return publicMetadata(m, m.sizeBytes)
      }
      if (op === 'attachment_attest_license' || op === 'attachment_revoke_license') {
        if (m.status !== 'ready' || m.kind !== 'image') fail('invalid_state')
        if (op === 'attachment_attest_license') {
          const evidenceId = body.evidenceAttachmentId
          if (
            !['owned', 'licensed', 'public_domain'].includes(body.license as string) ||
            !isId(evidenceId) ||
            evidenceId === id ||
            !entries.includes(evidenceId)
          )
            fail('invalid_request')
          const evidenceDir = join(doc, evidenceId)
          const evidence = await metadata(evidenceDir, evidenceId)
          if (evidence.status !== 'ready' || evidence.kind !== 'text') fail('invalid_state')
          const value = (await bytes(join(evidenceDir, 'text.txt'), TEXT_LIMIT * 4)).toString(
            'utf8',
          )
          if (value.length !== evidence.totalChars || hash(value) !== evidence.textDigest)
            fail('invalid_state')
          m = {
            ...m,
            licenseDeclaration: {
              kind: body.license as 'owned' | 'licensed' | 'public_domain',
              evidenceAttachmentId: evidenceId,
              assertedAt: Date.now(),
            },
          }
        } else {
          const { licenseDeclaration: _declaration, ...rest } = m
          m = rest
        }
        checkAbort(signal)
        await atomic(join(dir, 'metadata.json'), JSON.stringify(m))
        return publicMetadata(m, m.sizeBytes)
      }
      const rawPath = join(dir, `raw${extname(m.name).toLowerCase()}`)
      const received = await rawSize(rawPath)
      if (received > m.sizeBytes) fail('invalid_state')
      if (op === 'attachment_metadata') return publicMetadata(m, received)
      if (op === 'attachment_match_excerpt') {
        if (m.status !== 'ready') return { attachmentId: id, status: 'not_ready' }
        if (m.kind !== 'text') return { attachmentId: id, status: 'unsupported' }
        const value = (await bytes(join(dir, 'text.txt'), TEXT_LIMIT * 4)).toString('utf8')
        if (value.length !== m.totalChars || hash(value) !== m.textDigest) fail('invalid_state')
        checkAbort(signal)
        if (!(body.excerpt as string).trim()) return { attachmentId: id, status: 'empty_excerpt' }
        const sections = await textSections(dir, m, value)
        const excerpt = body.excerpt as string
        const preferred = sections?.find((item) => item.locator === body.locator)
        const matchedSection = (section: { start: number; end: number }) => {
          const found = value.indexOf(excerpt, section.start)
          return found >= 0 && found + excerpt.length <= section.end ? found : -1
        }
        const preferredOffset = preferred ? matchedSection(preferred) : -1
        const section =
          preferredOffset >= 0 ? preferred : sections?.find((item) => matchedSection(item) >= 0)
        const offset =
          preferredOffset >= 0
            ? preferredOffset
            : section
              ? matchedSection(section)
              : sections
                ? -1
                : value.indexOf(excerpt)
        return offset < 0 || (sections && !section)
          ? { attachmentId: id, status: 'not_found' }
          : {
              attachmentId: id,
              status: 'found',
              offset,
              ...(section ? { locator: section.locator } : {}),
            }
      }
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
              const expected = imageMime(m.name)
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
                !parsed.text.trim() ||
                parsed.text.length > TEXT_LIMIT
              )
                fail('parse_failed')
              let sectionsRaw: string | undefined
              if (parsed.sections) {
                if (
                  !['.pdf', '.docx', '.html', '.htm'].includes(extname(m.name).toLowerCase()) ||
                  parsed.sections.length < 1 ||
                  parsed.sections.length > 4096
                )
                  fail('parse_failed')
                sectionsRaw = JSON.stringify(parsed.sections)
                if (Buffer.byteLength(sectionsRaw) > 512 * 1024) fail('parse_failed')
                const { unit, separator } = sectionUnit(m.name)
                let previous = -separator
                for (const [index, section] of parsed.sections.entries()) {
                  if (
                    section.locator !== `第 ${index + 1} ${unit}` ||
                    !integer(section.start, 0, parsed.text.length) ||
                    !integer(section.end, section.start, parsed.text.length) ||
                    section.start !== previous + separator
                  )
                    fail('parse_failed')
                  previous = section.end
                }
                if (previous !== parsed.text.length) fail('parse_failed')
                await atomic(join(dir, 'sections.json'), sectionsRaw)
              }
              await atomic(join(dir, 'text.txt'), parsed.text)
              const pagesWithoutExtractedText =
                extname(m.name).toLowerCase() === '.pdf'
                  ? (parsed.sections ?? []).flatMap((section, index) =>
                      section.start === section.end ? [index + 1] : [],
                    )
                  : []
              const pagesWithSparseExtractedText =
                extname(m.name).toLowerCase() === '.pdf'
                  ? (parsed.sections ?? []).flatMap((section, index) => {
                      // A coverage hint only: short text does not establish OCR accuracy or a defect.
                      const chars = parsed.text!.slice(section.start, section.end).trim().length
                      return chars > 0 && chars < 200 ? [index + 1] : []
                    })
                  : []
              m = {
                attachmentId: id,
                sha256: id,
                name: m.name,
                sizeBytes: m.sizeBytes,
                status: 'ready',
                kind: 'text',
                totalChars: parsed.text.length,
                textDigest: hash(parsed.text),
                ...(sectionsRaw
                  ? { sectionCount: parsed.sections!.length, sectionsDigest: hash(sectionsRaw) }
                  : {}),
                ...(pagesWithoutExtractedText.length ? { pagesWithoutExtractedText } : {}),
                ...(pagesWithSparseExtractedText.length ? { pagesWithSparseExtractedText } : {}),
              }
            }
          } catch (error) {
            checkAbort(signal)
            m = {
              attachmentId: id,
              sha256: id,
              name: m.name,
              sizeBytes: m.sizeBytes,
              status: 'failed',
              error:
                error instanceof Error && error.message === 'animated_image_unsupported'
                  ? 'animated_image_unsupported'
                  : 'parse_failed',
            }
          }
          await atomic(join(dir, 'metadata.json'), JSON.stringify(m))
        }
        if (m.kind === 'image') await cachedImage(dir, m)
        checkAbort(signal)
        return publicMetadata(m, received)
      }
      if (op === 'attachment_original') {
        if (!integer(body.offset, 0, m.sizeBytes) || !integer(body.length, 1, CHUNK_LIMIT))
          fail('invalid_request')
        if (
          m.status !== 'ready' ||
          m.kind !== 'image' ||
          m.animationHandling !== undefined ||
          m.sizeBytes > 2 * 1024 * 1024
        )
          fail('invalid_state')
        const raw = await bytes(rawPath, 2 * 1024 * 1024)
        if (raw.length !== m.sizeBytes || hash(raw) !== id) fail('digest_mismatch')
        const info = inspectPresentationImage(raw)
        const expected = imageMime(m.name)
        if (info.mime !== expected || info.width !== m.width || info.height !== m.height)
          fail('invalid_state')
        checkAbort(signal)
        return {
          attachmentId: id,
          offset: body.offset,
          sizeBytes: raw.length,
          sha256: id,
          mime: info.mime,
          base64: raw.subarray(body.offset, body.offset + body.length).toString('base64'),
        }
      }
      if (op === 'attachment_asset') {
        const image = await cachedImage(dir, m)
        if (m.licenseDeclaration) {
          const evidenceId = m.licenseDeclaration.evidenceAttachmentId
          const evidenceDir = join(doc, evidenceId)
          const evidence = await metadata(evidenceDir, evidenceId)
          if (evidence.status !== 'ready' || evidence.kind !== 'text') fail('invalid_state')
          const value = (await bytes(join(evidenceDir, 'text.txt'), TEXT_LIMIT * 4)).toString(
            'utf8',
          )
          if (value.length !== evidence.totalChars || hash(value) !== evidence.textDigest)
            fail('invalid_state')
        }
        checkAbort(signal)
        return {
          id,
          mime: 'image/png',
          base64: image.toString('base64'),
          width: m.width,
          height: m.height,
          source: m.source ?? `attachment:${id}`,
          ...(m.licenseDeclaration
            ? {
                license: m.licenseDeclaration.kind,
                licenseEvidence: `attachment:${m.licenseDeclaration.evidenceAttachmentId}`,
              }
            : {}),
          ...(m.sourceAliases?.length
            ? { sources: [m.source!, ...m.sourceAliases.map((alias) => alias.source)] }
            : {}),
        }
      }
      if (!integer(body.offset, 0, TEXT_LIMIT) || !integer(body.maxChars, 1, 24000))
        fail('invalid_request')
      if (m.status !== 'ready' || m.kind !== 'text') fail('invalid_state')
      const text = (await bytes(join(dir, 'text.txt'), TEXT_LIMIT * 4)).toString('utf8')
      if (text.length !== m.totalChars || hash(text) !== m.textDigest || body.offset > text.length)
        fail('invalid_state')
      const sections = await textSections(dir, m, text)
      const offset = body.offset as number
      const end = Math.min(text.length, offset + (body.maxChars as number))
      checkAbort(signal)
      return {
        attachmentId: id,
        name: m.name,
        offset: body.offset,
        totalChars: text.length,
        text: text.slice(body.offset, body.offset + body.maxChars),
        sourceUri: `attachment:${id}`,
        ...(sections
          ? { pageSpans: sections.filter((section) => section.end > offset && section.start < end) }
          : {}),
      }
    } catch (e) {
      const code = e instanceof Error ? e.message : ''
      if (acquisition)
        await history
          .finish(body.documentId as string, acquisition.id, {
            state: 'rejected',
            error: presentationAcquisitionErrors.includes(code as PresentationAcquisitionError)
              ? (code as PresentationAcquisitionError)
              : 'acquisition_failed',
          })
          .catch(() => undefined)
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
          'animated_image_unsupported',
          'remote_image_unavailable',
          'remote_image_source_conflict',
          'remote_webpage_unavailable',
          'remote_webpage_source_conflict',
          'attachment_in_use',
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
