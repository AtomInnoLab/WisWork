import JSZip from 'jszip'
import { parseSkillPackage, type ParsedSkill } from './skill-registry.js'

export const SKILL_PACKAGE_LIMITS = Object.freeze({
  maxCompressedBytes: 2 * 1024 * 1024,
  maxUncompressedBytes: 8 * 1024 * 1024,
  maxFileBytes: 2 * 1024 * 1024,
  maxEntries: 64,
  maxPathBytes: 256,
})

export interface SkillPackageFile {
  readonly path: string
  readonly bytes: Uint8Array
}

export interface ParsedSkillArchive {
  readonly skill: ParsedSkill
  readonly files: readonly SkillPackageFile[]
}

interface ZipStreamEntry extends JSZip.JSZipObject {
  internalStream(type: 'uint8array'): JSZip.JSZipStreamHelper<Uint8Array>
}

const encoder = new TextEncoder()
const utf8 = new TextDecoder('utf-8', { fatal: true })
const TEXT_EXTENSIONS = new Set(['md', 'txt', 'json', 'yaml', 'yml', 'csv'])
const IMAGE_EXTENSIONS = new Set(['png', 'jpg', 'jpeg', 'webp'])

function invalid(): never {
  throw new Error('invalid_skill_package')
}

function limited(): never {
  throw new Error('skill_package_limit')
}

function checkCancelled(signal?: AbortSignal): void {
  if (signal?.aborted) throw new Error('upload_cancelled')
}

function validatePath(rawPath: string): string {
  const path = rawPath.endsWith('/') ? rawPath.slice(0, -1) : rawPath
  if (!path || path.startsWith('/') || path.includes('\\') || path.includes('\0')) invalid()
  const normalized = path.normalize('NFC')
  if (normalized !== path || encoder.encode(rawPath).byteLength > SKILL_PACKAGE_LIMITS.maxPathBytes)
    invalid()
  const parts = path.split('/')
  if (parts.some((part) => !part || part === '.' || part === '..')) invalid()
  return path
}

function declaredEntryCount(source: Uint8Array): number {
  const view = new DataView(source.buffer, source.byteOffset, source.byteLength)
  for (
    let offset = source.byteLength - 22;
    offset >= Math.max(0, source.byteLength - 65_557);
    offset -= 1
  ) {
    if (view.getUint32(offset, true) === 0x06054b50) {
      if (offset + 22 + view.getUint16(offset + 20, true) !== source.byteLength) invalid()
      const diskEntries = view.getUint16(offset + 8, true)
      const entries = view.getUint16(offset + 10, true)
      if (diskEntries !== entries || entries === 0xffff) invalid()
      return entries
    }
  }
  invalid()
}

function validateMode(mode: number | string | null | undefined, directory: boolean): void {
  const numeric = typeof mode === 'string' ? Number.parseInt(mode, 8) : mode
  if (typeof numeric !== 'number' || !Number.isFinite(numeric)) return
  const type = numeric & 0o170000
  if (
    (type !== 0 && type !== 0o100000 && type !== 0o040000) ||
    (!directory && (numeric & 0o111) !== 0)
  )
    invalid()
}

function validateImage(extension: string, bytes: Uint8Array): void {
  const starts = (...values: number[]) => values.every((value, index) => bytes[index] === value)
  if (extension === 'png' && starts(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a)) return
  if ((extension === 'jpg' || extension === 'jpeg') && starts(0xff, 0xd8, 0xff)) return
  if (
    extension === 'webp' &&
    starts(0x52, 0x49, 0x46, 0x46) &&
    bytes[8] === 0x57 &&
    bytes[9] === 0x45 &&
    bytes[10] === 0x42 &&
    bytes[11] === 0x50
  )
    return
  invalid()
}

async function inflateBounded(
  entry: JSZip.JSZipObject,
  currentTotal: number,
  signal?: AbortSignal,
): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const stream = (entry as ZipStreamEntry).internalStream('uint8array')
    const chunks: Uint8Array[] = []
    let size = 0
    let settled = false
    const failOnce = (code: string) => {
      if (settled) return
      settled = true
      stream.pause()
      reject(new Error(code))
    }
    const cancel = () => failOnce('upload_cancelled')
    signal?.addEventListener('abort', cancel, { once: true })
    stream.on('data', (chunk) => {
      if (settled) return
      if (
        size + chunk.byteLength > SKILL_PACKAGE_LIMITS.maxFileBytes ||
        currentTotal + size + chunk.byteLength > SKILL_PACKAGE_LIMITS.maxUncompressedBytes
      )
        return failOnce('skill_package_limit')
      chunks.push(chunk.slice())
      size += chunk.byteLength
    })
    stream.on('error', () => failOnce('invalid_skill_package'))
    stream.on('end', () => {
      signal?.removeEventListener('abort', cancel)
      if (settled) return
      settled = true
      const output = new Uint8Array(size)
      let offset = 0
      for (const chunk of chunks) {
        output.set(chunk, offset)
        offset += chunk.byteLength
      }
      resolve(output)
    })
    stream.resume()
  })
}

export async function parseSkillArchive(
  source: Uint8Array,
  signal?: AbortSignal,
): Promise<ParsedSkillArchive> {
  checkCancelled(signal)
  if (!(source instanceof Uint8Array)) invalid()
  if (source.byteLength > SKILL_PACKAGE_LIMITS.maxCompressedBytes) limited()
  const declaredEntries = declaredEntryCount(source)
  if (declaredEntries > SKILL_PACKAGE_LIMITS.maxEntries) limited()
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(source, { createFolders: false })
  } catch {
    invalid()
  }
  checkCancelled(signal)
  const allEntries = Object.values(zip.files)
  if (allEntries.length !== declaredEntries) invalid()
  const entries = allEntries.filter((entry) => !entry.dir)
  const names = new Set<string>()
  let manifestCount = 0
  let declaredTotal = 0
  for (const entry of allEntries) {
    const unsafeName = (entry as typeof entry & { unsafeOriginalName?: string }).unsafeOriginalName
    if (unsafeName && unsafeName !== entry.name) invalid()
    const path = validatePath(entry.name)
    const collisionKey = path.normalize('NFC').toLocaleLowerCase('en-US')
    if (names.has(collisionKey)) invalid()
    names.add(collisionKey)
    if (!entry.dir && path.toLocaleLowerCase('en-US') === 'skill.md') manifestCount += 1
    validateMode(entry.unixPermissions, entry.dir)
    validateMode(entry.dosPermissions, entry.dir)
    if (entry.dir) continue
    const internal = entry as typeof entry & {
      _data?: { compressedSize?: number; uncompressedSize?: number }
    }
    const compressed = internal._data?.compressedSize
    const uncompressed = internal._data?.uncompressedSize
    if (typeof compressed === 'number' && compressed > SKILL_PACKAGE_LIMITS.maxCompressedBytes)
      limited()
    if (typeof uncompressed === 'number') {
      if (uncompressed > SKILL_PACKAGE_LIMITS.maxFileBytes) limited()
      declaredTotal += uncompressed
      if (declaredTotal > SKILL_PACKAGE_LIMITS.maxUncompressedBytes) limited()
    }
  }
  if (manifestCount !== 1 || !zip.file('SKILL.md')) invalid()

  const files: SkillPackageFile[] = []
  let total = 0
  for (const entry of entries) {
    checkCancelled(signal)
    const extension = entry.name.split('.').pop()?.toLowerCase() ?? ''
    if (
      entry.name !== 'SKILL.md' &&
      !TEXT_EXTENSIONS.has(extension) &&
      !IMAGE_EXTENSIONS.has(extension)
    )
      invalid()
    const bytes = await inflateBounded(entry, total, signal)
    checkCancelled(signal)
    if (bytes.byteLength > SKILL_PACKAGE_LIMITS.maxFileBytes) limited()
    total += bytes.byteLength
    if (total > SKILL_PACKAGE_LIMITS.maxUncompressedBytes) limited()
    if (TEXT_EXTENSIONS.has(extension)) {
      try {
        utf8.decode(bytes)
      } catch {
        invalid()
      }
    } else if (IMAGE_EXTENSIONS.has(extension)) {
      validateImage(extension, bytes)
    }
    files.push(Object.freeze({ path: entry.name, bytes: bytes.slice() }))
  }
  files.sort((left, right) => (left.path < right.path ? -1 : left.path > right.path ? 1 : 0))
  const manifestFile = files.find((file) => file.path === 'SKILL.md')
  if (!manifestFile) invalid()
  let manifest: string
  try {
    manifest = utf8.decode(manifestFile.bytes)
  } catch {
    invalid()
  }
  return Object.freeze({ skill: parseSkillPackage(manifest), files: Object.freeze(files) })
}
