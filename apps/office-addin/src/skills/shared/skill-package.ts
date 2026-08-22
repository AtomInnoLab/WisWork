import JSZip from 'jszip'
import { parseSkillPackage, type ParsedSkill } from './skill-registry.js'
import { validateSkillZip, type SkillZipEntry } from './skill-zip-validator.js'

export const SKILL_PACKAGE_LIMITS = Object.freeze({
  maxCompressedBytes: 2 * 1024 * 1024,
  maxUncompressedBytes: 8 * 1024 * 1024,
  maxFileBytes: 2 * 1024 * 1024,
  maxEntries: 64,
  maxPathBytes: 256,
  maxImagePixels: 16_000_000,
})

export interface SkillPackageFile {
  readonly path: string
  readonly bytes: Uint8Array
}

export interface ParsedSkillArchive {
  readonly skill: ParsedSkill
  readonly files: readonly SkillPackageFile[]
}

export interface DecodedSkillImage {
  readonly width: number
  readonly height: number
  close(): void
}
export type SkillImageDecoder = (
  bytes: Uint8Array,
  mime: 'image/png' | 'image/jpeg' | 'image/webp',
) => Promise<DecodedSkillImage>

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

function validateUnixMode(mode: number | string | null | undefined, directory: boolean): void {
  const numeric = typeof mode === 'string' ? Number.parseInt(mode, 8) : mode
  if (typeof numeric !== 'number' || !Number.isFinite(numeric)) return
  const type = numeric & 0o170000
  if (
    (type !== 0 && type !== 0o100000 && type !== 0o040000) ||
    (!directory && (numeric & 0o111) !== 0)
  )
    invalid()
}

function validateDosAttributes(attributes: number | null | undefined, directory: boolean): void {
  if (attributes === null || attributes === undefined) return
  if (
    !Number.isInteger(attributes) ||
    (attributes & 0x04) !== 0 ||
    (attributes & 0x08) !== 0 ||
    Boolean(attributes & 0x10) !== directory ||
    (attributes & ~(0x01 | 0x02 | 0x10 | 0x20)) !== 0
  )
    invalid()
}

export async function validateSkillPackageImage(
  extension: string,
  bytes: Uint8Array,
  decode: SkillImageDecoder = decodeBrowserImage,
): Promise<{ width: number; height: number }> {
  const starts = (...values: number[]) => values.every((value, index) => bytes[index] === value)
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let width = 0
  let height = 0
  if (extension === 'png') {
    if (bytes.length < 45 || !starts(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a)) invalid()
    let offset = 8
    let first = true
    let ended = false
    let hasImageData = false
    let imageDataEnded = false
    let bitDepth = 0
    let colorType = 0
    let palette = false
    const compressed: Uint8Array[] = []
    while (offset + 12 <= bytes.length) {
      const size = view.getUint32(offset)
      const type = String.fromCharCode(...bytes.subarray(offset + 4, offset + 8))
      if (offset + 12 + size > bytes.length) invalid()
      if (!/^[A-Za-z]{4}$/.test(type) || type[2] !== type[2].toUpperCase()) invalid()
      if (/^[A-Z]/.test(type) && !['IHDR', 'PLTE', 'IDAT', 'IEND'].includes(type)) invalid()
      if (
        pngCrc(bytes.subarray(offset + 4, offset + 8 + size)) !== view.getUint32(offset + 8 + size)
      )
        invalid()
      if (first) {
        if (type !== 'IHDR' || size !== 13) invalid()
        width = view.getUint32(offset + 8)
        height = view.getUint32(offset + 12)
        bitDepth = bytes[offset + 16]
        colorType = bytes[offset + 17]
        if (
          !new Set([
            '0:1',
            '0:2',
            '0:4',
            '0:8',
            '0:16',
            '2:8',
            '2:16',
            '3:1',
            '3:2',
            '3:4',
            '3:8',
            '4:8',
            '4:16',
            '6:8',
            '6:16',
          ]).has(`${colorType}:${bitDepth}`) ||
          bytes[offset + 18] ||
          bytes[offset + 19] ||
          bytes[offset + 20]
        )
          invalid()
        first = false
      } else if (type === 'IHDR') invalid()
      if (type === 'PLTE') {
        if (
          palette ||
          hasImageData ||
          !size ||
          size % 3 ||
          size > 768 ||
          [0, 4].includes(colorType) ||
          (colorType === 3 && size / 3 > 2 ** bitDepth)
        )
          invalid()
        palette = true
      }
      if (type === 'IDAT') {
        if (imageDataEnded || !size) invalid()
        hasImageData = true
        compressed.push(bytes.slice(offset + 8, offset + 8 + size))
      } else if (hasImageData && type !== 'IEND') imageDataEnded = true
      offset += 12 + size
      if (type === 'IEND') {
        if (size !== 0 || offset !== bytes.length) invalid()
        ended = true
        break
      }
    }
    if (!ended || !hasImageData || (colorType === 3 && !palette)) invalid()
    const channels = ({ 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 } as Record<number, number>)[colorType]
    const rowBytes = Math.ceil((width * channels * bitDepth) / 8)
    const expected = height * (rowBytes + 1)
    if (!Number.isSafeInteger(expected) || expected > SKILL_PACKAGE_LIMITS.maxImagePixels * 8)
      limited()
    const joined = new Uint8Array(compressed.reduce((sum, value) => sum + value.length, 0))
    let joinedOffset = 0
    for (const value of compressed) {
      joined.set(value, joinedOffset)
      joinedOffset += value.length
    }
    const scanlines = await inflatePng(joined, expected)
    for (let row = 0; row < height; row++) if (scanlines[row * (rowBytes + 1)] > 4) invalid()
  } else if (extension === 'jpg' || extension === 'jpeg') {
    if (bytes.length < 10 || !starts(0xff, 0xd8) || bytes.at(-2) !== 0xff || bytes.at(-1) !== 0xd9)
      invalid()
    let offset = 2
    const components = new Set<number>()
    let sawSof = false
    let sawSos = false
    while (offset + 4 <= bytes.length - 2) {
      if (bytes[offset] !== 0xff) invalid()
      const marker = bytes[offset + 1]
      offset += 2
      if (marker === 0xda) {
        const size = view.getUint16(offset)
        const count = bytes[offset + 2]
        if (!sawSof || size !== 6 + 2 * count || count < 1 || offset + size > bytes.length - 2)
          invalid()
        for (let index = 0; index < count; index++)
          if (!components.has(bytes[offset + 3 + index * 2])) invalid()
        offset += size
        const entropyStart = offset
        while (offset < bytes.length - 2) {
          if (bytes[offset] !== 0xff) {
            offset++
            continue
          }
          const next = bytes[offset + 1]
          if (next === 0 || (next >= 0xd0 && next <= 0xd7)) {
            offset += 2
            continue
          }
          invalid()
        }
        if (offset === entropyStart) invalid()
        sawSos = true
        break
      }
      if (marker === 0xd8 || marker === 0xd9 || (marker >= 0xd0 && marker <= 0xd7)) continue
      const size = view.getUint16(offset)
      if (size < 2 || offset + size > bytes.length) invalid()
      if (
        [0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(
          marker,
        )
      ) {
        if (sawSof || size < 11 || bytes[offset + 2] !== 8) invalid()
        height = view.getUint16(offset + 3)
        width = view.getUint16(offset + 5)
        const count = bytes[offset + 7]
        if (count < 1 || count > 4 || size !== 8 + 3 * count) invalid()
        for (let index = 0; index < count; index++) {
          const id = bytes[offset + 8 + 3 * index]
          const sampling = bytes[offset + 9 + 3 * index]
          if (
            components.has(id) ||
            !(sampling >> 4) ||
            !(sampling & 15) ||
            sampling >> 4 > 4 ||
            (sampling & 15) > 4 ||
            bytes[offset + 10 + 3 * index] > 3
          )
            invalid()
          components.add(id)
        }
        sawSof = true
      }
      offset += size
    }
    if (!sawSof || !sawSos) invalid()
  } else if (extension === 'webp') {
    if (
      !starts(0x52, 0x49, 0x46, 0x46) ||
      bytes.length < 30 ||
      String.fromCharCode(...bytes.subarray(8, 12)) !== 'WEBP' ||
      view.getUint32(4, true) + 8 !== bytes.length
    )
      invalid()
    let offset = 12
    let canvasWidth = 0
    let canvasHeight = 0
    let payloads = 0
    while (offset + 8 <= bytes.length) {
      const type = String.fromCharCode(...bytes.subarray(offset, offset + 4))
      const size = view.getUint32(offset + 4, true)
      if (offset + 8 + size > bytes.length) invalid()
      if (type === 'VP8X' && size >= 10) {
        if (canvasWidth || size !== 10) invalid()
        canvasWidth =
          1 + bytes[offset + 12] + (bytes[offset + 13] << 8) + (bytes[offset + 14] << 16)
        canvasHeight =
          1 + bytes[offset + 15] + (bytes[offset + 16] << 8) + (bytes[offset + 17] << 16)
      } else if (
        type === 'VP8 ' &&
        size > 10 &&
        bytes[offset + 11] === 0x9d &&
        bytes[offset + 12] === 0x01 &&
        bytes[offset + 13] === 0x2a
      ) {
        const tag = bytes[offset + 8] | (bytes[offset + 9] << 8) | (bytes[offset + 10] << 16)
        if (tag & 1 || tag >>> 5 > size - 3) invalid()
        width = view.getUint16(offset + 14, true) & 0x3fff
        height = view.getUint16(offset + 16, true) & 0x3fff
        payloads++
      } else if (type === 'VP8L' && size > 5 && bytes[offset + 8] === 0x2f) {
        if (bytes[offset + 12] & 0xe0) invalid()
        width = 1 + bytes[offset + 9] + ((bytes[offset + 10] & 0x3f) << 8)
        height =
          1 +
          (bytes[offset + 10] >> 6) +
          (bytes[offset + 11] << 2) +
          ((bytes[offset + 12] & 0x0f) << 10)
        payloads++
      }
      offset += 8 + size + (size & 1)
    }
    if (offset !== bytes.length) invalid()
    if (payloads !== 1 || (canvasWidth && (canvasWidth !== width || canvasHeight !== height)))
      invalid()
  }
  if (!width || !height) invalid()
  if (width * height > SKILL_PACKAGE_LIMITS.maxImagePixels) limited()
  const mime =
    extension === 'png' ? 'image/png' : extension === 'webp' ? 'image/webp' : 'image/jpeg'
  let decoded: DecodedSkillImage | undefined
  try {
    decoded = await decode(bytes, mime)
    if (
      typeof decoded.close !== 'function' ||
      decoded.width !== width ||
      decoded.height !== height ||
      decoded.width * decoded.height > SKILL_PACKAGE_LIMITS.maxImagePixels
    )
      invalid()
  } catch {
    invalid()
  } finally {
    try {
      decoded?.close()
    } catch {
      invalid()
    }
  }
  return { width, height }
}

async function decodeBrowserImage(
  bytes: Uint8Array,
  mime: 'image/png' | 'image/jpeg' | 'image/webp',
): Promise<DecodedSkillImage> {
  if (typeof createImageBitmap !== 'function') invalid()
  const copy = new Uint8Array(bytes.length)
  copy.set(bytes)
  return createImageBitmap(new Blob([copy.buffer], { type: mime }))
}

function pngCrc(bytes: Uint8Array): number {
  let crc = 0xffffffff
  for (const byte of bytes) crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8)
  return (crc ^ 0xffffffff) >>> 0
}

async function inflatePng(bytes: Uint8Array, expected: number): Promise<Uint8Array> {
  let stream: ReadableStream<Uint8Array>
  try {
    const copy = new Uint8Array(bytes.byteLength)
    copy.set(bytes)
    stream = new Blob([copy.buffer]).stream().pipeThrough(new DecompressionStream('deflate'))
  } catch {
    invalid()
  }
  const reader = stream.getReader()
  const chunks: Uint8Array[] = []
  let total = 0
  try {
    while (true) {
      const result = await reader.read()
      if (result.done) break
      total += result.value.length
      if (total > expected) {
        await reader.cancel()
        invalid()
      }
      chunks.push(result.value.slice())
    }
  } catch {
    invalid()
  }
  if (total !== expected) invalid()
  const output = new Uint8Array(total)
  let offset = 0
  for (const chunk of chunks) {
    output.set(chunk, offset)
    offset += chunk.length
  }
  return output
}

const CRC_TABLE = Uint32Array.from({ length: 256 }, (_, value) => {
  let crc = value
  for (let bit = 0; bit < 8; bit++) crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1
  return crc >>> 0
})

async function inflateBounded(
  entry: JSZip.JSZipObject,
  currentTotal: number,
  metadata: SkillZipEntry,
  signal?: AbortSignal,
): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const stream = (entry as ZipStreamEntry).internalStream('uint8array')
    const chunks: Uint8Array[] = []
    let size = 0
    let settled = false
    let crc = 0xffffffff
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
      for (const byte of chunk) crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8)
    })
    stream.on('error', () => failOnce('invalid_skill_package'))
    stream.on('end', () => {
      signal?.removeEventListener('abort', cancel)
      if (settled) return
      settled = true
      if (size !== metadata.uncompressed || (crc ^ 0xffffffff) >>> 0 !== metadata.crc)
        return reject(new Error('invalid_skill_package'))
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
  imageDecoder?: SkillImageDecoder,
): Promise<ParsedSkillArchive> {
  checkCancelled(signal)
  if (!(source instanceof Uint8Array)) invalid()
  if (source.byteLength > SKILL_PACKAGE_LIMITS.maxCompressedBytes) limited()
  const metadata = validateSkillZip(source, SKILL_PACKAGE_LIMITS)
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(source, { createFolders: false })
  } catch {
    invalid()
  }
  checkCancelled(signal)
  const allEntries = Object.values(zip.files)
  if (
    allEntries.length !== metadata.size ||
    [...metadata.keys()].some((name) => !Object.hasOwn(zip.files, name))
  )
    invalid()
  const entries = allEntries.filter((entry) => !entry.dir)
  const names = new Set<string>()
  let manifestCount = 0
  for (const entry of allEntries) {
    const unsafeName = (entry as typeof entry & { unsafeOriginalName?: string }).unsafeOriginalName
    if (unsafeName && unsafeName !== entry.name) invalid()
    const path = validatePath(entry.name)
    const collisionKey = path.normalize('NFC').toLocaleLowerCase('en-US')
    if (names.has(collisionKey)) invalid()
    names.add(collisionKey)
    if (!entry.dir && path.toLocaleLowerCase('en-US') === 'skill.md') manifestCount += 1
    validateUnixMode(entry.unixPermissions, entry.dir)
    validateDosAttributes(entry.dosPermissions, entry.dir)
    if (entry.dir) continue
    const declared = metadata.get(entry.name)
    if (!declared || declared.directory !== entry.dir) invalid()
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
    const declared = metadata.get(entry.name)
    if (!declared) invalid()
    const bytes = await inflateBounded(entry, total, declared, signal)
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
      await validateSkillPackageImage(extension, bytes, imageDecoder)
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
