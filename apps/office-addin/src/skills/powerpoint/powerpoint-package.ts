import { XMLValidator } from 'fast-xml-parser'
import JSZip from 'jszip'

export const MAX_PPTX_PACKAGE_BYTES = 8 * 1024 * 1024
export const MAX_PPTX_ENTRY_BYTES = 2 * 1024 * 1024
export const MAX_PPTX_ENTRIES = 256
export const MAX_PPTX_XML_BYTES = 512 * 1024

export type PackageEditKind = 'slide' | 'chart' | 'master'
export interface XmlReplacement {
  path: string
  xml: string
}
export interface PackageEditResult {
  base64: string
  changedPaths: string[]
  beforeHashes: Record<string, string>
  afterHashes: Record<string, string>
}

function hash(value: string | Uint8Array): string {
  const bytes = typeof value === 'string' ? new TextEncoder().encode(value) : value
  let result = 0x811c9dc5
  for (const byte of bytes) {
    result ^= byte
    result = Math.imul(result, 0x01000193)
  }
  return `${bytes.byteLength}:${(result >>> 0).toString(16).padStart(8, '0')}`
}

function validPath(path: string): boolean {
  return (
    path.length > 0 &&
    path.length <= 256 &&
    !path.startsWith('/') &&
    !path.includes('\\') &&
    !path.includes('\0') &&
    path.split('/').every((part) => part && part !== '.' && part !== '..')
  )
}

function allowed(kind: PackageEditKind, path: string): boolean {
  if (kind === 'slide') return path === 'ppt/slides/slide1.xml'
  if (kind === 'chart') return /^ppt\/charts\/(chart|style|colors)\d+\.xml$/.test(path)
  return (
    /^ppt\/(slideMasters|slideLayouts)\/[A-Za-z0-9._-]+\.xml$/.test(path) ||
    /^ppt\/theme\/theme\d+\.xml$/.test(path)
  )
}

function compressedMetadata(file: unknown): { compressed?: number; uncompressed?: number } {
  const data = (file as { _data?: { compressedSize?: unknown; uncompressedSize?: unknown } })._data
  return {
    compressed: typeof data?.compressedSize === 'number' ? data.compressedSize : undefined,
    uncompressed: typeof data?.uncompressedSize === 'number' ? data.uncompressedSize : undefined,
  }
}

export async function editPowerPointPackage(
  base64: string,
  kind: PackageEditKind,
  replacements: XmlReplacement[],
  signal?: AbortSignal,
): Promise<PackageEditResult> {
  if (signal?.aborted) throw new Error('cancelled')
  if (
    !base64 ||
    base64.length > Math.ceil(MAX_PPTX_PACKAGE_BYTES / 3) * 4 ||
    replacements.length < 1 ||
    replacements.length > 32
  )
    throw new Error('invalid_tool_input')
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(base64, { base64: true, checkCRC32: true, createFolders: false })
  } catch {
    throw new Error('invalid_tool_input')
  }
  if (signal?.aborted) throw new Error('cancelled')
  const files = Object.values(zip.files)
  if (files.length > MAX_PPTX_ENTRIES) throw new Error('invalid_tool_input')
  let total = 0
  for (const file of files) {
    const originalName = (file as typeof file & { unsafeOriginalName?: string }).unsafeOriginalName
    if (
      (originalName && originalName !== file.name) ||
      !validPath(file.dir ? file.name.replace(/\/$/, '') : file.name)
    )
      throw new Error('invalid_tool_input')
    if (file.dir) continue
    const metadata = compressedMetadata(file)
    if (
      metadata.uncompressed === undefined ||
      metadata.compressed === undefined ||
      metadata.uncompressed > MAX_PPTX_ENTRY_BYTES
    )
      throw new Error('invalid_tool_input')
    total += metadata.uncompressed
    if (total > MAX_PPTX_PACKAGE_BYTES) throw new Error('invalid_tool_input')
  }
  const changedPaths = new Set<string>()
  const beforeHashes: Record<string, string> = {}
  const afterHashes: Record<string, string> = {}
  for (const replacement of replacements) {
    if (signal?.aborted) throw new Error('cancelled')
    if (
      !validPath(replacement.path) ||
      !allowed(kind, replacement.path) ||
      changedPaths.has(replacement.path) ||
      new TextEncoder().encode(replacement.xml).byteLength > MAX_PPTX_XML_BYTES ||
      /<!\s*(?:DOCTYPE|ENTITY)\b/i.test(replacement.xml) ||
      XMLValidator.validate(replacement.xml) !== true
    )
      throw new Error('invalid_tool_input')
    const file = zip.file(replacement.path)
    if (!file) throw new Error('invalid_tool_input')
    const before = await file.async('string')
    if (signal?.aborted) throw new Error('cancelled')
    beforeHashes[replacement.path] = hash(before)
    afterHashes[replacement.path] = hash(replacement.xml)
    zip.file(replacement.path, replacement.xml)
    changedPaths.add(replacement.path)
  }
  const output = await zip.generateAsync({
    type: 'base64',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
  if (signal?.aborted) throw new Error('cancelled')
  if (!output || output.length > Math.ceil(MAX_PPTX_PACKAGE_BYTES / 3) * 4)
    throw new Error('invalid_tool_input')
  return { base64: output, changedPaths: [...changedPaths], beforeHashes, afterHashes }
}

export async function verifyPowerPointPackage(
  base64: string,
  expected: Pick<PackageEditResult, 'changedPaths' | 'afterHashes'>,
  signal?: AbortSignal,
): Promise<boolean> {
  if (signal?.aborted) throw new Error('cancelled')
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(base64, { base64: true, checkCRC32: true })
  } catch {
    return false
  }
  for (const path of expected.changedPaths) {
    if (signal?.aborted) throw new Error('cancelled')
    const file = zip.file(path)
    if (!file) return false
    if (hash(await file.async('string')) !== expected.afterHashes[path]) return false
  }
  return true
}
