import { XMLParser, XMLValidator } from 'fast-xml-parser'
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
  beforeXml: Record<string, string>
  afterXml: Record<string, string>
  preservedHashes: Record<string, string>
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

const xmlParser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  processEntities: false,
})

function stableValue(value: unknown, preserveWhitespace = false): unknown {
  if (Array.isArray(value))
    return value
      .filter(
        (child) =>
          preserveWhitespace ||
          !(
            child &&
            typeof child === 'object' &&
            Object.keys(child).length === 1 &&
            typeof (child as Record<string, unknown>)['#text'] === 'string' &&
            /^\s*$/.test((child as Record<string, string>)['#text'])
          ),
      )
      .map((child) => stableValue(child, preserveWhitespace))
  if (!value || typeof value !== 'object') return value
  const record = value as Record<string, unknown>
  const attributes =
    record[':@'] && typeof record[':@'] === 'object'
      ? (record[':@'] as Record<string, unknown>)
      : undefined
  return Object.fromEntries(
    Object.entries(record)
      .sort(([first], [second]) => first.localeCompare(second))
      .map(([key, child]) => {
        const localName = key.includes(':') ? key.slice(key.lastIndexOf(':') + 1) : key
        const keepWhitespace =
          preserveWhitespace ||
          attributes?.['@_xml:space'] === 'preserve' ||
          ['t', 'instrText', 'delText'].includes(localName)
        return [key, stableValue(child, keepWhitespace)]
      }),
  )
}

function canonicalXml(value: string): string | undefined {
  try {
    if (XMLValidator.validate(value) !== true) return undefined
    return JSON.stringify(stableValue(xmlParser.parse(value)))
  } catch {
    return undefined
  }
}

function backgroundXml(value: string): string | undefined {
  return /<p:bg\b[^>]*\/>|<p:bg\b[^>]*>[\s\S]*?<\/p:bg\s*>/.exec(value)?.[0]
}

function withoutBackground(value: string): string | undefined {
  return canonicalXml(value.replace(/<p:bg\b[^>]*\/>|<p:bg\b[^>]*>[\s\S]*?<\/p:bg\s*>/, ''))
}

function backgroundOnly(before: string, after: string): boolean {
  const expected = backgroundXml(after)
  return (
    expected !== undefined &&
    canonicalXml(backgroundXml(before) ?? '') !== canonicalXml(expected) &&
    withoutBackground(before) === withoutBackground(after)
  )
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

function masterLayoutIdentity(xml: string): string[] {
  return [...xml.matchAll(/<p:sldLayoutId\b[^>]*\br:id=["']([^"']+)["'][^>]*\/?\s*>/g)].map(
    (match) => match[1],
  )
}

function compressedMetadata(file: unknown): { compressed?: number; uncompressed?: number } {
  const data = (file as { _data?: { compressedSize?: unknown; uncompressedSize?: unknown } })._data
  return {
    compressed: typeof data?.compressedSize === 'number' ? data.compressedSize : undefined,
    uncompressed: typeof data?.uncompressedSize === 'number' ? data.uncompressedSize : undefined,
  }
}

async function loadBoundedZip(base64: string, signal?: AbortSignal): Promise<JSZip> {
  if (signal?.aborted) throw new Error('cancelled')
  if (!base64 || base64.length > Math.ceil(MAX_PPTX_PACKAGE_BYTES / 3) * 4)
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
  return zip
}

export async function editPowerPointPackage(
  base64: string,
  kind: PackageEditKind,
  replacements: XmlReplacement[],
  signal?: AbortSignal,
): Promise<PackageEditResult> {
  if (replacements.length < 1 || replacements.length > 32) throw new Error('invalid_tool_input')
  const zip = await loadBoundedZip(base64, signal)
  const changedPaths = new Set<string>()
  const beforeHashes: Record<string, string> = {}
  const afterHashes: Record<string, string> = {}
  const beforeXml: Record<string, string> = {}
  const afterXml: Record<string, string> = {}
  const preservedHashes: Record<string, string> = {}
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
    if (
      kind === 'master' &&
      replacement.path.startsWith('ppt/slideMasters/') &&
      JSON.stringify(masterLayoutIdentity(before)) !==
        JSON.stringify(masterLayoutIdentity(replacement.xml))
    )
      throw new Error('office_api_unsupported')
    beforeHashes[replacement.path] = hash(before)
    afterHashes[replacement.path] = hash(replacement.xml)
    beforeXml[replacement.path] = before
    afterXml[replacement.path] = replacement.xml
    zip.file(replacement.path, replacement.xml)
    changedPaths.add(replacement.path)
  }
  for (const [path, file] of Object.entries(zip.files)) {
    if (signal?.aborted) throw new Error('cancelled')
    if (!file.dir && !changedPaths.has(path))
      preservedHashes[path] = hash(await file.async('uint8array'))
  }
  const output = await zip.generateAsync({
    type: 'base64',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
  if (signal?.aborted) throw new Error('cancelled')
  if (!output || output.length > Math.ceil(MAX_PPTX_PACKAGE_BYTES / 3) * 4)
    throw new Error('invalid_tool_input')
  return {
    base64: output,
    changedPaths: [...changedPaths],
    beforeHashes,
    afterHashes,
    beforeXml,
    afterXml,
    preservedHashes,
  }
}

export async function verifyImportedPowerPointPackage(
  base64: string,
  expected: Pick<
    PackageEditResult,
    'changedPaths' | 'beforeXml' | 'afterXml' | 'afterHashes' | 'preservedHashes'
  >,
  signal?: AbortSignal,
): Promise<boolean> {
  if (signal?.aborted) throw new Error('cancelled')
  try {
    const zip = await loadBoundedZip(base64, signal)
    const expectedPaths = new Set([
      ...expected.changedPaths,
      ...Object.keys(expected.preservedHashes),
    ])
    const actualPaths = Object.values(zip.files)
      .filter((file) => !file.dir)
      .map((file) => file.name)
    if (
      actualPaths.length !== expectedPaths.size ||
      actualPaths.some((path) => !expectedPaths.has(path))
    )
      return false
    for (const path of expected.changedPaths) {
      if (signal?.aborted) throw new Error('cancelled')
      const file = zip.file(path)
      const before = expected.beforeXml[path]
      const after = expected.afterXml[path]
      if (!file || before === undefined || after === undefined) return false
      const actual = await file.async('string')
      if (backgroundOnly(before, after)) {
        const actualBackground = backgroundXml(actual)
        const expectedBackground = backgroundXml(after)
        if (
          !actualBackground ||
          !expectedBackground ||
          canonicalXml(actualBackground) !== canonicalXml(expectedBackground)
        )
          return false
        if (withoutBackground(actual) !== withoutBackground(after)) return false
      } else if (hash(actual) !== expected.afterHashes[path]) {
        return false
      }
    }
    for (const [path, expectedHash] of Object.entries(expected.preservedHashes)) {
      if (signal?.aborted) throw new Error('cancelled')
      const file = zip.file(path)
      if (!file || hash(await file.async('uint8array')) !== expectedHash) return false
    }
    return true
  } catch (error) {
    if (error instanceof Error && error.message === 'cancelled') throw error
    return false
  }
}

export async function verifyPowerPointPackage(
  base64: string,
  expected: Pick<PackageEditResult, 'changedPaths' | 'afterHashes' | 'preservedHashes'>,
  signal?: AbortSignal,
): Promise<boolean> {
  if (signal?.aborted) throw new Error('cancelled')
  try {
    const zip = await loadBoundedZip(base64, signal)
    const expectedPaths = new Set([
      ...expected.changedPaths,
      ...Object.keys(expected.preservedHashes),
    ])
    const actualPaths = Object.values(zip.files)
      .filter((file) => !file.dir)
      .map((file) => file.name)
    if (
      actualPaths.length !== expectedPaths.size ||
      actualPaths.some((path) => !expectedPaths.has(path))
    )
      return false
    for (const path of expected.changedPaths) {
      if (signal?.aborted) throw new Error('cancelled')
      const file = zip.file(path)
      if (!file || hash(await file.async('string')) !== expected.afterHashes[path]) return false
    }
    for (const [path, expectedHash] of Object.entries(expected.preservedHashes)) {
      if (signal?.aborted) throw new Error('cancelled')
      const file = zip.file(path)
      if (!file || hash(await file.async('uint8array')) !== expectedHash) return false
    }
    return true
  } catch (error) {
    if (error instanceof Error && error.message === 'cancelled') throw error
    return false
  }
}

export async function verifyPowerPointPackageInputs(
  base64: string,
  expectedHashes: Readonly<Record<string, string>>,
  signal?: AbortSignal,
): Promise<boolean> {
  if (signal?.aborted) throw new Error('cancelled')
  try {
    const zip = await loadBoundedZip(base64, signal)
    for (const [path, expectedHash] of Object.entries(expectedHashes)) {
      if (signal?.aborted) throw new Error('cancelled')
      const file = zip.file(path)
      if (!file || hash(await file.async('string')) !== expectedHash) return false
    }
    return true
  } catch (error) {
    if (error instanceof Error && error.message === 'cancelled') throw error
    return false
  }
}

export async function capturePowerPointPackage(
  base64: string,
  signal?: AbortSignal,
): Promise<Pick<PackageEditResult, 'changedPaths' | 'afterHashes' | 'preservedHashes'>> {
  const zip = await loadBoundedZip(base64, signal)
  const preservedHashes: Record<string, string> = {}
  for (const [path, file] of Object.entries(zip.files)) {
    if (signal?.aborted) throw new Error('cancelled')
    if (!file.dir) preservedHashes[path] = hash(await file.async('uint8array'))
  }
  return { changedPaths: [], afterHashes: {}, preservedHashes }
}
