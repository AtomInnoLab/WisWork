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

const embeddedWorkbookPath = /^ppt\/embeddings\/[A-Za-z0-9_.-]+\.xlsx$/

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

// Chart XML edits may change presentation styling, but cached values and their data links
// require a separate workbook-aware operation with a savepoint and source readback.
function chartDataIdentity(xml: string): string {
  if (new TextEncoder().encode(xml).byteLength > MAX_PPTX_XML_BYTES ||
    /<!\s*(?:DOCTYPE|ENTITY)\b/i.test(xml) || XMLValidator.validate(xml) !== true)
    throw new Error('office_api_unsupported')
  const protectedTags = new Set([
    'ser', 'f', 'numCache', 'strCache', 'multiLvlStrCache',
    'externalData', 'pivotSource', 'extLst',
  ])
  const found: unknown[] = []
  const visit = (nodes: unknown): void => {
    if (!Array.isArray(nodes)) return
    for (const node of nodes) {
      if (!node || typeof node !== 'object') continue
      for (const [tag, value] of Object.entries(node)) {
        if (protectedTags.has(tag.split(':').at(-1)!)) found.push([tag, stableValue(node)])
        else visit(value)
      }
    }
  }
  visit(xmlParser.parse(xml))
  return JSON.stringify(found)
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

export async function loadBoundedZip(
  base64: string,
  signal?: AbortSignal,
  checkCRC32 = true,
  maxBytes = MAX_PPTX_PACKAGE_BYTES,
): Promise<JSZip> {
  if (signal?.aborted) throw new Error('cancelled')
  if (!base64 || base64.length > Math.ceil(maxBytes / 3) * 4)
    throw new Error('invalid_tool_input')
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(base64, { base64: true, checkCRC32, createFolders: false })
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
    if (total > maxBytes) throw new Error('invalid_tool_input')
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
    if (kind === 'chart' && /^ppt\/charts\/chart\d+\.xml$/.test(replacement.path) &&
      chartDataIdentity(before) !== chartDataIdentity(replacement.xml))
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

/** Build exact import/readback evidence for a synchronized chart XML and embedded XLSX edit. */
export async function captureChartValuePackageEdit(beforeBase64: string, afterBase64: string, signal?: AbortSignal): Promise<PackageEditResult> {
  const [before, after] = await Promise.all([loadBoundedZip(beforeBase64, signal), loadBoundedZip(afterBase64, signal)])
  const beforePaths = Object.keys(before.files).filter((path) => !before.files[path]!.dir).sort()
  const afterPaths = Object.keys(after.files).filter((path) => !after.files[path]!.dir).sort()
  if (JSON.stringify(beforePaths) !== JSON.stringify(afterPaths)) throw new Error('office_api_unsupported')
  const result: PackageEditResult = { base64: afterBase64, changedPaths: [], beforeHashes: {}, afterHashes: {}, beforeXml: {}, afterXml: {}, preservedHashes: {} }
  for (const path of beforePaths) {
    if (signal?.aborted) throw new Error('cancelled')
    const oldBytes = await before.file(path)!.async('uint8array')
    const newBytes = await after.file(path)!.async('uint8array')
    const oldHash = hash(oldBytes), newHash = hash(newBytes)
    if (oldHash === newHash) { result.preservedHashes[path] = oldHash; continue }
    if (!/^ppt\/charts\/chart\d+\.xml$/.test(path) && !embeddedWorkbookPath.test(path)) throw new Error('office_api_unsupported')
    result.changedPaths.push(path)
    result.beforeHashes[path] = oldHash
    result.afterHashes[path] = newHash
    result.beforeXml[path] = embeddedWorkbookPath.test(path) ? '' : new TextDecoder().decode(oldBytes)
    result.afterXml[path] = embeddedWorkbookPath.test(path) ? '' : new TextDecoder().decode(newBytes)
  }
  if (result.changedPaths.length !== 2 || result.changedPaths.filter((path) => embeddedWorkbookPath.test(path)).length !== 1) throw new Error('office_api_unsupported')
  return result
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
      if (embeddedWorkbookPath.test(path)) {
        if (hash(await file.async('uint8array')) !== expected.afterHashes[path]) return false
        continue
      }
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
      if (!file || hash(await file.async('uint8array')) !== expected.afterHashes[path]) return false
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
      if (!file || hash(await file.async('uint8array')) !== expectedHash) return false
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

export interface PowerPointPicturePackageInspection {
  pictureFingerprint: string
  mediaDigest: string
  shapeIds: string[]
}
/** Inspect a conservative, lossless subset: one ordinary top-level embedded raster picture. */
export async function inspectPowerPointPicturePackage(
  base64: string,
  shapeId: string,
  signal?: AbortSignal,
  original?: (base64: string) => void,
  options: { slideIndex?: number; maxBytes?: number } = {},
): Promise<PowerPointPicturePackageInspection> {
  const unsupported = (): never => {
    throw new Error('office_api_unsupported')
  }
  const zip = await loadBoundedZip(base64, signal, true, options.maxBytes)
  const paths = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  const path = options.slideIndex === undefined
    ? paths.length === 1 ? paths[0] : undefined
    : paths.includes(`ppt/slides/slide${options.slideIndex + 1}.xml`)
      ? `ppt/slides/slide${options.slideIndex + 1}.xml` : undefined
  if (!path) return unsupported()
  const xml = await zip.file(path)!.async('string')
  if (
    xml.length > MAX_PPTX_XML_BYTES ||
    /<!\s*(?:DOCTYPE|ENTITY)\b/i.test(xml) ||
    XMLValidator.validate(xml) !== true
  )
    unsupported()
  type Node = Record<string, unknown>
  const children = (node: Node, tag: string): Node[] =>
    Array.isArray(node[tag]) ? (node[tag] as Node[]) : []
  const child = (nodes: Node[], tag: string): Node =>
    nodes.find((node) => Object.hasOwn(node, tag)) ?? {}
  const root = xmlParser.parse(xml) as Node[]
  const slide = children(child(root, 'p:sld'), 'p:sld')
  // A slide animation can target the picture indirectly through nested timing nodes.
  if (slide.some((node) => Object.hasOwn(node, 'p:timing'))) unsupported()
  const tree = children(child(children(child(slide, 'p:cSld'), 'p:cSld'), 'p:spTree'), 'p:spTree')
  const shapes = tree.filter((node) =>
    ['p:sp', 'p:pic', 'p:graphicFrame', 'p:cxnSp', 'p:grpSp'].some((tag) =>
      Object.hasOwn(node, tag),
    ),
  )
  const findNv = (node: Node): Node | undefined => {
    if (Object.hasOwn(node, 'p:cNvPr')) return node
    for (const value of Object.values(node))
      if (Array.isArray(value))
        for (const nested of value as Node[]) {
          const found = findNv(nested)
          if (found) return found
        }
    return undefined
  }
  const ids = shapes.map((node) => (findNv(node)?.[':@'] as Node | undefined)?.['@_id'])
  if (
    ids.length > 100 ||
    ids.some((id) => typeof id !== 'string' || !id.length || id.length > 256) ||
    new Set(ids).size !== ids.length
  )
    unsupported()
  const picture = shapes[ids.indexOf(shapeId)]
  if (!picture || !Object.hasOwn(picture, 'p:pic')) unsupported()
  const allowed: Record<string, string[]> = {
    'p:pic': [],
    'p:nvPicPr': [],
    'p:cNvPr': ['id', 'name', 'descr', 'title'],
    'p:cNvPicPr': [],
    'p:nvPr': [],
    'a:picLocks': ['noChangeAspect'],
    'p:blipFill': ['dpi', 'rotWithShape'],
    'a:blip': ['r:embed', 'cstate'],
    'a:stretch': [],
    'a:fillRect': [],
    'a:srcRect': ['l', 't', 'r', 'b'],
    'p:spPr': ['bwMode'],
    'a:xfrm': ['rot'],
    'a:off': ['x', 'y'],
    'a:ext': ['cx', 'cy'],
    'a:prstGeom': ['prst'],
    'a:avLst': [],
  }
  let embed: string | undefined
  const inspect = (node: Node) => {
    const tag = Object.keys(node).find((key) => key !== ':@' && key !== '#text')
    if (!tag) {
      if (typeof node['#text'] === 'string' && !/^\s*$/.test(node['#text'])) unsupported()
      return
    }
    if (!Object.hasOwn(allowed, tag)) unsupported()
    const attributes = (node[':@'] ?? {}) as Node
    if (Object.keys(attributes).some((key) => !allowed[tag]!.includes(key.slice(2)))) unsupported()
    if (
      tag === 'p:spPr' &&
      attributes['@_bwMode'] !== undefined &&
      attributes['@_bwMode'] !== 'auto'
    )
      unsupported()
    if (
      tag === 'p:blipFill' &&
      attributes['@_rotWithShape'] !== undefined &&
      !['1', 'true'].includes(String(attributes['@_rotWithShape']))
    )
      unsupported()
    if (tag === 'p:blipFill' && attributes['@_dpi'] !== undefined) unsupported()
    if (tag === 'a:srcRect' && Object.values(attributes).some((value) => Number(value) !== 0))
      unsupported()
    if (tag === 'a:prstGeom' && attributes['@_prst'] !== 'rect') unsupported()
    if (tag === 'a:blip') {
      if (embed || typeof attributes['@_r:embed'] !== 'string') unsupported()
      embed = attributes['@_r:embed'] as string
    }
    for (const nested of children(node, tag)) inspect(nested)
  }
  inspect(picture!)
  if (!embed) unsupported()
  const relPath = path.replace('/slides/', '/slides/_rels/') + '.rels'
  const relFile = zip.file(relPath)
  if (!relFile) unsupported()
  const relXml = await relFile!.async('string')
  if (
    relXml.length > MAX_PPTX_XML_BYTES ||
    /<!\s*(?:DOCTYPE|ENTITY)\b/i.test(relXml) ||
    XMLValidator.validate(relXml) !== true
  )
    unsupported()
  const rels = children(child(xmlParser.parse(relXml) as Node[], 'Relationships'), 'Relationships')
    .filter((node) => Object.hasOwn(node, 'Relationship'))
    .map((node) => (node[':@'] ?? {}) as Node)
  if (new Set(rels.map((rel) => rel['@_Id'])).size !== rels.length) unsupported()
  const rel = rels.find((item) => item['@_Id'] === embed)
  if (
    !rel ||
    rel['@_TargetMode'] !== undefined ||
    rel['@_Type'] !== 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' ||
    typeof rel['@_Target'] !== 'string' ||
    !/^\.\.\/media\/[A-Za-z0-9_.-]+\.(?:png|jpe?g)$/i.test(rel['@_Target'])
  )
    unsupported()
  const media = zip.file('ppt/' + (rel!['@_Target'] as string).slice(3))
  if (!media) unsupported()
  const bytes = await media!.async('uint8array')
  if (
    !bytes.length ||
    bytes.length > 2 * 1024 * 1024 ||
    (!(bytes[0] === 137 && bytes[1] === 80 && bytes[2] === 78 && bytes[3] === 71) &&
      !(bytes[0] === 255 && bytes[1] === 216))
  )
    unsupported()
  const sha = async (value: Uint8Array) =>
    Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', new Uint8Array(value))))
      .map((byte) => byte.toString(16).padStart(2, '0'))
      .join('')
  const mediaDigest = await sha(bytes)
  const pictureFingerprint = await sha(
    new TextEncoder().encode(JSON.stringify([stableValue(picture), mediaDigest])),
  )
  if (signal?.aborted) throw new Error('cancelled')
  if (original) original(btoa(Array.from(bytes, (byte) => String.fromCharCode(byte)).join('')))
  return { pictureFingerprint, mediaDigest, shapeIds: ids as string[] }
}

/** Read only bounded chunks: ZIP headers are untrusted and may understate inflated size. */
async function boundedEntryBytes(
  file: JSZip.JSZipObject,
  remaining: number,
  signal?: AbortSignal,
): Promise<Uint8Array> {
  if (signal?.aborted) throw new Error('cancelled')
  // JSZip 3.10 exposes this browser stream API, but omits it from JSZipObject's types.
  const source = file as JSZip.JSZipObject & {
    internalStream(type: 'uint8array'): JSZip.JSZipStreamHelper<Uint8Array>
  }
  const stream = source.internalStream('uint8array')
  return new Promise((resolve, reject) => {
    const chunks: Uint8Array[] = []
    let size = 0,
      settled = false
    const fail = (message: string) => {
      if (settled) return
      settled = true
      stream.pause()
      signal?.removeEventListener('abort', abort)
      chunks.length = 0
      reject(new Error(message))
    }
    const abort = () => fail('cancelled')
    signal?.addEventListener('abort', abort, { once: true })
    stream
      .on('data', (chunk) => {
        if (settled) return
        if (signal?.aborted) {
          abort()
          return
        }
        size += chunk.byteLength
        if (size > MAX_PPTX_ENTRY_BYTES || size > remaining) {
          fail('invalid_tool_input')
          return
        }
        chunks.push(chunk)
      })
      .on('error', () => fail('invalid_tool_input'))
      .on('end', () => {
        if (settled) return
        settled = true
        signal?.removeEventListener('abort', abort)
        const result = new Uint8Array(size)
        let offset = 0
        for (const chunk of chunks) {
          result.set(chunk, offset)
          offset += chunk.byteLength
        }
        resolve(result)
      })
    stream.resume()
  })
}

/** Hash entry bytes and paths, excluding ZIP compression and timestamp metadata. */
export async function presentationPackageDigest(
  base64: string,
  signal?: AbortSignal,
): Promise<string> {
  // Parse the index without CRC inflation; validate declared limits before reading any entry.
  const zip = await loadBoundedZip(base64, signal, false)
  const sha = async (bytes: Uint8Array): Promise<string> => {
    if (signal?.aborted) throw new Error('cancelled')
    const digest = await crypto.subtle.digest('SHA-256', new Uint8Array(bytes))
    if (signal?.aborted) throw new Error('cancelled')
    return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('')
  }
  const entries: [string, string][] = []
  let total = 0
  for (const path of Object.keys(zip.files).sort()) {
    const file = zip.files[path]!
    if (file.dir) continue
    const bytes = await boundedEntryBytes(file, MAX_PPTX_PACKAGE_BYTES - total, signal)
    total += bytes.byteLength
    entries.push([path, await sha(bytes)])
  }
  if (!entries.length) throw new Error('invalid_tool_input')
  return sha(new TextEncoder().encode(JSON.stringify(entries)))
}
