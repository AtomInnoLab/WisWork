import { XMLParser, XMLValidator } from 'fast-xml-parser'
import JSZip from 'jszip'

export type ConversionKind = 'pdf-to-text' | 'pdf-to-images' | 'docx-to-text' | 'xlsx-to-csv'

export interface ConversionLimits {
  maxInputBytes: number
  maxEntries: number
  maxEntryBytes: number
  maxArchiveBytes: number
  maxCompressionRatio: number
  maxPages: number
  maxSheets: number
  maxRows: number
  maxColumns: number
  maxCells: number
  maxPagePixels: number
  maxTotalPixels: number
  maxOutputBytes: number
}

export const DEFAULT_CONVERSION_LIMITS: Readonly<ConversionLimits> = Object.freeze({
  maxInputBytes: 8 * 1024 * 1024,
  maxEntries: 256,
  maxEntryBytes: 8 * 1024 * 1024,
  maxArchiveBytes: 32 * 1024 * 1024,
  maxCompressionRatio: 100,
  maxPages: 100,
  maxSheets: 32,
  maxRows: 20_000,
  maxColumns: 512,
  maxCells: 200_000,
  maxPagePixels: 16_000_000,
  maxTotalPixels: 64_000_000,
  maxOutputBytes: 2 * 1024 * 1024,
})

export interface ConversionInput {
  kind: ConversionKind
  inputName: string
  bytes: Uint8Array
}
export interface ConversionOutput {
  path: string
  bytes: Uint8Array
}

interface PdfPage {
  getViewport(options: { scale: number }): { width: number; height: number }
  getTextContent?(): Promise<{ items: Array<{ str?: unknown }> }>
}
interface PdfDocument {
  numPages: number
  getPage(page: number): Promise<PdfPage>
  destroy(): void | Promise<void>
}
export interface ConversionDependencies {
  loadPdf?(bytes: Uint8Array): Promise<PdfDocument>
  renderPdfPage?(page: PdfPage, width: number, height: number): Promise<Uint8Array>
}

const encoder = new TextEncoder()
const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseAttributeValue: false,
  parseTagValue: false,
})

function fail(code: string): never {
  throw new Error(code)
}

function safeBaseName(name: string): string {
  const last = name.replaceAll('\\', '/').split('/').pop() ?? ''
  const base = last
    .replace(/\.[^.]+$/, '')
    .replace(/[^A-Za-z0-9._-]+/g, '-')
    .slice(0, 80)
  return base || 'converted'
}

function boundedOutput(outputs: ConversionOutput[], limits: ConversionLimits): ConversionOutput[] {
  const total = outputs.reduce((sum, output) => sum + output.bytes.byteLength, 0)
  if (
    total > limits.maxOutputBytes ||
    outputs.some((output) => output.bytes.byteLength > limits.maxOutputBytes)
  )
    fail('conversion_limit')
  return outputs
}

interface ArchiveContext {
  zip: JSZip
  inflatedBytes: number
  cache: Map<string, Uint8Array>
  metadata: Map<string, { crc: number; uncompressed: number; directory: boolean }>
}

interface ZipStreamObject extends JSZip.JSZipObject {
  internalStream(type: 'uint8array'): JSZip.JSZipStreamHelper<Uint8Array>
}

const ZIP_LOCAL = 0x04034b50
const ZIP_CENTRAL = 0x02014b50
const ZIP_EOCD = 0x06054b50
const zipNameDecoder = new TextDecoder('utf-8', { fatal: true })

function u16(view: DataView, offset: number): number {
  if (offset < 0 || offset + 2 > view.byteLength) fail('conversion_archive_unsafe')
  return view.getUint16(offset, true)
}

function u32(view: DataView, offset: number): number {
  if (offset < 0 || offset + 4 > view.byteLength) fail('conversion_archive_unsafe')
  return view.getUint32(offset, true)
}

function decodeZipName(bytes: Uint8Array, utf8: boolean): string {
  if (!utf8 && bytes.some((byte) => byte > 0x7f)) fail('conversion_archive_unsafe')
  try {
    return zipNameDecoder.decode(bytes)
  } catch {
    fail('conversion_archive_unsafe')
  }
}

function validateZipPath(name: string): string {
  if (
    !name ||
    name.length > 512 ||
    name.includes('\0') ||
    name.includes('\\') ||
    name.startsWith('/') ||
    /^[A-Za-z]:/.test(name)
  )
    fail('conversion_archive_unsafe')
  const directory = name.endsWith('/')
  const parts = name.split('/')
  if (directory) parts.pop()
  if (!parts.length || parts.some((part) => !part || part === '.' || part === '..'))
    fail('conversion_archive_unsafe')
  const normalized = `${parts.join('/')}${directory ? '/' : ''}`.normalize('NFC')
  if (normalized !== name) fail('conversion_archive_unsafe')
  return parts.join('/').toLocaleLowerCase('en-US')
}

function validateExtraFields(bytes: Uint8Array, start: number, length: number): void {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const end = start + length
  if (end > bytes.byteLength) fail('conversion_archive_unsafe')
  const ids = new Set<number>()
  let offset = start
  while (offset < end) {
    if (offset + 4 > end) fail('conversion_archive_unsafe')
    const id = u16(view, offset)
    const size = u16(view, offset + 2)
    offset += 4
    if (offset + size > end || ids.has(id) || id === 0x0001 || id === 0x9901)
      fail('conversion_archive_unsafe')
    ids.add(id)
    offset += size
  }
}

function validateArchiveMetadata(
  bytes: Uint8Array,
  limits: ConversionLimits,
): Map<string, { crc: number; uncompressed: number; directory: boolean }> {
  if (bytes.byteLength < 22 || bytes.byteLength > limits.maxInputBytes) fail('conversion_limit')
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const searchStart = Math.max(0, bytes.byteLength - 65_557)
  let eocd = -1
  for (let offset = bytes.byteLength - 22; offset >= searchStart; offset -= 1) {
    if (u32(view, offset) === ZIP_EOCD) {
      eocd = offset
      break
    }
  }
  if (eocd < 0) fail('conversion_archive_unsafe')
  const commentLength = u16(view, eocd + 20)
  if (eocd + 22 + commentLength !== bytes.byteLength) fail('conversion_archive_unsafe')
  const disk = u16(view, eocd + 4)
  const centralDisk = u16(view, eocd + 6)
  const diskEntries = u16(view, eocd + 8)
  const entries = u16(view, eocd + 10)
  const centralSize = u32(view, eocd + 12)
  const centralOffset = u32(view, eocd + 16)
  if (
    disk !== 0 ||
    centralDisk !== 0 ||
    diskEntries !== entries ||
    entries === 0xffff ||
    entries > limits.maxEntries ||
    centralSize === 0xffffffff ||
    centralOffset === 0xffffffff ||
    centralOffset + centralSize !== eocd
  )
    fail(entries > limits.maxEntries ? 'conversion_limit' : 'conversion_archive_unsafe')

  const names = new Set<string>()
  const metadata = new Map<string, { crc: number; uncompressed: number; directory: boolean }>()
  const localRanges: Array<readonly [number, number]> = []
  let claimedAggregate = 0
  let offset = centralOffset
  for (let index = 0; index < entries; index += 1) {
    if (u32(view, offset) !== ZIP_CENTRAL) fail('conversion_archive_unsafe')
    const madeBy = bytes[offset + 5]
    const neededVersion = u16(view, offset + 6)
    const flags = u16(view, offset + 8)
    const method = u16(view, offset + 10)
    const crc = u32(view, offset + 16)
    const compressed = u32(view, offset + 20)
    const uncompressed = u32(view, offset + 24)
    const nameLength = u16(view, offset + 28)
    const extraLength = u16(view, offset + 30)
    const entryCommentLength = u16(view, offset + 32)
    const startDisk = u16(view, offset + 34)
    const external = u32(view, offset + 38)
    const localOffset = u32(view, offset + 42)
    const end = offset + 46 + nameLength + extraLength + entryCommentLength
    if (
      end > eocd ||
      startDisk !== 0 ||
      compressed === 0xffffffff ||
      uncompressed === 0xffffffff ||
      localOffset === 0xffffffff ||
      (flags & 0x0001) !== 0 ||
      (flags & ~0x080e) !== 0 ||
      (method === 0 && (flags & 0x0006) !== 0) ||
      neededVersion > 20 ||
      ![0, 8].includes(method)
    )
      fail('conversion_archive_unsafe')
    const name = decodeZipName(
      bytes.subarray(offset + 46, offset + 46 + nameLength),
      Boolean(flags & 0x0800),
    )
    const canonical = validateZipPath(name)
    if (names.has(canonical)) fail('conversion_archive_unsafe')
    names.add(canonical)
    const directory = name.endsWith('/')
    metadata.set(name, { crc, uncompressed, directory })
    validateExtraFields(bytes, offset + 46 + nameLength, extraLength)
    if (madeBy === 3) {
      const unixType = (external >>> 16) & 0xf000
      const expected = directory ? 0x4000 : 0x8000
      if (unixType !== 0 && unixType !== expected) fail('conversion_archive_unsafe')
    } else {
      if ((external & 0x08) !== 0 || Boolean(external & 0x10) !== directory)
        fail('conversion_archive_unsafe')
    }
    if (directory && (compressed !== 0 || uncompressed !== 0 || method !== 0))
      fail('conversion_archive_unsafe')
    if (uncompressed > limits.maxEntryBytes) fail('conversion_limit')
    claimedAggregate += uncompressed
    if (claimedAggregate > limits.maxArchiveBytes) fail('conversion_limit')
    if (
      uncompressed > 0 &&
      (compressed === 0 || uncompressed / compressed > limits.maxCompressionRatio)
    )
      fail('conversion_archive_unsafe')

    if (u32(view, localOffset) !== ZIP_LOCAL) fail('conversion_archive_unsafe')
    const localFlags = u16(view, localOffset + 6)
    const localMethod = u16(view, localOffset + 8)
    const localCrc = u32(view, localOffset + 14)
    const localCompressed = u32(view, localOffset + 18)
    const localUncompressed = u32(view, localOffset + 22)
    const localNameLength = u16(view, localOffset + 26)
    const localExtraLength = u16(view, localOffset + 28)
    const localData = localOffset + 30 + localNameLength + localExtraLength
    validateExtraFields(bytes, localOffset + 30 + localNameLength, localExtraLength)
    const localName = decodeZipName(
      bytes.subarray(localOffset + 30, localOffset + 30 + localNameLength),
      Boolean(localFlags & 0x0800),
    )
    if (
      localData + compressed > centralOffset ||
      localFlags !== flags ||
      localMethod !== method ||
      localName !== name
    )
      fail('conversion_archive_unsafe')
    let localEnd = localData + compressed
    if ((flags & 0x0008) === 0) {
      if (localCrc !== crc || localCompressed !== compressed || localUncompressed !== uncompressed)
        fail('conversion_archive_unsafe')
    } else {
      if (
        ![0, crc].includes(localCrc) ||
        ![0, compressed].includes(localCompressed) ||
        ![0, uncompressed].includes(localUncompressed)
      )
        fail('conversion_archive_unsafe')
      if (u32(view, localEnd) === 0x08074b50) localEnd += 4
      if (
        u32(view, localEnd) !== crc ||
        u32(view, localEnd + 4) !== compressed ||
        u32(view, localEnd + 8) !== uncompressed
      )
        fail('conversion_archive_unsafe')
      localEnd += 12
    }
    if (localEnd > centralOffset) fail('conversion_archive_unsafe')
    localRanges.push([localOffset, localEnd])
    offset = end
  }
  if (offset !== eocd) fail('conversion_archive_unsafe')
  localRanges.sort((a, b) => a[0] - b[0])
  if (localRanges.length && localRanges[0][0] !== 0) fail('conversion_archive_unsafe')
  for (let index = 1; index < localRanges.length; index += 1)
    if (localRanges[index][0] !== localRanges[index - 1][1]) fail('conversion_archive_unsafe')
  if (localRanges.length && localRanges.at(-1)?.[1] !== centralOffset)
    fail('conversion_archive_unsafe')
  return metadata
}

async function openArchive(bytes: Uint8Array, limits: ConversionLimits): Promise<ArchiveContext> {
  const metadata = validateArchiveMetadata(bytes, limits)
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(bytes, { checkCRC32: false, createFolders: false })
  } catch {
    fail('conversion_invalid_document')
  }
  if (
    Object.keys(zip.files).length !== metadata.size ||
    [...metadata.keys()].some((name) => !Object.hasOwn(zip.files, name))
  )
    fail('conversion_archive_unsafe')
  return { zip, inflatedBytes: 0, cache: new Map(), metadata }
}

const CRC_TABLE = Uint32Array.from({ length: 256 }, (_, value) => {
  let crc = value
  for (let bit = 0; bit < 8; bit += 1) crc = (crc & 1) !== 0 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1
  return crc >>> 0
})

function updateCrc32(crc: number, bytes: Uint8Array): number {
  let value = crc
  for (const byte of bytes) value = CRC_TABLE[(value ^ byte) & 0xff] ^ (value >>> 8)
  return value >>> 0
}

async function boundedEntry(
  archive: ArchiveContext,
  path: string,
  limits: ConversionLimits,
): Promise<Uint8Array> {
  const cached = archive.cache.get(path)
  if (cached) return cached
  const entry = archive.zip.file(path) as ZipStreamObject | null
  const metadata = archive.metadata.get(path)
  if (!entry || !metadata || metadata.directory) fail('conversion_invalid_document')
  const chunks: Uint8Array[] = []
  let entryBytes = 0
  let crc = 0xffffffff
  const bytes = await new Promise<Uint8Array>((resolve, reject) => {
    const stream = entry.internalStream('uint8array')
    let settled = false
    const rejectOnce = (code: string) => {
      if (settled) return
      settled = true
      stream.pause()
      reject(new Error(code))
    }
    stream.on('data', (chunk) => {
      if (settled) return
      if (
        entryBytes + chunk.byteLength > limits.maxEntryBytes ||
        archive.inflatedBytes + chunk.byteLength > limits.maxArchiveBytes
      )
        return rejectOnce('conversion_limit')
      const copy = chunk.slice()
      chunks.push(copy)
      entryBytes += copy.byteLength
      archive.inflatedBytes += copy.byteLength
      crc = updateCrc32(crc, copy)
    })
    stream.on('error', () => rejectOnce('conversion_invalid_document'))
    stream.on('end', () => {
      if (settled) return
      settled = true
      if (entryBytes !== metadata.uncompressed || (crc ^ 0xffffffff) >>> 0 !== metadata.crc)
        return reject(new Error('conversion_invalid_document'))
      const output = new Uint8Array(entryBytes)
      let offset = 0
      for (const chunk of chunks) {
        output.set(chunk, offset)
        offset += chunk.byteLength
      }
      resolve(output)
    })
    stream.resume()
  })
  archive.cache.set(path, bytes)
  return bytes
}

async function xmlEntry(
  archive: ArchiveContext,
  path: string,
  limits: ConversionLimits,
): Promise<string> {
  let text: string
  try {
    text = new TextDecoder('utf-8', { fatal: true }).decode(
      await boundedEntry(archive, path, limits),
    )
  } catch (error) {
    if (error instanceof Error && error.message === 'conversion_limit') throw error
    fail('conversion_invalid_document')
  }
  if (XMLValidator.validate(text) !== true) fail('conversion_invalid_document')
  return text
}

function xmlText(source: string): string {
  return source
    .replace(/<w:tab\b[^>]*\/>/g, '\t')
    .replace(/<w:br\b[^>]*\/>/g, '\n')
    .replace(/<\/w:p>/g, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
}

async function requireContentType(
  archive: ArchiveContext,
  limits: ConversionLimits,
  partName: string,
  contentType: string,
): Promise<void> {
  const source = await xmlEntry(archive, '[Content_Types].xml', limits)
  let contentTypes: Record<string, any>
  try {
    contentTypes = parser.parse(source)
  } catch {
    fail('conversion_invalid_document')
  }
  const overrides = asArray(contentTypes.Types?.Override)
  if (
    !overrides.some(
      (entry) => entry?.['@_PartName'] === partName && entry?.['@_ContentType'] === contentType,
    )
  )
    fail('conversion_invalid_document')
}

async function docxToText(
  input: ConversionInput,
  limits: ConversionLimits,
): Promise<ConversionOutput[]> {
  if (!/\.docx$/i.test(input.inputName)) fail('conversion_invalid_document')
  const archive = await openArchive(input.bytes, limits)
  await requireContentType(
    archive,
    limits,
    '/word/document.xml',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  )
  const document = await xmlEntry(archive, 'word/document.xml', limits)
  return boundedOutput(
    [{ path: `${safeBaseName(input.inputName)}.txt`, bytes: encoder.encode(xmlText(document)) }],
    limits,
  )
}

function asArray<T>(value: T | T[] | undefined): T[] {
  return value === undefined ? [] : Array.isArray(value) ? value : [value]
}

function columnIndex(reference: string): { row: number; column: number } {
  const match = /^([A-Z]{1,4})([1-9]\d{0,6})$/.exec(reference)
  if (!match) fail('conversion_invalid_document')
  let column = 0
  for (const character of match[1]) column = column * 26 + character.charCodeAt(0) - 64
  return { row: Number(match[2]), column }
}

function csvEscape(value: string): string {
  const safe = /^[=+\-@\t\r]/.test(value) ? `'${value}` : value
  return /[",\r\n]/.test(safe) ? `"${safe.replaceAll('"', '""')}"` : safe
}

async function xlsxToCsv(
  input: ConversionInput,
  limits: ConversionLimits,
): Promise<ConversionOutput[]> {
  if (!/\.xlsx$/i.test(input.inputName)) fail('conversion_invalid_document')
  const archive = await openArchive(input.bytes, limits)
  await requireContentType(
    archive,
    limits,
    '/xl/workbook.xml',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  )
  const workbookXml = await xmlEntry(archive, 'xl/workbook.xml', limits)
  const relationshipsXml = await xmlEntry(archive, 'xl/_rels/workbook.xml.rels', limits)
  let workbook: Record<string, any>
  let relationships: Record<string, any>
  try {
    workbook = parser.parse(workbookXml)
    relationships = parser.parse(relationshipsXml)
  } catch {
    fail('conversion_invalid_document')
  }
  const sheets = asArray(workbook.workbook?.sheets?.sheet)
  if (!sheets.length || sheets.length > limits.maxSheets) fail('conversion_limit')
  const relationMap = new Map(
    asArray(relationships.Relationships?.Relationship).map((item) => [
      item['@_Id'],
      item['@_Target'],
    ]),
  )
  let shared: string[] = []
  if (archive.zip.file('xl/sharedStrings.xml')) {
    const sharedXml = await xmlEntry(archive, 'xl/sharedStrings.xml', limits)
    try {
      const parsed = parser.parse(sharedXml)
      shared = asArray(parsed.sst?.si).map((item) => {
        const values: unknown[] = []
        const visit = (node: unknown): void => {
          if (Array.isArray(node)) {
            node.forEach(visit)
          } else if (node && typeof node === 'object') {
            for (const [key, value] of Object.entries(node)) {
              if (key === 't') values.push(value)
              else visit(value)
            }
          }
        }
        visit(item)
        return values.join('')
      })
    } catch {
      fail('conversion_invalid_document')
    }
  }
  const outputs: ConversionOutput[] = []
  let totalCells = 0
  const usedNames = new Set<string>()
  for (const sheet of sheets) {
    const target = relationMap.get(sheet['@_r:id'])
    if (typeof target !== 'string' || target.startsWith('/') || target.includes('..'))
      fail('conversion_archive_unsafe')
    const path = `xl/${target.replace(/^\.\//, '')}`
    const sheetXml = await xmlEntry(archive, path, limits)
    let parsed: Record<string, any>
    try {
      parsed = parser.parse(sheetXml)
    } catch {
      fail('conversion_invalid_document')
    }
    const cells = asArray(parsed.worksheet?.sheetData?.row).flatMap((row) => asArray(row.c))
    totalCells += cells.length
    if (totalCells > limits.maxCells) fail('conversion_limit')
    const rows = new Map<number, Map<number, string>>()
    let maxRow = 0
    let maxColumn = 0
    for (const cell of cells) {
      const coordinate = columnIndex(cell['@_r'])
      if (coordinate.row > limits.maxRows || coordinate.column > limits.maxColumns)
        fail('conversion_limit')
      maxRow = Math.max(maxRow, coordinate.row)
      maxColumn = Math.max(maxColumn, coordinate.column)
      if (maxRow * maxColumn > limits.maxCells) fail('conversion_limit')
      const raw = cell.v === undefined ? '' : String(cell.v)
      const value = cell['@_t'] === 's' ? shared[Number(raw)] : raw
      if (value === undefined) fail('conversion_invalid_document')
      const row = rows.get(coordinate.row) ?? new Map<number, string>()
      row.set(coordinate.column, value)
      rows.set(coordinate.row, row)
    }
    const lines: string[] = []
    for (let row = 1; row <= maxRow; row += 1) {
      const values: string[] = []
      for (let column = 1; column <= maxColumn; column += 1)
        values.push(csvEscape(rows.get(row)?.get(column) ?? ''))
      lines.push(values.join(','))
    }
    const name =
      String(sheet['@_name'] ?? 'Sheet')
        .replace(/[^A-Za-z0-9._-]+/g, '-')
        .slice(0, 40) || 'Sheet'
    if (usedNames.has(name.toLowerCase())) fail('conversion_invalid_document')
    usedNames.add(name.toLowerCase())
    outputs.push({
      path: `${safeBaseName(input.inputName)}-${name}.csv`,
      bytes: encoder.encode(lines.join('\n')),
    })
  }
  return boundedOutput(outputs, limits)
}

async function defaultLoadPdf(bytes: Uint8Array): Promise<PdfDocument> {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const options = {
    data: bytes.slice(),
    isEvalSupported: false,
    useWorkerFetch: false,
    disableFontFace: true,
  }
  return pdfjs.getDocument(options).promise as Promise<PdfDocument>
}

async function defaultRenderPdfPage(
  page: PdfPage,
  width: number,
  height: number,
): Promise<Uint8Array> {
  if (typeof OffscreenCanvas === 'undefined') fail('conversion_unsupported')
  const canvas = new OffscreenCanvas(width, height)
  const context = canvas.getContext('2d')
  if (!context || !('render' in page)) fail('conversion_unsupported')
  await (page as PdfPage & { render(options: object): { promise: Promise<void> } }).render({
    canvasContext: context,
    viewport: page.getViewport({ scale: 1 }),
  }).promise
  const blob = await canvas.convertToBlob({ type: 'image/png' })
  return new Uint8Array(await blob.arrayBuffer())
}

async function pdfConvert(
  input: ConversionInput,
  limits: ConversionLimits,
  dependencies: ConversionDependencies,
): Promise<ConversionOutput[]> {
  if (!/\.pdf$/i.test(input.inputName) || input.bytes.byteLength < 5)
    fail('conversion_invalid_document')
  if (input.bytes.byteLength > limits.maxInputBytes) fail('conversion_limit')
  if (new TextDecoder().decode(input.bytes.subarray(0, 5)) !== '%PDF-')
    fail('conversion_invalid_document')
  let pdf: PdfDocument
  try {
    pdf = await (dependencies.loadPdf ?? defaultLoadPdf)(input.bytes)
  } catch {
    fail('conversion_invalid_document')
  }
  try {
    if (!Number.isInteger(pdf.numPages) || pdf.numPages < 1 || pdf.numPages > limits.maxPages)
      fail('conversion_limit')
    if (input.kind === 'pdf-to-text') {
      const pages: string[] = []
      for (let index = 1; index <= pdf.numPages; index += 1) {
        const content = await (await pdf.getPage(index)).getTextContent?.()
        if (!content) fail('conversion_invalid_document')
        pages.push(
          content.items.map((item) => (typeof item.str === 'string' ? item.str : '')).join(' '),
        )
        if (encoder.encode(pages.join('\n\n')).byteLength > limits.maxOutputBytes)
          fail('conversion_limit')
      }
      return [
        { path: `${safeBaseName(input.inputName)}.txt`, bytes: encoder.encode(pages.join('\n\n')) },
      ]
    }
    const outputs: ConversionOutput[] = []
    let totalPixels = 0
    for (let index = 1; index <= pdf.numPages; index += 1) {
      const page = await pdf.getPage(index)
      const viewport = page.getViewport({ scale: 1 })
      const width = Math.ceil(viewport.width)
      const height = Math.ceil(viewport.height)
      const pixels = width * height
      if (!Number.isSafeInteger(pixels) || pixels <= 0 || pixels > limits.maxPagePixels)
        fail('conversion_limit')
      totalPixels += pixels
      if (totalPixels > limits.maxTotalPixels) fail('conversion_limit')
      const bytes = await (dependencies.renderPdfPage ?? defaultRenderPdfPage)(page, width, height)
      outputs.push({ path: `${safeBaseName(input.inputName)}-page-${index}.png`, bytes })
      boundedOutput(outputs, limits)
    }
    return outputs
  } finally {
    await pdf.destroy()
  }
}

export async function convertDocument(
  input: ConversionInput,
  limits: ConversionLimits = DEFAULT_CONVERSION_LIMITS,
  dependencies: ConversionDependencies = {},
): Promise<ConversionOutput[]> {
  if (
    !(input.bytes instanceof Uint8Array) ||
    !Object.values(limits).every((value) => Number.isFinite(value) && value > 0) ||
    !Object.entries(limits).every(([name, value]) =>
      name === 'maxCompressionRatio' ? true : Number.isInteger(value),
    )
  )
    fail('conversion_invalid_document')
  if (input.kind === 'docx-to-text') return docxToText(input, limits)
  if (input.kind === 'xlsx-to-csv') return xlsxToCsv(input, limits)
  if (input.kind === 'pdf-to-text' || input.kind === 'pdf-to-images')
    return pdfConvert(input, limits, dependencies)
  fail('conversion_invalid_document')
}
