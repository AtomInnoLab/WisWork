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

type ZipEntryWithMetadata = JSZip.JSZipObject & {
  unsafeOriginalName?: string
  _data?: { compressedSize?: number; uncompressedSize?: number }
}

async function openArchive(bytes: Uint8Array, limits: ConversionLimits): Promise<JSZip> {
  if (bytes.byteLength > limits.maxInputBytes) fail('conversion_limit')
  let archive: JSZip
  try {
    // CRC verification in JSZip eagerly inflates every entry. Metadata is validated first below;
    // only the exact allowlisted document entries are decompressed afterwards inside the worker.
    archive = await JSZip.loadAsync(bytes, { checkCRC32: false, createFolders: false })
  } catch {
    fail('conversion_invalid_document')
  }
  const entries = Object.values(archive.files) as ZipEntryWithMetadata[]
  if (entries.length > limits.maxEntries) fail('conversion_limit')
  let aggregate = 0
  for (const entry of entries) {
    const original = entry.unsafeOriginalName ?? entry.name
    if (
      original.startsWith('/') ||
      original.includes('\\') ||
      original.includes('\0') ||
      original.split('/').some((part) => part === '..' || part === '.')
    )
      fail('conversion_archive_unsafe')
    const compressed = entry._data?.compressedSize
    const uncompressed = entry._data?.uncompressedSize
    if (!entry.dir && (!Number.isSafeInteger(compressed) || !Number.isSafeInteger(uncompressed)))
      fail('conversion_archive_unsafe')
    if (entry.dir) continue
    if (uncompressed! > limits.maxEntryBytes) fail('conversion_limit')
    aggregate += uncompressed!
    if (aggregate > limits.maxArchiveBytes) fail('conversion_limit')
    if (
      uncompressed! > 0 &&
      (compressed! <= 0 || uncompressed! / compressed! > limits.maxCompressionRatio)
    )
      fail('conversion_archive_unsafe')
  }
  return archive
}

async function xmlEntry(archive: JSZip, path: string, limits: ConversionLimits): Promise<string> {
  const entry = archive.file(path)
  if (!entry) fail('conversion_invalid_document')
  const text = await entry.async('string')
  if (
    encoder.encode(text).byteLength > limits.maxEntryBytes ||
    XMLValidator.validate(text) !== true
  )
    fail('conversion_invalid_document')
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

async function docxToText(
  input: ConversionInput,
  limits: ConversionLimits,
): Promise<ConversionOutput[]> {
  const archive = await openArchive(input.bytes, limits)
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
  const archive = await openArchive(input.bytes, limits)
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
  if (archive.file('xl/sharedStrings.xml')) {
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
  if (input.bytes.byteLength > limits.maxInputBytes) fail('conversion_limit')
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
