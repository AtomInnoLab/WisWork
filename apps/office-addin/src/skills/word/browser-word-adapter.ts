import type { WordDocumentBlock, WordDocumentWrite, WordInlineSpan } from './word-markdown.js'
import { readUntilConverged } from '../shared/office-write-transaction.js'

export const MAX_WORD_PARAGRAPHS = 500
export const MAX_WORD_TEXT_LENGTH = 12_000
export const MAX_WORD_OOXML_BYTES = 1024 * 1024
export const MAX_WORD_RESULT_BYTES = 256 * 1024
export const MAX_WORD_PDF_BYTES = 16 * 1024 * 1024
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml'
const PKG_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage'
const XMLNS_NS = 'http://www.w3.org/2000/xmlns/'

export interface WordParagraph {
  index: number
  text: string
  style?: string
  alignment?: string
  listLevel?: number
  listString?: string
}

export interface WordDocumentTextResult {
  totalParagraphs: number
  totalParagraphsExact: boolean
  hasMore: boolean
  showing: { start: number; end: number }
  paragraphs: WordParagraph[]
}

export interface WordDocumentStructureResult {
  paragraphCount: number
  sectionCount: number
  tableCount: number
  contentControlCount: number
  truncated: {
    paragraphs: boolean
    sections: boolean
    tables: boolean
    contentControls: boolean
  }
  headings: Array<{ text: string; level: number; paragraphIndex: number }>
  tables: Array<{ index: number; rows: number; rowsTruncated: boolean; style: string }>
  contentControls: Array<{ id: number; title: string; tag: string; type: string }>
}

export interface WordOoxmlChild {
  index: number
  type: string
  line: number
  paragraphIndex?: number
  tableIndex?: number
  paragraphRange?: [number, number]
  rows?: number
  cols?: number
  text?: string
}

export interface WordOoxmlResult {
  xml: string
  children: WordOoxmlChild[]
}

export interface WordScreenshotResult {
  base64: string
  mime: 'image/png'
}

export interface WordDocumentSnapshot {
  text: string
  fingerprint: string
}

export type WordDeclarativeOperation =
  | { op: 'insert_text'; location: 'start' | 'end' | 'replace'; text: string }
  | { op: 'replace_all'; search: string; replacement: string; matchCase: boolean }

export interface WordAdapter {
  getDocumentSnapshot(signal?: AbortSignal): Promise<WordDocumentSnapshot>
  getDocumentText(
    options: { startParagraph?: number; endParagraph?: number; includeFormatting?: boolean },
    signal?: AbortSignal,
  ): Promise<WordDocumentTextResult>
  getDocumentStructure(signal?: AbortSignal): Promise<WordDocumentStructureResult>
  getOoxml(
    options: { startChild?: number; endChild?: number },
    signal?: AbortSignal,
  ): Promise<WordOoxmlResult>
  screenshotDocument(page: number, signal?: AbortSignal): Promise<WordScreenshotResult>
  fingerprint(signal?: AbortSignal): Promise<string>
  executeOperations(operations: WordDeclarativeOperation[], signal?: AbortSignal): Promise<void>
  verifyOperations(operations: WordDeclarativeOperation[], signal?: AbortSignal): Promise<boolean>
  executeDocumentWrite(write: WordDocumentWrite, signal?: AbortSignal): Promise<void>
  verifyDocumentWrite(write: WordDocumentWrite, signal?: AbortSignal): Promise<boolean>
}

type RuntimeRecord = Record<string, unknown>

function cancelled(signal?: AbortSignal): void {
  if (signal?.aborted) throw new Error('cancelled')
}

function runtime(): { office: RuntimeRecord; word: RuntimeRecord } {
  const root = globalThis as unknown as RuntimeRecord
  const office = root.Office as RuntimeRecord | undefined
  const word = root.Word as RuntimeRecord | undefined
  const context = office?.context as RuntimeRecord | undefined
  const requirements = context?.requirements as RuntimeRecord | undefined
  const supports = requirements?.isSetSupported
  if (
    !office ||
    !word ||
    context?.host !== 'Word' ||
    typeof supports !== 'function' ||
    !(supports as (name: string, version: string) => boolean).call(
      requirements,
      'WordApi',
      '1.3',
    ) ||
    typeof word.run !== 'function'
  ) {
    throw new Error('office_api_unsupported')
  }
  return { office, word }
}

function text(value: unknown, maximum = MAX_WORD_TEXT_LENGTH): string {
  return typeof value === 'string' ? value.slice(0, maximum) : ''
}

function finiteInteger(value: unknown, fallback = 0): number {
  return Number.isFinite(value) ? Math.trunc(value as number) : fallback
}

async function sync(context: RuntimeRecord, signal?: AbortSignal): Promise<void> {
  cancelled(signal)
  await (context.sync as () => Promise<void>)()
  cancelled(signal)
}

function elements(parent: Element): Element[] {
  return Array.from(parent.childNodes).filter((node): node is Element => node.nodeType === 1)
}

async function exportWordPdf(signal?: AbortSignal): Promise<Uint8Array> {
  cancelled(signal)
  const root = globalThis as unknown as RuntimeRecord
  const office = root.Office as RuntimeRecord | undefined
  const context = office?.context as RuntimeRecord | undefined
  const document = context?.document as RuntimeRecord | undefined
  const fileType = office?.FileType as RuntimeRecord | undefined
  if (typeof document?.getFileAsync !== 'function' || fileType?.Pdf === undefined)
    throw new Error('office_api_unsupported')
  const file = await new Promise<RuntimeRecord>((resolve, reject) => {
    ;(
      document.getFileAsync as (
        type: unknown,
        options: object,
        callback: (result: RuntimeRecord) => void,
      ) => void
    )(fileType.Pdf, { sliceSize: 64 * 1024 }, (result) => {
      if (result.status !== 'succeeded' || !result.value) reject(new Error('office_read_failed'))
      else resolve(result.value as RuntimeRecord)
    })
  })
  try {
    const size = finiteInteger(file.size, -1)
    const sliceCount = finiteInteger(file.sliceCount, -1)
    if (size < 1 || size > MAX_WORD_PDF_BYTES || sliceCount < 1 || sliceCount > 256)
      throw new Error('office_read_failed')
    const chunks: Uint8Array[] = []
    let total = 0
    for (let index = 0; index < sliceCount; index += 1) {
      cancelled(signal)
      const slice = await new Promise<RuntimeRecord>((resolve, reject) => {
        ;(file.getSliceAsync as (index: number, callback: (result: RuntimeRecord) => void) => void)(
          index,
          (result) => {
            if (result.status !== 'succeeded' || !result.value)
              reject(new Error('office_read_failed'))
            else resolve(result.value as RuntimeRecord)
          },
        )
      })
      cancelled(signal)
      const data = slice.data
      if (
        !Array.isArray(data) ||
        data.some((value) => !Number.isInteger(value) || value < 0 || value > 255)
      )
        throw new Error('office_read_failed')
      const bytes = Uint8Array.from(data as number[])
      total += bytes.byteLength
      if (total > MAX_WORD_PDF_BYTES) throw new Error('office_read_failed')
      chunks.push(bytes)
    }
    if (total !== size) throw new Error('office_read_failed')
    const result = new Uint8Array(total)
    let offset = 0
    for (const chunk of chunks) {
      result.set(chunk, offset)
      offset += chunk.byteLength
    }
    return result
  } finally {
    if (typeof file.closeAsync === 'function')
      await new Promise<void>((resolve) =>
        (file.closeAsync as (callback: () => void) => void)(() => resolve()),
      )
  }
}

function summarizeOoxml(
  xml: string,
  options: { startChild?: number; endChild?: number },
): WordOoxmlResult {
  if (new TextEncoder().encode(xml).byteLength > MAX_WORD_OOXML_BYTES)
    throw new Error('office_read_failed')
  if (typeof DOMParser === 'undefined' || typeof XMLSerializer === 'undefined')
    throw new Error('office_api_unsupported')
  const document = new DOMParser().parseFromString(xml, 'text/xml')
  if (document.getElementsByTagName('parsererror').length) throw new Error('office_read_failed')
  const body = Array.from(document.getElementsByTagName('*')).find(
    (item) => item.localName === 'body',
  )
  if (!body) throw new Error('office_read_failed')
  const all = elements(body)
  const start = options.startChild ?? 0
  const end = options.endChild ?? all.length - 1
  if (all.length === 0 || start < 0 || start >= all.length || end < start || end >= all.length)
    throw new Error('invalid_tool_input')

  let paragraphIndex = 0
  let tableIndex = 0
  let line = 2
  const blocks: string[] = []
  const children: WordOoxmlChild[] = []
  for (let index = 0; index < all.length; index += 1) {
    const element = all[index]
    const type = element.localName
    const paragraphs =
      type === 'p'
        ? 1
        : Array.from(element.getElementsByTagName('*')).filter((item) => item.localName === 'p')
            .length
    if (index >= start && index <= end) {
      const serialized = new XMLSerializer().serializeToString(element)
      const summary: WordOoxmlChild = { index, type: text(type, 32), line }
      if (type === 'p') {
        summary.paragraphIndex = paragraphIndex
        summary.text = text(element.textContent, 80)
      } else if (type === 'tbl') {
        const rows = Array.from(element.getElementsByTagName('*')).filter(
          (item) => item.localName === 'tr',
        )
        const cols = rows[0]
          ? Array.from(rows[0].getElementsByTagName('*')).filter((item) => item.localName === 'tc')
              .length
          : 0
        summary.tableIndex = tableIndex
        summary.rows = rows.length
        summary.cols = cols
        summary.paragraphRange = [paragraphIndex, paragraphIndex + Math.max(paragraphs - 1, 0)]
      } else if (paragraphs > 0) {
        summary.paragraphRange = [paragraphIndex, paragraphIndex + paragraphs - 1]
      }
      children.push(summary)
      blocks.push(serialized)
      line += serialized.split('\n').length + 1
    }
    paragraphIndex += paragraphs
    if (type === 'tbl') tableIndex += 1
  }
  return {
    xml: `<w:body xmlns:w="${W_NS}">\n${blocks.join('\n\n')}\n</w:body>`,
    children,
  }
}

async function readBodyOoxml(signal?: AbortSignal): Promise<string> {
  cancelled(signal)
  const { word } = runtime()
  return (word.run as (callback: (context: RuntimeRecord) => unknown) => Promise<unknown>)(
    async (context) => {
      const body = (context.document as RuntimeRecord).body as RuntimeRecord
      if (typeof body.getOoxml !== 'function') throw new Error('office_api_unsupported')
      const result = (body.getOoxml as () => RuntimeRecord)()
      await sync(context, signal)
      if (
        typeof result.value !== 'string' ||
        new TextEncoder().encode(result.value).byteLength > MAX_WORD_OOXML_BYTES
      )
        throw new Error('office_read_failed')
      return result.value
    },
  ) as Promise<string>
}

export class BrowserWordAdapter implements WordAdapter {
  private expectedText:
    { expected: string; fingerprint: string; originalFingerprint: string } | undefined
  private expectedDocument:
    | {
        fingerprint: string
        original: string
        originalFingerprint: string
        write: WordDocumentWrite
      }
    | undefined
  constructor(
    private readonly screenshotDependencies: {
      exportPdf?: (signal?: AbortSignal) => Promise<Uint8Array>
      renderPage?: (bytes: Uint8Array, page: number, signal?: AbortSignal) => Promise<string>
    } = {},
  ) {}

  async getDocumentText(
    options: { startParagraph?: number; endParagraph?: number; includeFormatting?: boolean },
    signal?: AbortSignal,
  ): Promise<WordDocumentTextResult> {
    cancelled(signal)
    const { word } = runtime()
    return (word.run as (callback: (context: RuntimeRecord) => unknown) => Promise<unknown>)(
      async (context) => {
        const document = context.document as RuntimeRecord
        const body = document.body as RuntimeRecord
        const paragraphs = body.paragraphs as RuntimeRecord
        const start = options.startParagraph ?? 0
        const requested =
          options.endParagraph === undefined
            ? MAX_WORD_PARAGRAPHS
            : Math.min(options.endParagraph - start, MAX_WORD_PARAGRAPHS)
        ;(paragraphs.load as (properties: unknown) => void)({
          $skip: start,
          $top: Math.max(requested + 1, 1),
        })
        await sync(context, signal)
        let loaded = paragraphs.items as RuntimeRecord[]
        if (loaded.length === 0 && start > 0) {
          ;(paragraphs.load as (properties: unknown) => void)({ $skip: start - 1, $top: 1 })
          await sync(context, signal)
          if ((paragraphs.items as RuntimeRecord[]).length === 0)
            throw new Error('invalid_tool_input')
          loaded = []
        }
        const hasMore = loaded.length > requested
        const items = loaded.slice(0, requested)
        const end = start + items.length
        const includeFormatting = options.includeFormatting !== false
        const selected = items
        for (const paragraph of selected) {
          ;(paragraph.load as (properties: string) => void)(
            includeFormatting ? 'text,style,alignment,outlineLevel' : 'text',
          )
          if (includeFormatting) {
            const list = paragraph.listItemOrNullObject as RuntimeRecord | undefined
            if (list && typeof list.load === 'function')
              (list.load as (properties: string) => void)('level,listString,isNullObject')
          }
        }
        await sync(context, signal)
        return {
          totalParagraphs: end,
          totalParagraphsExact: !hasMore,
          hasMore,
          showing: { start, end },
          paragraphs: selected.map((paragraph, offset) => {
            const result: WordParagraph = { index: start + offset, text: text(paragraph.text) }
            if (includeFormatting) {
              result.style = text(paragraph.style, 256)
              result.alignment = text(paragraph.alignment, 64)
              const list = paragraph.listItemOrNullObject as RuntimeRecord | undefined
              if (list && list.isNullObject === false) {
                result.listLevel = finiteInteger(list.level)
                result.listString = text(list.listString, 128)
              }
            }
            return result
          }),
        }
      },
    ) as Promise<WordDocumentTextResult>
  }

  async getDocumentStructure(signal?: AbortSignal): Promise<WordDocumentStructureResult> {
    cancelled(signal)
    const { word } = runtime()
    return (word.run as (callback: (context: RuntimeRecord) => unknown) => Promise<unknown>)(
      async (context) => {
        const document = context.document as RuntimeRecord
        const body = document.body as RuntimeRecord
        const paragraphs = body.paragraphs as RuntimeRecord
        const tables = body.tables as RuntimeRecord
        const contentControls = body.contentControls as RuntimeRecord
        const sections = document.sections as RuntimeRecord
        ;(paragraphs.load as (properties: unknown) => void)({ $top: MAX_WORD_PARAGRAPHS + 1 })
        ;(tables.load as (properties: unknown) => void)({ $top: 257 })
        ;(contentControls.load as (properties: unknown) => void)({ $top: 257 })
        ;(sections.load as (properties: unknown) => void)({ $top: 257 })
        await sync(context, signal)
        const paragraphItems = (paragraphs.items as RuntimeRecord[]).slice(0, MAX_WORD_PARAGRAPHS)
        const tableItems = (tables.items as RuntimeRecord[]).slice(0, 256)
        const controlItems = (contentControls.items as RuntimeRecord[]).slice(0, 256)
        for (const paragraph of paragraphItems)
          (paragraph.load as (properties: string) => void)('text,style,outlineLevel')
        for (const table of tableItems) {
          ;(table.load as (properties: string) => void)('style')
          ;((table.rows as RuntimeRecord).load as (properties: unknown) => void)({ $top: 257 })
        }
        for (const control of controlItems)
          (control.load as (properties: string) => void)('title,tag,type,id')
        await sync(context, signal)
        return {
          paragraphCount: Math.min((paragraphs.items as unknown[]).length, MAX_WORD_PARAGRAPHS),
          sectionCount: Math.min((sections.items as unknown[]).length, 256),
          tableCount: Math.min((tables.items as unknown[]).length, 256),
          contentControlCount: Math.min((contentControls.items as unknown[]).length, 256),
          truncated: {
            paragraphs: (paragraphs.items as unknown[]).length > MAX_WORD_PARAGRAPHS,
            sections: (sections.items as unknown[]).length > 256,
            tables: (tables.items as unknown[]).length > 256,
            contentControls: (contentControls.items as unknown[]).length > 256,
          },
          headings: paragraphItems.flatMap((paragraph, paragraphIndex) => {
            const level = finiteInteger(paragraph.outlineLevel)
            return level >= 1 && level <= 9
              ? [{ text: text(paragraph.text, 120), level, paragraphIndex }]
              : []
          }),
          tables: tableItems.map((table, index) => ({
            index,
            rows: Math.min((((table.rows as RuntimeRecord).items as unknown[]) ?? []).length, 256),
            rowsTruncated: (((table.rows as RuntimeRecord).items as unknown[]) ?? []).length > 256,
            style: text(table.style, 256),
          })),
          contentControls: controlItems.map((control) => ({
            id: finiteInteger(control.id),
            title: text(control.title, 256),
            tag: text(control.tag, 256),
            type: text(control.type, 64),
          })),
        }
      },
    ) as Promise<WordDocumentStructureResult>
  }

  async getOoxml(
    options: { startChild?: number; endChild?: number },
    signal?: AbortSignal,
  ): Promise<WordOoxmlResult> {
    return summarizeOoxml(await readBodyOoxml(signal), options)
  }

  async screenshotDocument(page: number, signal?: AbortSignal): Promise<WordScreenshotResult> {
    cancelled(signal)
    runtime()
    const bytes = await (this.screenshotDependencies.exportPdf?.(signal) ?? exportWordPdf(signal))
    cancelled(signal)
    const render =
      this.screenshotDependencies.renderPage ??
      (await import('../shared/browser-pdf.js')).renderPdfPageToPng
    return { base64: await render(bytes, page, signal), mime: 'image/png' }
  }

  async fingerprint(signal?: AbortSignal): Promise<string> {
    return (await this.getDocumentSnapshot(signal)).fingerprint
  }

  async getDocumentSnapshot(signal?: AbortSignal): Promise<WordDocumentSnapshot> {
    cancelled(signal)
    const value = await readBodyOoxml(signal)
    const digest = await crypto.subtle.digest(
      'SHA-256',
      new TextEncoder().encode(stableFingerprintOoxml(value)),
    )
    cancelled(signal)
    return {
      text: ooxmlText(value),
      fingerprint: [...new Uint8Array(digest)]
        .map((byte) => byte.toString(16).padStart(2, '0'))
        .join(''),
    }
  }

  async executeOperations(
    operations: WordDeclarativeOperation[],
    signal?: AbortSignal,
  ): Promise<void> {
    cancelled(signal)
    assertIndependentWordOperations(operations)
    this.expectedText = undefined
    const { word } = runtime()
    await (word.run as (callback: (context: RuntimeRecord) => unknown) => Promise<unknown>)(
      async (context) => {
        const body = (context.document as RuntimeRecord).body as RuntimeRecord
        if (
          typeof body.insertText !== 'function' ||
          typeof body.search !== 'function' ||
          typeof body.getOoxml !== 'function'
        )
          throw new Error('office_api_unsupported')
        const snapshot = (body.getOoxml as () => RuntimeRecord)()
        const searches = operations.map((operation) => {
          if (operation.op !== 'replace_all') return undefined
          const result = (body.search as (value: string, options: object) => RuntimeRecord)(
            operation.search,
            { matchCase: operation.matchCase },
          )
          ;(result.load as (properties: string) => void)('items')
          return result
        })
        await sync(context, signal)
        if (
          typeof snapshot.value !== 'string' ||
          new TextEncoder().encode(snapshot.value).byteLength > MAX_WORD_OOXML_BYTES
        )
          throw new Error('office_write_failed')
        let expected = ooxmlText(snapshot.value)
        let replacementCount = 0
        for (let index = 0; index < operations.length; index += 1) {
          const operation = operations[index]
          const items = searches[index]?.items as RuntimeRecord[] | undefined
          replacementCount += items?.length ?? 0
          if (
            replacementCount > 1_000 ||
            items?.some((item) => typeof item.insertText !== 'function')
          )
            throw new Error('office_write_failed')
          expected = applyTextOperation(expected, operation, items?.length)
        }
        if (new TextEncoder().encode(expected).byteLength > MAX_WORD_OOXML_BYTES)
          throw new Error('office_write_failed')
        cancelled(signal)
        for (let index = 0; index < operations.length; index += 1) {
          const operation = operations[index]
          if (operation.op === 'insert_text') {
            ;(body.insertText as (value: string, location: string) => void)(
              operation.text,
              { start: 'Start', end: 'End', replace: 'Replace' }[operation.location],
            )
          } else {
            for (const item of searches[index]?.items as RuntimeRecord[])
              (item.insertText as (value: string, location: string) => void)(
                operation.replacement,
                'Replace',
              )
          }
        }
        cancelled(signal)
        let writeError: unknown
        try {
          await (context.sync as () => Promise<void>)()
        } catch (error) {
          writeError = error
        }
        const current = await readUntilConverged({
          read: () => readWordBodyInContext(context, body),
          accept: (value) => ooxmlText(value) === expected,
        })
        const originalFingerprint = stableFingerprintOoxml(snapshot.value)
        if (ooxmlText(current) !== expected) {
          if (stableFingerprintOoxml(current) === originalFingerprint)
            throw new Error('office_write_failed', { cause: writeError })
          throw new Error('office_concurrent_change', { cause: writeError })
        }
        if (
          stableNonTextFingerprintOoxml(current) !== stableNonTextFingerprintOoxml(snapshot.value)
        )
          throw new Error('office_state_uncertain', { cause: writeError })
        this.expectedText = {
          expected,
          fingerprint: stableFingerprintOoxml(current),
          originalFingerprint,
        }
        if (signal?.aborted) throw new Error('office_state_uncertain')
      },
    )
  }

  async verifyOperations(
    operations: WordDeclarativeOperation[],
    signal?: AbortSignal,
  ): Promise<boolean> {
    cancelled(signal)
    void operations
    const expected = this.expectedText
    this.expectedText = undefined
    if (expected === undefined) return false
    const current = await readUntilConverged({
      read: () => readBodyOoxml(signal),
      accept: (value) => ooxmlText(value) === expected.expected,
      signal,
    })
    if (
      ooxmlText(current) === expected.expected &&
      stableFingerprintOoxml(current) === expected.fingerprint
    )
      return true
    if (stableFingerprintOoxml(current) === expected.originalFingerprint) return false
    throw new Error('office_concurrent_change')
  }

  async executeDocumentWrite(write: WordDocumentWrite, signal?: AbortSignal): Promise<void> {
    cancelled(signal)
    this.expectedDocument = undefined
    const { word } = runtime()
    await (word.run as (callback: (context: RuntimeRecord) => unknown) => Promise<unknown>)(
      async (context) => {
        const body = (context.document as RuntimeRecord).body as RuntimeRecord
        if (typeof body.insertOoxml !== 'function' || typeof body.getOoxml !== 'function')
          throw new Error('office_api_unsupported')
        const snapshot = (body.getOoxml as () => RuntimeRecord)()
        await sync(context, signal)
        if (
          typeof snapshot.value !== 'string' ||
          new TextEncoder().encode(snapshot.value).byteLength > MAX_WORD_OOXML_BYTES
        )
          throw new Error('office_write_failed')
        const original = snapshot.value
        const beforeFingerprint = stableFingerprintOoxml(original)
        cancelled(signal)
        const target = buildDocumentWriteOoxml(original, write)
        ;(body.insertOoxml as (xml: string, location: string) => void)(target, 'Replace')
        let writeError: unknown
        try {
          await (context.sync as () => Promise<void>)()
        } catch (error) {
          writeError = error
        }
        let postValue: string | undefined
        let readError: unknown
        try {
          postValue = await readUntilConverged({
            read: () => readWordBodyInContext(context, body),
            accept: (value) => verifyNativeDocumentWrite(original, value, write),
          })
        } catch (error) {
          readError = error
        }
        if (readError) throw new Error('office_state_uncertain', { cause: readError })
        const verified =
          postValue !== undefined &&
          new TextEncoder().encode(postValue).byteLength <= MAX_WORD_OOXML_BYTES &&
          verifyNativeDocumentWrite(original, postValue, write)
        if (!verified) {
          if (postValue && stableFingerprintOoxml(postValue) === beforeFingerprint)
            throw new Error('office_write_failed', { cause: writeError })
          throw new Error('office_concurrent_change', { cause: writeError })
        }
        this.expectedDocument = {
          fingerprint: stableFingerprintOoxml(postValue!),
          original,
          originalFingerprint: beforeFingerprint,
          write,
        }
        if (signal?.aborted) throw new Error('office_state_uncertain')
      },
    )
  }

  async verifyDocumentWrite(_write: WordDocumentWrite, signal?: AbortSignal): Promise<boolean> {
    const expected = this.expectedDocument
    this.expectedDocument = undefined
    if (!expected) return false
    let current: string
    try {
      current = await readUntilConverged({
        read: () => readBodyOoxml(signal),
        accept: (value) =>
          stableFingerprintOoxml(value) === expected.fingerprint ||
          verifyNativeDocumentWrite(expected.original, value, expected.write),
      })
    } catch (error) {
      throw new Error('office_state_uncertain', { cause: error })
    }
    const currentFingerprint = stableFingerprintOoxml(current)
    if (currentFingerprint === expected.fingerprint && !signal?.aborted) return true
    if (currentFingerprint === expected.originalFingerprint) return false
    if (!signal?.aborted && verifyNativeDocumentWrite(expected.original, current, expected.write))
      return true
    throw new Error('office_concurrent_change')
  }
}

function assertIndependentWordOperations(operations: WordDeclarativeOperation[]): void {
  for (let laterIndex = 0; laterIndex < operations.length; laterIndex += 1) {
    const later = operations[laterIndex]
    if (later.op !== 'replace_all') continue
    const normalize = (value: string) => (later.matchCase ? value : value.toLocaleLowerCase())
    const laterSearch = normalize(later.search)
    for (const earlier of operations.slice(0, laterIndex)) {
      const produced = normalize(earlier.op === 'insert_text' ? earlier.text : earlier.replacement)
      if (produced.includes(laterSearch)) throw new Error('invalid_tool_input')
      if (earlier.op === 'replace_all') {
        const earlierSearch = normalize(earlier.search)
        if (earlierSearch.includes(laterSearch) || laterSearch.includes(earlierSearch))
          throw new Error('invalid_tool_input')
      }
    }
  }
}

async function readWordBodyInContext(context: RuntimeRecord, body: RuntimeRecord): Promise<string> {
  const result = (body.getOoxml as () => RuntimeRecord)()
  await (context.sync as () => Promise<void>)()
  if (
    typeof result.value !== 'string' ||
    new TextEncoder().encode(result.value).byteLength > MAX_WORD_OOXML_BYTES
  )
    throw new Error('office_read_failed')
  return result.value
}

function buildDocumentWriteOoxml(original: string, write: WordDocumentWrite): string {
  if (write.blocks.some((block) => block.type === 'list')) throw new Error('office_api_unsupported')
  if (typeof DOMParser === 'undefined' || typeof XMLSerializer === 'undefined')
    throw new Error('office_api_unsupported')
  const document = new DOMParser().parseFromString(original, 'text/xml')
  if (document.getElementsByTagName('parsererror').length) throw new Error('office_read_failed')
  const body = wordBodyElementFromDocument(document)
  const section = directElements(body).find(
    (element) => element.namespaceURI === W_NS && element.localName === 'sectPr',
  )
  const existing = directElements(body).filter((element) => element !== section)
  const headingStyleIds = new Map<number, string>()
  for (const [styleId, level] of headingStyleLevels(original))
    if (!headingStyleIds.has(level)) headingStyleIds.set(level, styleId)
  const nodes = write.blocks.map((block) => createWordBlock(document, block, headingStyleIds))
  const separated: Element[] = []
  for (const node of nodes) {
    if (separated.at(-1)?.localName === 'tbl' && node.localName === 'tbl')
      separated.push(createWordParagraph(document, []))
    separated.push(node)
  }
  if (write.mode === 'append') {
    const boundary = withoutTrailingTableParagraphs(existing)
    for (const element of existing.slice(boundary.length)) element.remove()
    existing.splice(boundary.length)
  }
  if (
    write.mode === 'append' &&
    existing.at(-1)?.localName === 'tbl' &&
    separated[0]?.localName === 'tbl'
  )
    separated.unshift(createWordParagraph(document, []))
  if (
    write.mode === 'prepend' &&
    separated.at(-1)?.localName === 'tbl' &&
    existing[0]?.localName === 'tbl'
  )
    separated.push(createWordParagraph(document, []))
  if (write.mode === 'replace') {
    for (const child of directElements(body)) if (child !== section) child.remove()
  }
  const reference = write.mode === 'prepend' ? body.firstChild : (section ?? null)
  for (const node of separated) body.insertBefore(node, reference)
  return new XMLSerializer().serializeToString(document)
}

function createWordBlock(
  document: Document,
  block: WordDocumentBlock,
  headingStyleIds: Map<number, string>,
): Element {
  if (block.type === 'list') throw new Error('office_api_unsupported')
  if (block.type === 'table') {
    const table = document.createElementNS(W_NS, 'w:tbl')
    const properties = document.createElementNS(W_NS, 'w:tblPr')
    const style = document.createElementNS(W_NS, 'w:tblStyle')
    style.setAttributeNS(W_NS, 'w:val', 'TableGrid')
    properties.append(style)
    table.append(properties)
    for (const [rowIndex, row] of block.rows.entries()) {
      const rowElement = document.createElementNS(W_NS, 'w:tr')
      if (rowIndex < block.headerRows) {
        const rowProperties = document.createElementNS(W_NS, 'w:trPr')
        rowProperties.append(document.createElementNS(W_NS, 'w:tblHeader'))
        rowElement.append(rowProperties)
      }
      for (const value of row) {
        const cell = document.createElementNS(W_NS, 'w:tc')
        cell.append(createWordParagraph(document, [{ text: value }]))
        rowElement.append(cell)
      }
      table.append(rowElement)
    }
    return table
  }
  return createWordParagraph(
    document,
    block.spans,
    block.type === 'heading'
      ? (headingStyleIds.get(block.level - 1) ?? `Heading${block.level}`)
      : undefined,
    block.type === 'heading' ? block.level - 1 : undefined,
  )
}

function createWordParagraph(
  document: Document,
  spans: WordInlineSpan[],
  styleId?: string,
  outlineLevel?: number,
): Element {
  const paragraph = document.createElementNS(W_NS, 'w:p')
  if (styleId) {
    const properties = document.createElementNS(W_NS, 'w:pPr')
    const style = document.createElementNS(W_NS, 'w:pStyle')
    style.setAttributeNS(W_NS, 'w:val', styleId)
    properties.append(style)
    const outline = document.createElementNS(W_NS, 'w:outlineLvl')
    outline.setAttributeNS(W_NS, 'w:val', String(outlineLevel))
    properties.append(outline)
    const runProperties = document.createElementNS(W_NS, 'w:rPr')
    runProperties.append(document.createElementNS(W_NS, 'w:b'))
    const size = document.createElementNS(W_NS, 'w:sz')
    size.setAttributeNS(W_NS, 'w:val', String(Math.max(22, 34 - (outlineLevel ?? 0) * 2)))
    runProperties.append(size)
    properties.append(runProperties)
    paragraph.append(properties)
  }
  for (const span of spans) {
    const run = document.createElementNS(W_NS, 'w:r')
    if (span.bold || span.italic || span.code) {
      const properties = document.createElementNS(W_NS, 'w:rPr')
      if (span.bold) properties.append(document.createElementNS(W_NS, 'w:b'))
      if (span.italic) properties.append(document.createElementNS(W_NS, 'w:i'))
      if (span.code) {
        const fonts = document.createElementNS(W_NS, 'w:rFonts')
        fonts.setAttributeNS(W_NS, 'w:ascii', 'Courier New')
        fonts.setAttributeNS(W_NS, 'w:hAnsi', 'Courier New')
        properties.append(fonts)
      }
      run.append(properties)
    }
    const text = document.createElementNS(W_NS, 'w:t')
    text.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve')
    text.textContent = span.text
    run.append(text)
    paragraph.append(run)
  }
  return paragraph
}

type NativeUnit =
  | {
      type: 'paragraph'
      spans: WordInlineSpan[]
      styleBuiltIn?: string
      list?: boolean
      ordered?: boolean
    }
  | { type: 'table'; rows: string[][]; headerRows: number }

function nativeUnits(blocks: WordDocumentBlock[]): NativeUnit[] {
  return blocks.flatMap((block): NativeUnit[] => {
    if (block.type === 'table')
      return [{ type: 'table', rows: block.rows, headerRows: block.headerRows }]
    if (block.type === 'list')
      return block.items.map((spans) => ({
        type: 'paragraph',
        spans,
        list: true,
        ordered: block.ordered,
      }))
    return [
      {
        type: 'paragraph',
        spans: block.spans,
        ...(block.type === 'heading' ? { styleBuiltIn: `Heading${block.level}` } : {}),
      },
    ]
  })
}

interface ParagraphSignature {
  type: 'paragraph'
  text: string
  style: string
  outlineLevel?: number
  listKind?: 'ordered' | 'unordered'
  characters: Array<{ bold: boolean; italic: boolean; code: boolean }>
}

interface TableSignature {
  type: 'table'
  rows: string[][]
  headerRows: number
}

type DocumentSignature = ParagraphSignature | TableSignature
type RunFormatting = { bold?: boolean; italic?: boolean; code?: boolean }
type WordStyleSemantics = {
  headingLevels: Map<string, number>
  paragraphRuns: Map<string, RunFormatting>
  characterRuns: Map<string, RunFormatting>
  defaults: RunFormatting
}

function wordText(element: Element): string {
  return Array.from(element.getElementsByTagNameNS(W_NS, '*'))
    .filter((child) => ['t', 'tab', 'br', 'cr'].includes(child.localName))
    .map((child) =>
      child.localName === 't' ? (child.textContent ?? '') : child.localName === 'tab' ? '\t' : '\n',
    )
    .join('')
}

function directWordChild(parent: Element, localName: string): Element | undefined {
  return directElements(parent).find(
    (element) => element.namespaceURI === W_NS && element.localName === localName,
  )
}

function paragraphSignature(
  paragraph: Element,
  numbering: Map<string, 'ordered' | 'unordered'>,
  styles: WordStyleSemantics,
): ParagraphSignature {
  const propertiesElement = directWordChild(paragraph, 'pPr') ?? paragraph
  const style = directWordChild(propertiesElement, 'pStyle')
  const outlineValue = directWordChild(propertiesElement, 'outlineLvl')?.getAttributeNS(W_NS, 'val')
  const styleId = style?.getAttributeNS(W_NS, 'val') ?? ''
  const numberId = directWordChild(
    directWordChild(propertiesElement, 'numPr') ?? propertiesElement,
    'numId',
  )?.getAttributeNS(W_NS, 'val')
  const characters: ParagraphSignature['characters'] = []
  const paragraphFormatting = mergeFormatting(
    styles.defaults,
    styles.paragraphRuns.get(styleId),
    runFormatting(directWordChild(propertiesElement, 'rPr')),
  )
  for (const run of Array.from(paragraph.getElementsByTagNameNS(W_NS, 'r'))) {
    const properties = directWordChild(run, 'rPr')
    const characterStyleId = properties
      ? directWordChild(properties, 'rStyle')?.getAttributeNS(W_NS, 'val')
      : undefined
    const formatting = mergeFormatting(
      paragraphFormatting,
      characterStyleId ? styles.characterRuns.get(characterStyleId) : undefined,
      runFormatting(properties),
    )
    for (const _character of Array.from(wordText(run))) {
      characters.push({
        bold: formatting.bold ?? false,
        italic: formatting.italic ?? false,
        code: formatting.code ?? false,
      })
    }
  }
  return {
    type: 'paragraph',
    text: wordText(paragraph),
    style: styleId,
    ...(outlineValue && /^\d+$/.test(outlineValue)
      ? { outlineLevel: Number(outlineValue) }
      : styles.headingLevels.has(styleId)
        ? { outlineLevel: styles.headingLevels.get(styleId) }
        : {}),
    ...(numberId && numbering.has(numberId) ? { listKind: numbering.get(numberId) } : {}),
    characters,
  }
}

function elementSignature(
  element: Element,
  numbering: Map<string, 'ordered' | 'unordered'>,
  styles: WordStyleSemantics,
): DocumentSignature | undefined {
  if (element.namespaceURI !== W_NS) return undefined
  if (element.localName === 'p') return paragraphSignature(element, numbering, styles)
  if (element.localName !== 'tbl') return undefined
  const rows = Array.from(element.getElementsByTagNameNS(W_NS, 'tr')).map((row) =>
    Array.from(row.getElementsByTagNameNS(W_NS, 'tc')).map(wordText),
  )
  const headerRows = Array.from(element.getElementsByTagNameNS(W_NS, 'tr')).filter((row) =>
    Boolean(row.getElementsByTagNameNS(W_NS, 'tblHeader').length),
  ).length
  return { type: 'table', rows, headerRows }
}

function numberingKinds(xml: string): Map<string, 'ordered' | 'unordered'> {
  if (typeof DOMParser === 'undefined') throw new Error('office_api_unsupported')
  const document = new DOMParser().parseFromString(xml, 'text/xml')
  const abstractKinds = new Map<string, 'ordered' | 'unordered'>()
  for (const abstract of Array.from(document.getElementsByTagNameNS(W_NS, 'abstractNum'))) {
    const id = abstract.getAttributeNS(W_NS, 'abstractNumId')
    const format = abstract.getElementsByTagNameNS(W_NS, 'numFmt')[0]?.getAttributeNS(W_NS, 'val')
    if (id && format) abstractKinds.set(id, format === 'bullet' ? 'unordered' : 'ordered')
  }
  const result = new Map<string, 'ordered' | 'unordered'>()
  for (const num of Array.from(document.getElementsByTagNameNS(W_NS, 'num'))) {
    const id = num.getAttributeNS(W_NS, 'numId')
    const abstractId = num
      .getElementsByTagNameNS(W_NS, 'abstractNumId')[0]
      ?.getAttributeNS(W_NS, 'val')
    const kind = abstractId ? abstractKinds.get(abstractId) : undefined
    if (id && kind) result.set(id, kind)
  }
  return result
}

function styleSemantics(xml: string): WordStyleSemantics {
  if (typeof DOMParser === 'undefined') throw new Error('office_api_unsupported')
  const document = new DOMParser().parseFromString(xml, 'text/xml')
  const definitions = new Map<
    string,
    {
      type: string
      basedOn?: string
      link?: string
      default: boolean
      outline?: number
      run: RunFormatting
    }
  >()
  const docDefaults = document.getElementsByTagNameNS(W_NS, 'docDefaults')[0]
  const defaultRunProperties = docDefaults
    ? docDefaults
        .getElementsByTagNameNS(W_NS, 'rPrDefault')[0]
        ?.getElementsByTagNameNS(W_NS, 'rPr')[0]
    : undefined
  const documentDefaults = runFormatting(defaultRunProperties)
  for (const style of Array.from(document.getElementsByTagNameNS(W_NS, 'style'))) {
    const id = style.getAttributeNS(W_NS, 'styleId')
    const type = style.getAttributeNS(W_NS, 'type') ?? ''
    if (!id || (type !== 'paragraph' && type !== 'character')) continue
    const properties = directWordChild(style, 'pPr')
    const value = properties
      ? directWordChild(properties, 'outlineLvl')?.getAttributeNS(W_NS, 'val')
      : undefined
    const basedOn = directWordChild(style, 'basedOn')?.getAttributeNS(W_NS, 'val') ?? undefined
    const link = directWordChild(style, 'link')?.getAttributeNS(W_NS, 'val') ?? undefined
    definitions.set(id, {
      type,
      ...(basedOn ? { basedOn } : {}),
      ...(link ? { link } : {}),
      default: ['1', 'true', 'on'].includes(
        (style.getAttributeNS(W_NS, 'default') ?? '').toLowerCase(),
      ),
      ...(value && /^\d+$/.test(value) && Number(value) >= 0 && Number(value) <= 8
        ? { outline: Number(value) }
        : {}),
      run: runFormatting(directWordChild(style, 'rPr')),
    })
  }
  const headingLevels = new Map<string, number>()
  const paragraphRuns = new Map<string, RunFormatting>()
  const characterRuns = new Map<string, RunFormatting>()
  const resolve = (
    id: string,
    seen = new Set<string>(),
  ): { outline?: number; run: RunFormatting } => {
    if (seen.has(id)) return { run: {} }
    seen.add(id)
    const definition = definitions.get(id)
    if (!definition) return { run: {} }
    const base = definition.basedOn ? resolve(definition.basedOn, seen) : { run: {} }
    return {
      ...(definition.outline !== undefined
        ? { outline: definition.outline }
        : base.outline !== undefined
          ? { outline: base.outline }
          : {}),
      run: mergeFormatting(base.run, definition.run),
    }
  }
  const defaultParagraph = [...definitions].find(
    ([, definition]) => definition.type === 'paragraph' && definition.default,
  )?.[0]
  const defaultCharacter = [...definitions].find(
    ([, definition]) => definition.type === 'character' && definition.default,
  )?.[0]
  const defaults = mergeFormatting(
    documentDefaults,
    defaultParagraph ? resolve(defaultParagraph).run : undefined,
    defaultCharacter ? resolve(defaultCharacter).run : undefined,
  )
  for (const [id, definition] of definitions) {
    const resolved = resolve(id)
    if (definition.type === 'paragraph') {
      if (resolved.outline !== undefined) headingLevels.set(id, resolved.outline)
      paragraphRuns.set(
        id,
        mergeFormatting(
          defaults,
          defaultParagraph && defaultParagraph !== id ? resolve(defaultParagraph).run : undefined,
          resolved.run,
          definition.link ? resolve(definition.link).run : undefined,
        ),
      )
    } else
      characterRuns.set(
        id,
        mergeFormatting(
          defaults,
          defaultCharacter && defaultCharacter !== id ? resolve(defaultCharacter).run : undefined,
          resolved.run,
          definition.link ? resolve(definition.link).run : undefined,
        ),
      )
  }
  return { headingLevels, paragraphRuns, characterRuns, defaults }
}

function headingStyleLevels(xml: string): Map<string, number> {
  return styleSemantics(xml).headingLevels
}

function mergeFormatting(...values: Array<RunFormatting | undefined>): RunFormatting {
  return Object.assign({}, ...values.filter(Boolean))
}

function onOff(properties: Element | undefined, localName: string): boolean | undefined {
  const element = properties ? directWordChild(properties, localName) : undefined
  if (!element) return undefined
  const value = element.getAttributeNS(W_NS, 'val')
  return value === null || !['0', 'false', 'off'].includes(value.toLowerCase())
}

function runFormatting(properties: Element | undefined): RunFormatting {
  const bold = onOff(properties, 'b')
  const italic = onOff(properties, 'i')
  const fonts = properties ? directWordChild(properties, 'rFonts') : undefined
  const fontNames = [
    fonts?.getAttributeNS(W_NS, 'ascii'),
    fonts?.getAttributeNS(W_NS, 'hAnsi'),
  ].filter((name): name is string => Boolean(name))
  return {
    ...(bold !== undefined ? { bold } : {}),
    ...(italic !== undefined ? { italic } : {}),
    ...(fonts ? { code: fontNames.some((name) => name.toLowerCase().includes('courier')) } : {}),
  }
}

function spansMatch(
  expected: WordInlineSpan[],
  actual: ParagraphSignature,
  strict = true,
): boolean {
  const text = expected.map((span) => span.text).join('')
  if (actual.text !== text || actual.characters.length !== Array.from(text).length) return false
  let offset = 0
  for (const span of expected) {
    for (const _character of Array.from(span.text)) {
      const properties = actual.characters[offset]
      if (
        !properties ||
        (strict
          ? properties.bold !== Boolean(span.bold) ||
            properties.italic !== Boolean(span.italic) ||
            properties.code !== Boolean(span.code)
          : (span.bold && !properties.bold) ||
            (span.italic && !properties.italic) ||
            (span.code && !properties.code))
      )
        return false
      offset += 1
    }
  }
  return true
}

function signatureMatches(unit: NativeUnit, actual: DocumentSignature): boolean {
  if (unit.type === 'table')
    return (
      actual.type === 'table' &&
      JSON.stringify(actual.rows) === JSON.stringify(unit.rows) &&
      actual.headerRows === unit.headerRows
    )
  if (
    actual.type !== 'paragraph' ||
    !spansMatch(unit.spans, actual, unit.styleBuiltIn === undefined)
  )
    return false
  if (unit.styleBuiltIn) {
    const level = Number(unit.styleBuiltIn.replace('Heading', '')) - 1
    return actual.outlineLevel === level
  }
  if (unit.list) return actual.listKind === (unit.ordered ? 'ordered' : 'unordered')
  return true
}

function normalizedDocumentText(value: string): string {
  return value
    .replace(/\r\n?|\u2028|\u2029/g, '\n')
    .replace(/[\t ]+$/gm, '')
    .replace(/^\n+|\n+$/g, '')
}

type NativeVerificationStage = 'text' | 'body_shape' | 'content' | 'boundary'
type NativeVerificationResult = { valid: true } | { valid: false; stage: NativeVerificationStage }

function verifyNativeDocumentWriteDetailed(
  before: string,
  after: string,
  write: WordDocumentWrite,
): NativeVerificationResult {
  const beforeParts = bodyParts(before)
  const afterParts = bodyParts(after)
  const expected: NativeUnit[] = []
  for (const unit of nativeUnits(write.blocks)) {
    if (expected.at(-1)?.type === 'table' && unit.type === 'table')
      expected.push({ type: 'paragraph', spans: [] })
    expected.push(unit)
  }
  const firstBefore = beforeParts.content[0]
  const appendBoundary = withoutTrailingTableParagraphs(beforeParts.content)
  const lastBefore = appendBoundary.at(-1)
  if (write.mode === 'append' && lastBefore?.localName === 'tbl' && expected[0]?.type === 'table')
    expected.unshift({ type: 'paragraph', spans: [] })
  if (
    write.mode === 'prepend' &&
    firstBefore?.localName === 'tbl' &&
    expected.at(-1)?.type === 'table'
  )
    expected.push({ type: 'paragraph', spans: [] })
  const insertedText = normalizedDocumentText(
    expected
      .map((unit) =>
        unit.type === 'table'
          ? unit.rows.flat().join('\n')
          : unit.spans.map((span) => span.text).join(''),
      )
      .join('\n'),
  )
  const beforeText = normalizedDocumentText(ooxmlText(before))
  const afterText = normalizedDocumentText(ooxmlText(after))
  if (
    (write.mode === 'replace' && afterText !== insertedText) ||
    (write.mode === 'append' &&
      (!afterText.startsWith(beforeText) || !afterText.endsWith(insertedText))) ||
    (write.mode === 'prepend' &&
      (!afterText.startsWith(insertedText) || !afterText.endsWith(beforeText)))
  )
    return { valid: false, stage: 'text' }
  const insertedMatches = (elements: Element[]) =>
    elements.length === expected.length &&
    expected.every((unit, offset) => {
      const signature = elementSignature(elements[offset], afterParts.numbering, afterParts.styles)
      return signature ? signatureMatches(unit, signature) : false
    })
  const canonicalEqual = (left: Element[], right: Element[]) =>
    left.length === right.length &&
    left.every(
      (element, index) =>
        JSON.stringify(canonicalWordNode(element)) ===
        JSON.stringify(canonicalWordNode(right[index])),
    )
  let valid: boolean
  if (write.mode === 'replace') {
    valid = insertedMatches(afterParts.content)
    if (
      !valid &&
      afterParts.content.length > expected.length &&
      insertedMatches(afterParts.content.slice(0, expected.length)) &&
      expected[expected.length - 1]?.type === 'table'
    )
      valid = afterParts.content
        .slice(expected.length)
        .every((element) =>
          isEmptyParagraph(elementSignature(element, afterParts.numbering, afterParts.styles)),
        )
  } else if (write.mode === 'append') {
    const afterContent = withoutTrailingEmptyParagraphs(afterParts.content)
    const inserted = afterContent.slice(-expected.length)
    const currentBoundary = withoutTrailingTableParagraphs(afterContent.slice(0, -expected.length))
    valid = canonicalEqual(appendBoundary, currentBoundary) && insertedMatches(inserted)
  } else {
    valid =
      afterParts.content.length === expected.length + beforeParts.content.length &&
      insertedMatches(afterParts.content.slice(0, expected.length)) &&
      canonicalEqual(beforeParts.content, afterParts.content.slice(expected.length))
  }
  if (valid) return { valid: true }
  const expectedLength =
    expected.length + (write.mode === 'replace' ? 0 : beforeParts.content.length)
  if (
    afterParts.content.length !== expectedLength &&
    !(
      write.mode === 'replace' &&
      expected.at(-1)?.type === 'table' &&
      afterParts.content.length === expectedLength + 1
    )
  )
    return { valid: false, stage: 'body_shape' }
  if (write.mode !== 'replace') return { valid: false, stage: 'boundary' }
  return { valid: false, stage: 'content' }
}

function verifyNativeDocumentWrite(
  before: string,
  after: string,
  write: WordDocumentWrite,
): boolean {
  return verifyNativeDocumentWriteDetailed(before, after, write).valid
}

function bodyParts(xml: string): {
  content: Element[]
  numbering: Map<string, 'ordered' | 'unordered'>
  styles: WordStyleSemantics
} {
  const elements = directElements(wordBodyElement(xml))
  if (elements.at(-1)?.namespaceURI === W_NS && elements.at(-1)?.localName === 'sectPr')
    elements.pop()
  return {
    content: elements,
    numbering: numberingKinds(xml),
    styles: styleSemantics(xml),
  }
}

function isEmptyParagraph(signature: DocumentSignature | undefined): boolean {
  return signature?.type === 'paragraph' && signature.text === ''
}

function withoutTrailingTableParagraphs(elements: Element[]): Element[] {
  const boundary = withoutTrailingEmptyParagraphs(elements)
  return boundary.length < elements.length && boundary.at(-1)?.localName === 'tbl'
    ? boundary
    : elements
}

function withoutTrailingEmptyParagraphs(elements: Element[]): Element[] {
  let boundary = elements.length
  while (
    boundary > 0 &&
    elements[boundary - 1]?.namespaceURI === W_NS &&
    elements[boundary - 1]?.localName === 'p' &&
    wordText(elements[boundary - 1]!) === ''
  )
    boundary -= 1
  return elements.slice(0, boundary)
}

function stableFingerprintOoxml(xml: string): string {
  return JSON.stringify({
    body: canonicalWordNode(wordBodyElement(xml)),
    styles: wordStylesElement(xml) ? canonicalWordNode(wordStylesElement(xml)!) : undefined,
  })
}

function stableNonTextFingerprintOoxml(xml: string): string {
  return JSON.stringify({
    body: canonicalWordNode(wordBodyElement(xml), true),
    styles: wordStylesElement(xml) ? canonicalWordNode(wordStylesElement(xml)!) : undefined,
  })
}

function wordStylesElement(xml: string): Element | undefined {
  if (typeof DOMParser === 'undefined') throw new Error('office_api_unsupported')
  const document = new DOMParser().parseFromString(xml, 'text/xml')
  if (document.getElementsByTagName('parsererror').length) throw new Error('office_read_failed')
  const root = document.documentElement
  if (root.namespaceURI !== PKG_NS || root.localName !== 'package') return undefined
  const parts = Array.from(document.getElementsByTagNameNS(PKG_NS, 'part')).filter(
    (part) => part.getAttributeNS(PKG_NS, 'name') === '/word/styles.xml',
  )
  if (parts.length === 0) return undefined
  if (parts.length !== 1) throw new Error('office_read_failed')
  const xmlData = directElements(parts[0]).filter(
    (element) => element.namespaceURI === PKG_NS && element.localName === 'xmlData',
  )
  if (xmlData.length !== 1) throw new Error('office_read_failed')
  const styles = directElements(xmlData[0]).filter(
    (element) => element.namespaceURI === W_NS && element.localName === 'styles',
  )
  if (styles.length !== 1) throw new Error('office_read_failed')
  return styles[0]
}

function wordBodyElement(xml: string): Element {
  if (typeof DOMParser === 'undefined') throw new Error('office_api_unsupported')
  const document = new DOMParser().parseFromString(xml, 'text/xml')
  if (document.getElementsByTagName('parsererror').length) throw new Error('office_read_failed')
  return wordBodyElementFromDocument(document)
}

function wordBodyElementFromDocument(document: Document): Element {
  const root = document.documentElement
  let wordDocument: Element | undefined
  if (root.namespaceURI === PKG_NS && root.localName === 'package') {
    const parts = Array.from(document.getElementsByTagNameNS(PKG_NS, 'part')).filter(
      (part) => part.getAttributeNS(PKG_NS, 'name') === '/word/document.xml',
    )
    if (parts.length !== 1) throw new Error('office_read_failed')
    const xmlData = directElements(parts[0]).filter(
      (element) => element.namespaceURI === PKG_NS && element.localName === 'xmlData',
    )
    if (xmlData.length !== 1) throw new Error('office_read_failed')
    const documents = directElements(xmlData[0]).filter(
      (element) => element.namespaceURI === W_NS && element.localName === 'document',
    )
    if (documents.length !== 1) throw new Error('office_read_failed')
    wordDocument = documents[0]
  } else if (root.namespaceURI === W_NS && root.localName === 'document') {
    wordDocument = root
  }
  const bodies = wordDocument
    ? directElements(wordDocument).filter(
        (element) => element.namespaceURI === W_NS && element.localName === 'body',
      )
    : root.namespaceURI === W_NS && root.localName === 'body'
      ? [root]
      : []
  if (bodies.length !== 1) throw new Error('office_read_failed')
  return bodies[0]
}

function directElements(parent: Element): Element[] {
  return Array.from(parent.childNodes).filter((node): node is Element => node.nodeType === 1)
}

type CanonicalWordNode =
  | { element: string; attributes: Array<[string, string]>; children: CanonicalWordNode[] }
  | { text: string }

function canonicalWordNode(node: Node, redactText = false): CanonicalWordNode {
  if (node.nodeType === 3) return { text: redactText ? '' : (node.nodeValue ?? '') }
  const element = node as Element
  const preserveWhitespace =
    element.namespaceURI === W_NS &&
    (element.localName === 't' ||
      element.localName === 'delText' ||
      element.localName === 'instrText')
  const children = Array.from(element.childNodes).flatMap((child): CanonicalWordNode[] => {
    if (child.nodeType === 1) {
      const childElement = child as Element
      if (
        childElement.namespaceURI === W_NS &&
        (childElement.localName === 'proofErr' ||
          childElement.localName === 'lastRenderedPageBreak')
      )
        return []
      return [canonicalWordNode(child, redactText)]
    }
    if (child.nodeType !== 3) return []
    const value = child.nodeValue ?? ''
    if (!preserveWhitespace && /^\s*$/.test(value)) return []
    return [{ text: redactText && preserveWhitespace ? '' : value }]
  })
  const attributes = Array.from(element.attributes)
    .filter((attribute) => {
      if (attribute.namespaceURI === XMLNS_NS || attribute.name === 'xmlns') return false
      if (attribute.namespaceURI === W_NS && attribute.localName.startsWith('rsid')) return false
      return !(
        attribute.namespaceURI === W14_NS &&
        (attribute.localName === 'paraId' || attribute.localName === 'textId')
      )
    })
    .map((attribute): [string, string] => [
      `{${attribute.namespaceURI ?? ''}}${attribute.localName}`,
      attribute.value,
    ])
    .sort(([left], [right]) => left.localeCompare(right))
  return {
    element: `{${element.namespaceURI ?? ''}}${element.localName}`,
    attributes,
    children,
  }
}

function ooxmlText(xml: string): string {
  const body = wordBodyElement(xml)
  const paragraphs = Array.from(body.getElementsByTagNameNS(W_NS, 'p'))
  if (paragraphs.length === 0 && body.getElementsByTagNameNS(W_NS, 't').length > 0)
    throw new Error('office_read_failed')
  return paragraphs.map(wordText).join('\n')
}

function applyTextOperation(
  value: string,
  operation: WordDeclarativeOperation,
  expectedMatches?: number,
): string {
  if (operation.op === 'insert_text')
    return operation.location === 'start'
      ? `${operation.text}${value}`
      : operation.location === 'end'
        ? `${value}${operation.text}`
        : operation.text
  if (expectedMatches === 0) return value
  const escaped = operation.search.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const expression = new RegExp(escaped, operation.matchCase ? 'g' : 'gi')
  let replaced = 0
  const result = value.replace(expression, (match) => {
    if (expectedMatches !== undefined && replaced >= expectedMatches) return match
    replaced += 1
    return operation.replacement
  })
  if (expectedMatches !== undefined && replaced !== expectedMatches)
    throw new Error('office_write_failed')
  return result
}
