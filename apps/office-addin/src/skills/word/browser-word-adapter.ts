import type { WordDocumentBlock, WordDocumentWrite, WordInlineSpan } from './word-markdown.js'

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
  private expectedText: string | undefined
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
        await sync(context, signal)
        this.expectedText = expected
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
    return ooxmlText(await readBodyOoxml(signal)) === expected
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
        const beforeFingerprint = stableFingerprintOoxml(snapshot.value)
        cancelled(signal)
        const target = buildDocumentWriteOoxml(snapshot.value, write)
        ;(body.insertOoxml as (xml: string, location: string) => void)(target, 'Replace')
        const result = (body.getOoxml as () => RuntimeRecord)()
        let writeError: unknown
        try {
          await (context.sync as () => Promise<void>)()
        } catch (error) {
          writeError = error
        }
        if (writeError || signal?.aborted) {
          await reconcileAtomicDocumentWrite(context, body, snapshot.value, write)
          throw writeError ?? new Error('cancelled')
        }
        const postValue = typeof result.value === 'string' ? result.value : undefined
        const verified =
          postValue !== undefined &&
          new TextEncoder().encode(postValue).byteLength <= MAX_WORD_OOXML_BYTES &&
          verifyNativeDocumentWrite(snapshot.value, postValue, write)
        if (!verified || signal?.aborted) {
          await reconcileAtomicDocumentWrite(context, body, snapshot.value, write)
          throw signal?.aborted ? new Error('cancelled') : new Error('office_verify_failed')
        }
        this.expectedDocument = {
          fingerprint: stableFingerprintOoxml(postValue!),
          original: snapshot.value,
          originalFingerprint: beforeFingerprint,
          write,
        }
      },
    )
  }

  async verifyDocumentWrite(_write: WordDocumentWrite, signal?: AbortSignal): Promise<boolean> {
    const expected = this.expectedDocument
    this.expectedDocument = undefined
    if (!expected) return false
    let current: string
    try {
      current = await readBodyOoxml()
    } catch (error) {
      throw new Error('office_recovery_failed', { cause: error })
    }
    const currentFingerprint = stableFingerprintOoxml(current)
    if (currentFingerprint === expected.fingerprint && !signal?.aborted) return true
    if (currentFingerprint === expected.originalFingerprint) return false
    if (!signal?.aborted && verifyNativeDocumentWrite(expected.original, current, expected.write))
      return true
    const { word } = runtime()
    await (word.run as (callback: (context: RuntimeRecord) => unknown) => Promise<unknown>)(
      async (context) => {
        const body = (context.document as RuntimeRecord).body as RuntimeRecord
        if (typeof body.insertOoxml !== 'function' || typeof body.getOoxml !== 'function')
          throw new Error('office_recovery_failed')
        await restoreWordBody(context, body, expected.original, expected.originalFingerprint)
      },
    )
    return false
  }
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

async function restoreWordBody(
  context: RuntimeRecord,
  body: RuntimeRecord,
  original: string,
  fingerprint: string,
): Promise<void> {
  try {
    ;(body.insertOoxml as (xml: string, location: string) => void)(original, 'Replace')
    await (context.sync as () => Promise<void>)()
    const restored = (body.getOoxml as () => RuntimeRecord)()
    await (context.sync as () => Promise<void>)()
    if (
      typeof restored.value !== 'string' ||
      stableFingerprintOoxml(restored.value) !== fingerprint
    )
      throw new Error('office_recovery_failed')
  } catch (error) {
    if (error instanceof Error && error.message === 'office_recovery_failed') throw error
    throw new Error('office_recovery_failed', { cause: error })
  }
}

async function reconcileAtomicDocumentWrite(
  context: RuntimeRecord,
  body: RuntimeRecord,
  original: string,
  write: WordDocumentWrite,
): Promise<void> {
  try {
    const current = (body.getOoxml as () => RuntimeRecord)()
    await (context.sync as () => Promise<void>)()
    if (typeof current.value !== 'string') throw new Error('office_recovery_failed')
    const fingerprint = stableFingerprintOoxml(current.value)
    const originalFingerprint = stableFingerprintOoxml(original)
    if (fingerprint === originalFingerprint) return
    if (!verifyNativeDocumentWrite(original, current.value, write))
      throw new Error('office_recovery_failed')
    await restoreWordBody(context, body, original, originalFingerprint)
  } catch (error) {
    if (error instanceof Error && error.message === 'office_recovery_failed') throw error
    throw new Error('office_recovery_failed', { cause: error })
  }
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
  headingStyles: Map<string, number>,
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
  for (const run of Array.from(paragraph.getElementsByTagNameNS(W_NS, 'r'))) {
    const properties = directWordChild(run, 'rPr')
    const bold = Boolean(properties?.getElementsByTagNameNS(W_NS, 'b').length)
    const italic = Boolean(properties?.getElementsByTagNameNS(W_NS, 'i').length)
    const fonts = properties?.getElementsByTagNameNS(W_NS, 'rFonts')[0]
    const code = [fonts?.getAttributeNS(W_NS, 'ascii'), fonts?.getAttributeNS(W_NS, 'hAnsi')]
      .filter(Boolean)
      .some((name) => name!.toLowerCase().includes('courier'))
    for (const _character of Array.from(wordText(run))) {
      characters.push({ bold, italic, code })
    }
  }
  return {
    type: 'paragraph',
    text: wordText(paragraph),
    style: styleId,
    ...(outlineValue && /^\d+$/.test(outlineValue)
      ? { outlineLevel: Number(outlineValue) }
      : headingStyles.has(styleId)
        ? { outlineLevel: headingStyles.get(styleId) }
        : {}),
    ...(numberId && numbering.has(numberId) ? { listKind: numbering.get(numberId) } : {}),
    characters,
  }
}

function elementSignature(
  element: Element,
  numbering: Map<string, 'ordered' | 'unordered'>,
  headingStyles: Map<string, number>,
): DocumentSignature | undefined {
  if (element.namespaceURI !== W_NS) return undefined
  if (element.localName === 'p') return paragraphSignature(element, numbering, headingStyles)
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

function headingStyleLevels(xml: string): Map<string, number> {
  if (typeof DOMParser === 'undefined') throw new Error('office_api_unsupported')
  const document = new DOMParser().parseFromString(xml, 'text/xml')
  const result = new Map<string, number>()
  for (const style of Array.from(document.getElementsByTagNameNS(W_NS, 'style'))) {
    if (style.getAttributeNS(W_NS, 'type') !== 'paragraph') continue
    const id = style.getAttributeNS(W_NS, 'styleId')
    const properties = directWordChild(style, 'pPr')
    const value = properties
      ? directWordChild(properties, 'outlineLvl')?.getAttributeNS(W_NS, 'val')
      : undefined
    if (id && value && /^\d+$/.test(value) && Number(value) >= 0 && Number(value) <= 8)
      result.set(id, Number(value))
  }
  return result
}

function spansMatch(expected: WordInlineSpan[], actual: ParagraphSignature): boolean {
  const text = expected.map((span) => span.text).join('')
  if (actual.text !== text || actual.characters.length !== Array.from(text).length) return false
  let offset = 0
  for (const span of expected) {
    for (const _character of Array.from(span.text)) {
      const properties = actual.characters[offset]
      if (
        !properties ||
        (span.bold && !properties.bold) ||
        (span.italic && !properties.italic) ||
        (span.code && !properties.code)
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
  if (actual.type !== 'paragraph' || !spansMatch(unit.spans, actual)) return false
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

function verifyNativeDocumentWrite(
  before: string,
  after: string,
  write: WordDocumentWrite,
): boolean {
  const beforeParts = bodyParts(before)
  const afterParts = bodyParts(after)
  const expected: NativeUnit[] = []
  for (const unit of nativeUnits(write.blocks)) {
    if (expected.at(-1)?.type === 'table' && unit.type === 'table')
      expected.push({ type: 'paragraph', spans: [] })
    expected.push(unit)
  }
  const firstBefore = beforeParts.content[0]
  const lastBefore = beforeParts.content.at(-1)
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
    return false
  const insertedMatches = (elements: Element[]) =>
    elements.length === expected.length &&
    expected.every((unit, offset) => {
      const signature = elementSignature(
        elements[offset],
        afterParts.numbering,
        afterParts.headingStyles,
      )
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
      afterParts.content.length === expected.length + 1 &&
      insertedMatches(afterParts.content.slice(0, -1)) &&
      isEmptyParagraph(
        elementSignature(
          afterParts.content[afterParts.content.length - 1],
          afterParts.numbering,
          afterParts.headingStyles,
        ),
      ) &&
      expected[expected.length - 1]?.type === 'table'
    )
      valid = true
  } else if (write.mode === 'append') {
    const boundary = beforeParts.content.length
    valid =
      afterParts.content.length === boundary + expected.length &&
      canonicalEqual(beforeParts.content, afterParts.content.slice(0, boundary)) &&
      insertedMatches(afterParts.content.slice(boundary))
  } else {
    valid =
      afterParts.content.length === expected.length + beforeParts.content.length &&
      insertedMatches(afterParts.content.slice(0, expected.length)) &&
      canonicalEqual(beforeParts.content, afterParts.content.slice(expected.length))
  }
  return valid
}

function bodyParts(xml: string): {
  content: Element[]
  numbering: Map<string, 'ordered' | 'unordered'>
  headingStyles: Map<string, number>
} {
  const elements = directElements(wordBodyElement(xml))
  if (elements.at(-1)?.namespaceURI === W_NS && elements.at(-1)?.localName === 'sectPr')
    elements.pop()
  return {
    content: elements,
    numbering: numberingKinds(xml),
    headingStyles: headingStyleLevels(xml),
  }
}

function isEmptyParagraph(signature: DocumentSignature | undefined): boolean {
  return signature?.type === 'paragraph' && signature.text === ''
}

function stableFingerprintOoxml(xml: string): string {
  return JSON.stringify(canonicalWordNode(wordBodyElement(xml)))
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

function canonicalWordNode(node: Node): CanonicalWordNode {
  if (node.nodeType === 3) return { text: node.nodeValue ?? '' }
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
      return [canonicalWordNode(child)]
    }
    if (child.nodeType !== 3) return []
    const value = child.nodeValue ?? ''
    if (!preserveWhitespace && /^\s*$/.test(value)) return []
    return [{ text: value }]
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
  const paragraphs = [...xml.matchAll(/<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g)]
  if (paragraphs.length === 0 && /<w:t(?:\s[^>]*)?>/.test(xml))
    throw new Error('office_read_failed')
  return paragraphs
    .map((paragraph) =>
      [...paragraph[1].matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:(tab|br|cr)\s*\/>/g)]
        .map((token) => (token[2] === 'tab' ? '\t' : token[2] ? '\n' : decodeXml(token[1])))
        .join(''),
    )
    .join('\n')
}

function decodeXml(value: string): string {
  return value.replace(/&(lt|gt|amp|quot|apos|#\d+|#x[\da-f]+);/gi, (entity, code: string) => {
    const named: Record<string, string> = {
      lt: '<',
      gt: '>',
      amp: '&',
      quot: '"',
      apos: "'",
    }
    const normalized = code.toLowerCase()
    if (normalized in named) return named[normalized]
    const numeric = normalized.startsWith('#x')
      ? Number.parseInt(normalized.slice(2), 16)
      : Number.parseInt(normalized.slice(1), 10)
    return Number.isFinite(numeric) && numeric >= 0 && numeric <= 0x10ffff
      ? String.fromCodePoint(numeric)
      : entity
  })
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
