export const MAX_WORD_PARAGRAPHS = 500
export const MAX_WORD_TEXT_LENGTH = 12_000
export const MAX_WORD_OOXML_BYTES = 1024 * 1024
export const MAX_WORD_RESULT_BYTES = 256 * 1024
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

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

export interface WordAdapter {
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

export class BrowserWordAdapter implements WordAdapter {
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
    cancelled(signal)
    const { word } = runtime()
    const xml = (word.run as (callback: (context: RuntimeRecord) => unknown) => Promise<unknown>)(
      async (context) => {
        const document = context.document as RuntimeRecord
        const body = document.body as RuntimeRecord
        if (typeof body.getOoxml !== 'function') throw new Error('office_api_unsupported')
        const result = (body.getOoxml as () => RuntimeRecord)()
        await sync(context, signal)
        if (typeof result.value !== 'string') throw new Error('office_read_failed')
        return result.value
      },
    ) as Promise<string>
    return summarizeOoxml(await xml, options)
  }

  async screenshotDocument(_page: number, signal?: AbortSignal): Promise<WordScreenshotResult> {
    cancelled(signal)
    runtime()
    // Release blocker: PDF export plus bounded browser rendering has not yet received its
    // dependency/CSP audit. The compatibility tool intentionally fails closed until it does.
    throw new Error('office_api_unsupported')
  }
}
