export interface WordInlineSpan {
  text: string
  bold?: true
  italic?: true
  code?: true
}

export type WordDocumentBlock =
  | { type: 'paragraph'; spans: WordInlineSpan[] }
  | { type: 'heading'; level: number; spans: WordInlineSpan[] }
  | { type: 'list'; ordered: boolean; items: WordInlineSpan[][] }
  | { type: 'table'; rows: string[][]; headerRows: 1 }

export interface WordDocumentWrite {
  mode: 'replace' | 'append' | 'prepend'
  blocks: WordDocumentBlock[]
  semanticText: string
  structure: { headings: number; lists: number; tables: number }
}

const TABLE_SEPARATOR = /^\s*\|?\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)+\|?\s*$/
const ORDERED_ITEM = /^\s*\d+[.)]\s+(.+)$/
const UNORDERED_ITEM = /^\s*[-*+]\s+(.+)$/
const MAX_BLOCKS = 500
const MAX_TABLE_COLUMNS = 20
const MAX_TABLE_ROWS = 100
const MAX_TABLE_CELLS = 1_000
const MAX_INLINE_SPANS = 1_000

function invalid(): never {
  throw new Error('invalid_tool_input')
}

function literalInline(value: string): boolean {
  return value.includes('[') || value.includes(']') || /<[^>]*>/.test(value)
}

function inline(value: string): WordInlineSpan[] {
  if (literalInline(value)) return [{ text: value }]
  const spans: WordInlineSpan[] = []
  let plain = ''
  const flush = () => {
    if (plain) spans.push({ text: plain })
    plain = ''
  }
  for (let index = 0; index < value.length;) {
    const marker = value[index]
    const double = value.slice(index, index + 2)
    const delimiter = marker === '`' ? '`' : double === '**' || double === '__' ? double : marker
    const format =
      marker === '`'
        ? 'code'
        : double === '**' || double === '__'
          ? 'bold'
          : marker === '*' || marker === '_'
            ? 'italic'
            : undefined
    if (!format) {
      plain += marker
      index += 1
      continue
    }
    const start = index + delimiter.length
    const end = value.indexOf(delimiter, start)
    if (end <= start || (delimiter.length === 1 && value[end + 1] === delimiter)) {
      plain += marker
      index += 1
      continue
    }
    flush()
    spans.push({ text: value.slice(start, end), [format]: true })
    index = end + delimiter.length
  }
  flush()
  return spans.length ? spans : [{ text: '' }]
}

function spanText(spans: WordInlineSpan[]): string {
  return spans.map((span) => span.text).join('')
}

function tableCells(line: string): string[] {
  return line
    .trim()
    .replace(/^\|/, '')
    .replace(/\|$/, '')
    .split('|')
    .map((cell) => spanText(inline(cell.trim())))
}

function hasFormattedTableCell(line: string): boolean {
  return line
    .trim()
    .replace(/^\|/, '')
    .replace(/\|$/, '')
    .split('|')
    .some((cell) => inline(cell.trim()).some((span) => span.bold || span.italic || span.code))
}

function assertSpanBudget(blocks: WordDocumentBlock[]): void {
  const count = blocks.reduce(
    (total, block) =>
      total +
      (block.type === 'table'
        ? 0
        : block.type === 'list'
          ? block.items.reduce((sum, item) => sum + item.length, 0)
          : block.spans.length),
    0,
  )
  if (count > MAX_INLINE_SPANS) invalid()
}

export function parseWordMarkdown(text: string): Omit<WordDocumentWrite, 'mode'> {
  const lines = text.replace(/\r\n?/g, '\n').split('\n')
  if (lines.length > MAX_BLOCKS) invalid()
  const blocks: WordDocumentBlock[] = []
  const semantic: string[] = []
  let headings = 0
  let lists = 0
  let tables = 0
  let index = 0

  while (index < lines.length) {
    const line = lines[index]
    if (line.trim() === '') {
      index += 1
      continue
    }
    const heading = /^(#{1,6})\s+(.+)$/.exec(line)
    if (heading) {
      const spans = inline(heading[2].trim())
      blocks.push({ type: 'heading', level: heading[1].length, spans })
      semantic.push(spanText(spans))
      headings += 1
      index += 1
      continue
    }
    const header = tableCells(line)
    if (header.length > 1 && index + 1 < lines.length && TABLE_SEPARATOR.test(lines[index + 1])) {
      if (hasFormattedTableCell(line)) invalid()
      if (header.length > MAX_TABLE_COLUMNS) invalid()
      const rows = [header]
      index += 2
      while (index < lines.length && lines[index].includes('|') && lines[index].trim() !== '') {
        if (hasFormattedTableCell(lines[index])) invalid()
        const row = tableCells(lines[index])
        if (row.length !== header.length) break
        rows.push(row)
        if (rows.length > MAX_TABLE_ROWS || rows.length * header.length > MAX_TABLE_CELLS) invalid()
        index += 1
      }
      blocks.push({ type: 'table', rows, headerRows: 1 })
      semantic.push(...rows.flat())
      tables += 1
      continue
    }
    const ordered = ORDERED_ITEM.exec(line)
    const unordered = UNORDERED_ITEM.exec(line)
    if (ordered || unordered) {
      const expression = ordered ? ORDERED_ITEM : UNORDERED_ITEM
      const items: WordInlineSpan[][] = []
      while (index < lines.length) {
        const item = expression.exec(lines[index])
        if (!item) break
        items.push(inline(item[1].trim()))
        if (items.length > MAX_BLOCKS) invalid()
        index += 1
      }
      blocks.push({ type: 'list', ordered: Boolean(ordered), items })
      semantic.push(...items.map(spanText))
      lists += 1
      continue
    }
    const spans = inline(line)
    blocks.push({ type: 'paragraph', spans })
    semantic.push(spanText(spans))
    index += 1
  }
  if (blocks.length === 0 || blocks.length > MAX_BLOCKS) invalid()
  assertSpanBudget(blocks)
  return {
    blocks,
    semanticText: semantic.join('\n'),
    structure: { headings, lists, tables },
  }
}

export function parseWordPlainText(text: string): Omit<WordDocumentWrite, 'mode'> {
  const lines = text.replace(/\r\n?/g, '\n').split('\n')
  if (lines.length > MAX_BLOCKS) invalid()
  const blocks = lines.map((line) => ({ type: 'paragraph' as const, spans: [{ text: line }] }))
  assertSpanBudget(blocks)
  return {
    blocks,
    semanticText: lines.join('\n'),
    structure: { headings: 0, lists: 0, tables: 0 },
  }
}
