import { XMLParser } from 'fast-xml-parser'
import JSZip from 'jszip'
import type { InMemoryVfs } from './vfs.js'

const MAX_PDF_PAGES = 20
const MAX_OUTPUT_BYTES = 2 * 1024 * 1024
const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  trimValues: false,
  parseTagValue: false,
})
const array = <T>(value: T | T[] | undefined): T[] =>
  value === undefined ? [] : Array.isArray(value) ? value : [value]
const nodeText = (value: unknown): string =>
  value == null
    ? ''
    : typeof value === 'object'
      ? nodeText((value as Record<string, unknown>)['#text'])
      : String(value)
const bounded = (value: string): string => {
  if (new TextEncoder().encode(value).byteLength > MAX_OUTPUT_BYTES) throw new Error('vfs_limit')
  return value
}
const csv = (value: string): string =>
  /[",\r\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value
const check = (signal?: AbortSignal) => {
  if (signal?.aborted) throw new Error('cancelled')
}

export async function docxToText(bytes: Uint8Array, signal?: AbortSignal): Promise<string> {
  check(signal)
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('word/document.xml')?.async('text')
  check(signal)
  if (!xml) throw new Error('command_failed')
  const document = parser.parse(xml) as Record<string, any>
  const paragraphs = array(document['w:document']?.['w:body']?.['w:p']) as Record<string, any>[]
  return bounded(paragraphs.map((paragraph) => collectText(paragraph)).join('\n'))
}

function collectText(value: unknown): string {
  if (value == null) return ''
  if (Array.isArray(value)) return value.map(collectText).join('')
  if (typeof value !== 'object') return ''
  return Object.entries(value as Record<string, unknown>)
    .map(([key, child]) =>
      key === 'w:t' ? nodeText(child) : key.startsWith('@_') ? '' : collectText(child),
    )
    .join('')
}

export async function xlsxToCsv(bytes: Uint8Array, signal?: AbortSignal): Promise<string> {
  check(signal)
  const zip = await JSZip.loadAsync(bytes)
  const workbookXml = await zip.file('xl/workbook.xml')?.async('text')
  if (!workbookXml) throw new Error('command_failed')
  const workbook = parser.parse(workbookXml) as Record<string, any>
  const sheets = array(workbook.workbook?.sheets?.sheet) as Record<string, unknown>[]
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('text')
  const targets = new Map<string, string>()
  if (relsXml) {
    const rels = parser.parse(relsXml) as Record<string, any>
    for (const rel of array(rels.Relationships?.Relationship) as Record<string, unknown>[])
      targets.set(String(rel['@_Id']), `xl/${String(rel['@_Target']).replace(/^\/?xl\//, '')}`)
  }
  const shared: string[] = []
  const sharedXml = await zip.file('xl/sharedStrings.xml')?.async('text')
  if (sharedXml) {
    const parsed = parser.parse(sharedXml) as Record<string, any>
    for (const item of array(parsed.sst?.si) as Record<string, unknown>[])
      shared.push(collectSpreadsheetText(item))
  }
  const sections: string[] = []
  for (const sheet of sheets.slice(0, 32)) {
    check(signal)
    const path = targets.get(String(sheet['@_r:id']))
    const xml = path ? await zip.file(path)?.async('text') : undefined
    if (!xml) continue
    const parsed = parser.parse(xml) as Record<string, any>
    const lines = [`# ${csv(String(sheet['@_name'] ?? ''))}`]
    for (const row of array(parsed.worksheet?.sheetData?.row) as Record<string, any>[]) {
      const cells: string[] = []
      for (const cell of array(row.c) as Record<string, unknown>[]) {
        const reference = String(cell['@_r'] ?? '')
        const letters = /^[A-Z]+/i.exec(reference)?.[0]?.toUpperCase() ?? ''
        const column =
          letters.split('').reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0) - 1
        while (cells.length < Math.max(column, 0)) cells.push('')
        const raw = nodeText(cell.v)
        const value =
          cell['@_t'] === 's'
            ? (shared[Number(raw)] ?? '')
            : cell['@_t'] === 'inlineStr'
              ? collectSpreadsheetText(cell.is)
              : raw
        cells[Math.max(column, 0)] = csv(value)
      }
      lines.push(cells.join(','))
    }
    sections.push(lines.join('\n'))
  }
  return bounded(sections.join('\n\n'))
}

function collectSpreadsheetText(value: unknown): string {
  if (!value || typeof value !== 'object') return nodeText(value)
  const record = value as Record<string, unknown>
  if (record.t !== undefined) return nodeText(record.t)
  return array(record.r)
    .map((run) => collectSpreadsheetText(run))
    .join('')
}

export function createConversionCommands(vfs: InMemoryVfs) {
  return {
    'pdf-to-text': async (args: string[], signal?: AbortSignal) =>
      convert(
        args,
        vfs,
        async (bytes, activeSignal) =>
          (await import('./browser-pdf.js')).pdfToText(bytes, activeSignal),
        signal,
      ),
    'docx-to-text': async (args: string[], signal?: AbortSignal) =>
      convert(args, vfs, docxToText, signal),
    'xlsx-to-csv': async (args: string[], signal?: AbortSignal) =>
      convert(args, vfs, xlsxToCsv, signal),
    'pdf-to-images': async (args: string[], signal?: AbortSignal) => {
      if (args.length < 2 || args.length > 3) return { output: '', error: 'sandbox_denied' }
      const pages = args[2] === undefined ? 1 : Number(args[2])
      if (!Number.isInteger(pages) || pages < 1 || pages > MAX_PDF_PAGES)
        return { output: '', error: 'sandbox_denied' }
      const bytes = vfs.readBytes(args[0])
      const { renderPdfPageToPng } = await import('./browser-pdf.js')
      const outputs: string[] = []
      for (let page = 1; page <= pages; page += 1) {
        check(signal)
        const base64 = await renderPdfPageToPng(bytes, page, signal)
        const path = `${args[1].replace(/\/$/, '')}/page-${page}.png`
        vfs.writeFile(
          path,
          Uint8Array.from(atob(base64), (character) => character.charCodeAt(0)),
        )
        outputs.push(path)
      }
      return { output: outputs.join('\n') }
    },
  }
}

async function convert(
  args: string[],
  vfs: InMemoryVfs,
  converter: (bytes: Uint8Array, signal?: AbortSignal) => Promise<string>,
  signal?: AbortSignal,
) {
  if (args.length !== 2) return { output: '', error: 'sandbox_denied' }
  const output = await converter(vfs.readBytes(args[0]), signal)
  check(signal)
  vfs.writeFile(args[1], output)
  return { output: args[1] }
}
