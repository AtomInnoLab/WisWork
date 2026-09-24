import { XMLParser, XMLValidator } from 'fast-xml-parser'
import JSZip from 'jszip'
import { loadBoundedZip, MAX_PPTX_XML_BYTES } from './powerpoint-package.js'

type Node = Record<string, unknown>
const parser = new XMLParser({ preserveOrder: true, ignoreAttributes: false, attributeNamePrefix: '@_', parseTagValue: false, parseAttributeValue: false, processEntities: false })
const children = (nodes: Node[], tag: string): Node[][] => nodes.flatMap((node) => Object.entries(node).flatMap(([key, value]) => Array.isArray(value) ? [...(key === tag ? [value as Node[]] : []), ...children(value as Node[], tag)] : []))
const tags = (nodes: Node[], tag: string): Node[] => nodes.flatMap((node) => Object.entries(node).flatMap(([key, value]) => Array.isArray(value) ? [...(key === tag ? [node] : []), ...tags(value as Node[], tag)] : []))
const attr = (node: Node, name: string): string | undefined => { const value = (node[':@'] as Node | undefined)?.[`@_${name}`]; return typeof value === 'string' ? value : undefined }
const text = (nodes: Node[]): string => nodes.map((node) => node['#text'] ?? '').join('')
const firstText = (nodes: Node[], tag: string): string | undefined => { const found = children(nodes, tag)[0]; return found ? text(found) : undefined }
function xml(source: string): Node[] {
  if (new TextEncoder().encode(source).byteLength > MAX_PPTX_XML_BYTES || /<!\s*(?:DOCTYPE|ENTITY)\b/i.test(source) || XMLValidator.validate(source) !== true) throw new Error('office_api_unsupported')
  return parser.parse(source) as Node[]
}
async function read(zip: JSZip, path: string, signal?: AbortSignal): Promise<Node[]> {
  if (signal?.aborted) throw new Error('cancelled')
  const file = zip.file(path)
  if (!file) throw new Error('office_api_unsupported')
  return xml(await file.async('string'))
}
function relation(nodes: Node[], id: string, type: string): Node {
  const all = tags(nodes, 'Relationship')
  if (new Set(all.map((node) => attr(node, 'Id'))).size !== all.length) throw new Error('office_api_unsupported')
  const found = all.filter((node) => attr(node, 'Id') === id)
  if (found.length !== 1 || !attr(found[0]!, 'Type')?.endsWith(`/${type}`)) throw new Error('office_api_unsupported')
  return found[0]!
}
function points(nodes: Node[], container: string): string[] | undefined {
  const group = children(nodes, container)[0]
  if (!group) return undefined
  const cache = children(group, 'c:strCache')[0] ?? children(group, 'c:numCache')[0]
  if (!cache || children(group, 'c:multiLvlStrCache').length) return undefined
  const all = tags(cache, 'c:pt')
  if (all.length > 32 || all.some((node, index) => attr(node, 'idx') !== String(index))) return undefined
  return all.map((node) => firstText(node['c:pt'] as Node[], 'c:v') ?? '')
}
function formula(nodes: Node[], container: string): string | undefined { const group = children(nodes, container)[0]; return group && firstText(group, 'c:f') }
function range(formulaText: string): { col: string; start: number; end: number } | undefined {
  const match = /^(?:Sheet1|'Sheet1')!\$?([A-Z]{1,2})\$?([1-9]\d*):\$?([A-Z]{1,2})\$?([1-9]\d*)$/.exec(formulaText)
  if (!match || match[1] !== match[3]) return undefined
  const start = Number(match[2]), end = Number(match[4])
  return end >= start && end - start < 32 && end <= 100000 ? { col: match[1]!, start, end } : undefined
}
function cellValues(sheet: Node[], shared: string[]): Map<string, string> {
  const cells = tags(sheet, 'c')
  if (cells.length > 4096) throw new Error('office_api_unsupported')
  const result = new Map<string, string>()
  for (const cell of cells) {
    const address = attr(cell, 'r')
    if (!address || !/^[A-Z]{1,2}[1-9]\d*$/.test(address) || result.has(address)) throw new Error('office_api_unsupported')
    const body = cell['c'] as Node[]
    const kind = attr(cell, 't')
    const raw = kind === 'inlineStr' ? firstText(body, 't') : firstText(body, 'v')
    if (raw === undefined || raw.length > 128 || tags(body, 'f').length) throw new Error('office_api_unsupported')
    if (kind === 's') {
      const index = Number(raw)
      if (!Number.isSafeInteger(index) || index < 0 || index >= shared.length) throw new Error('office_api_unsupported')
      result.set(address, shared[index]!)
    } else if (kind === 'inlineStr' || kind === 'str' || kind === undefined || kind === 'n') result.set(address, raw)
    else throw new Error('office_api_unsupported')
  }
  return result
}
export interface ChartSourceReport {
  shapeId: string
  sourceKind: 'embedded_xlsx' | 'external_link' | 'cache_only' | 'unsupported'
  verification: 'matches' | 'mismatch' | 'not_verified'
  reason?: string
  sourceDigest?: string
  series: Array<{ categories: string[]; values: string[] }>
}
/** Read a single chart's bounded cache and embedded workbook; external targets are classified, never fetched. */
export async function inspectPowerPointChartSourcePackage(base64: string, shapeId: string, signal?: AbortSignal): Promise<ChartSourceReport> {
  if (!/^[1-9]\d{0,9}$/.test(shapeId)) throw new Error('invalid_tool_input')
  const zip = await loadBoundedZip(base64, signal)
  const slides = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  if (slides.length !== 1) throw new Error('office_api_unsupported')
  const slidePath = slides[0]!
  const slide = await read(zip, slidePath, signal)
  const frames = children(slide, 'p:graphicFrame').filter((frame) => tags(frame, 'p:cNvPr').some((node) => attr(node, 'id') === shapeId))
  if (frames.length !== 1) throw new Error('office_api_unsupported')
  const chartNodes = tags(frames[0]!, 'c:chart')
  if (chartNodes.length !== 1) throw new Error('office_api_unsupported')
  const chartId = attr(chartNodes[0]!, 'r:id')
  if (!chartId) throw new Error('office_api_unsupported')
  const slideRels = await read(zip, slidePath.replace('/slides/', '/slides/_rels/') + '.rels', signal)
  const chartRel = relation(slideRels, chartId, 'chart')
  if (attr(chartRel, 'TargetMode') !== undefined || !/^\.\.\/charts\/chart\d+\.xml$/.test(attr(chartRel, 'Target') ?? '')) throw new Error('office_api_unsupported')
  const chartPath = `ppt/charts/${attr(chartRel, 'Target')!.slice('../charts/'.length)}`
  const chart = await read(zip, chartPath, signal)
  const allSeries = children(chart, 'c:ser')
  if (allSeries.length > 8) throw new Error('office_api_unsupported')
  const series = allSeries.map((item) => ({ categories: points(item, 'c:cat') ?? [], values: points(item, 'c:val') ?? [] }))
  const result: ChartSourceReport = { shapeId, sourceKind: 'cache_only', verification: 'not_verified', series }
  if (JSON.stringify(result).length > 128 * 1024) throw new Error('office_api_unsupported')
  const externalData = tags(chart, 'c:externalData')
  if (!externalData.length) return result
  if (externalData.length !== 1) throw new Error('office_api_unsupported')
  const sourceId = attr(externalData[0]!, 'r:id')
  if (!sourceId) throw new Error('office_api_unsupported')
  const chartRels = await read(zip, chartPath.replace('/charts/', '/charts/_rels/') + '.rels', signal)
  const sourceRel = relation(chartRels, sourceId, 'package')
  if (attr(sourceRel, 'TargetMode') === 'External') return { ...result, sourceKind: 'external_link', reason: 'external_source_not_fetched' }
  if (attr(sourceRel, 'TargetMode') !== undefined || !/^\.\.\/embeddings\/[A-Za-z0-9_.-]+\.xlsx$/.test(attr(sourceRel, 'Target') ?? '')) throw new Error('office_api_unsupported')
  const sourcePath = `ppt/embeddings/${attr(sourceRel, 'Target')!.slice('../embeddings/'.length)}`
  const bytes = await zip.file(sourcePath)?.async('uint8array')
  if (!bytes || bytes.byteLength > 2 * 1024 * 1024) throw new Error('office_api_unsupported')
  const digest = await crypto.subtle.digest('SHA-256', new Uint8Array(bytes))
  const sourceDigest = Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('')
  const embedded: ChartSourceReport = { ...result, sourceKind: 'embedded_xlsx', sourceDigest }
  const formulaPairs = allSeries.map((item) => [formula(item, 'c:cat'), formula(item, 'c:val')] as const)
  const ranges = formulaPairs.map(([cat, val]) => [cat ? range(cat) : undefined, val ? range(val) : undefined] as const)
  if (!series.length || ranges.some(([cat, val]) => !cat || !val) || series.some((item) => !item.categories.length || !item.values.length))
    return { ...embedded, reason: 'unsupported_formula_or_cache' }
  const book = await loadBoundedZip(btoa(Array.from(bytes, (byte) => String.fromCharCode(byte)).join('')), signal)
  const workbook = await read(book, 'xl/workbook.xml', signal)
  const sheets = tags(workbook, 'sheet')
  if (sheets.length !== 1 || attr(sheets[0]!, 'name') !== 'Sheet1') return { ...embedded, reason: 'unsupported_workbook_structure' }
  const sheetId = attr(sheets[0]!, 'r:id')
  if (!sheetId) throw new Error('office_api_unsupported')
  const workbookRels = await read(book, 'xl/_rels/workbook.xml.rels', signal)
  const sheetRel = relation(workbookRels, sheetId, 'worksheet')
  if (attr(sheetRel, 'TargetMode') !== undefined || !/^worksheets\/sheet\d+\.xml$/.test(attr(sheetRel, 'Target') ?? '')) throw new Error('office_api_unsupported')
  const sheet = await read(book, `xl/${attr(sheetRel, 'Target')}`, signal)
  const sharedXml = book.file('xl/sharedStrings.xml') ? await read(book, 'xl/sharedStrings.xml', signal) : []
  const shared = tags(sharedXml, 'si').map((node) => children(node['si'] as Node[], 't').map(text).join(''))
  if (shared.length > 4096 || shared.some((item) => item.length > 128)) throw new Error('office_api_unsupported')
  const cells = cellValues(sheet, shared)
  const fromRange = (item: {col:string;start:number;end:number}) => Array.from({length:item.end-item.start+1}, (_, index) => cells.get(`${item.col}${item.start+index}`))
  const compares = ranges.map(([cat, val], index) => {
    const categories = fromRange(cat!), values = fromRange(val!)
    if (categories.some((value) => value === undefined) || values.some((value) => value === undefined)) return undefined
    return JSON.stringify(categories) === JSON.stringify(series[index]!.categories) && JSON.stringify(values) === JSON.stringify(series[index]!.values)
  })
  if (signal?.aborted) throw new Error('cancelled')
  return { ...embedded, verification: compares.includes(undefined) ? 'not_verified' : compares.every(Boolean) ? 'matches' : 'mismatch', ...(compares.includes(undefined) ? {reason:'missing_source_cell'} : {}) }
}
