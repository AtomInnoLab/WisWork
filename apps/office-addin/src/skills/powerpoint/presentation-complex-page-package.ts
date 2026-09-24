import { XMLParser, XMLValidator } from 'fast-xml-parser'
import { loadBoundedZip, MAX_PPTX_XML_BYTES } from './powerpoint-package.js'

type Node = Record<string, unknown>
const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: false,
  processEntities: false,
})

function elements(nodes: Node[], tag: string): Node[][] {
  const found: Node[][] = []
  for (const node of nodes) {
    for (const [key, value] of Object.entries(node)) {
      if (!Array.isArray(value)) continue
      if (key === tag) found.push(value as Node[])
      found.push(...elements(value as Node[], tag))
    }
  }
  return found
}

function value(nodes: Node[]): string {
  return elements(nodes, 'a:t').map((parts) => parts.map((part) => part['#text'] ?? '').join('')).join('')
}

function xmlNodes(xml: string): Node[] {
  if (
    new TextEncoder().encode(xml).byteLength > MAX_PPTX_XML_BYTES ||
    /<!\s*(?:DOCTYPE|ENTITY)\b/i.test(xml) ||
    XMLValidator.validate(xml) !== true
  ) throw new Error('office_api_unsupported')
  return parser.parse(xml) as Node[]
}

export interface ComplexPagePackageSummary {
  tables: Array<{ shapeId: string; rows: string[][]; simpleCells: boolean[][]; truncated?: boolean }>
  charts: Array<{
    shapeId: string
    series: Array<{ name?: string; categories: string[]; values: string[] }>
    cacheOnly: true
    truncated?: boolean
  }>
  truncated: boolean
}

/** Read only the exported page and embedded chart caches; workbook links are never followed. */
export async function inspectPowerPointComplexPagePackage(
  base64: string,
  signal?: AbortSignal,
): Promise<ComplexPagePackageSummary> {
  const zip = await loadBoundedZip(base64, signal)
  const slides = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  if (slides.length !== 1) throw new Error('office_api_unsupported')
  const slidePath = slides[0]!
  const slide = xmlNodes(await zip.file(slidePath)!.async('string'))
  const frames = elements(slide, 'p:graphicFrame')
  const output: ComplexPagePackageSummary = { tables: [], charts: [], truncated: false }
  // Fixed text budget also covers JSON escaping and the bounded array structure below.
  let remainingText = 8 * 1024
  const clip = (text: string, limit = 128): string => {
    const length = Math.min(text.length, limit, remainingText)
    if (text.length > length) output.truncated = true
    remainingText -= length
    return text.slice(0, length)
  }
  const relPath = slidePath.replace('/slides/', '/slides/_rels/') + '.rels'
  let relationships: Node[] | undefined
  const findChart = async (id: string): Promise<Node[]> => {
    if (!relationships) {
      const file = zip.file(relPath)
      if (!file) throw new Error('office_api_unsupported')
      relationships = xmlNodes(await file.async('string'))
    }
    const relNodes = relationshipNodes(relationships).filter((node) => (node[':@'] as Node | undefined)?.['@_Id'] === id)
    if (relNodes.length !== 1) throw new Error('office_api_unsupported')
    const attrs = relNodes[0]![':@'] as Node
    const target = attrs['@_Target']
    const type = attrs['@_Type']
    if (
      attrs['@_TargetMode'] === 'External' ||
      typeof target !== 'string' ||
      !/^\.\.\/charts\/chart\d+\.xml$/.test(target) ||
      typeof type !== 'string' ||
      !type.endsWith('/chart')
    ) throw new Error('office_api_unsupported')
    const file = zip.file(`ppt/charts/${target.slice('../charts/'.length)}`)
    if (!file) throw new Error('office_api_unsupported')
    return xmlNodes(await file.async('string'))
  }
  for (const frame of frames) {
    if (signal?.aborted) throw new Error('cancelled')
    const rawShapeId = frameId(frame)
    const shapeId = rawShapeId ? clip(rawShapeId, 64) : undefined
    if (!shapeId) continue
    const tables = elements(frame, 'a:tbl')
    if (tables.length) {
      if (output.tables.length >= 16) { output.truncated = true; continue }
      const hasMergedCells = tagNodes(tables[0]!, 'a:tc').some((node) => {
        const attrs = node[':@'] as Node | undefined
        return attrs && ['@_gridSpan', '@_rowSpan', '@_hMerge', '@_vMerge'].some((name) => Object.hasOwn(attrs, name))
      })
      const allRows = elements(tables[0]!, 'a:tr')
      const simpleCells: boolean[][] = []
      const rows = allRows.slice(0, 20).map((row) => {
        const cells = elements(row, 'a:tc')
        if (cells.length > 12) output.truncated = true
        const selected = cells.slice(0, 12)
        simpleCells.push(selected.map((cell) =>
          !hasMergedCells && elements(cell, 'a:p').length <= 1 &&
          elements(cell, 'a:r').length <= 1 &&
          elements(cell, 'a:br').length === 0 &&
          elements(cell, 'a:fld').length === 0 &&
          elements(cell, 'a:rPr').length === 0 &&
          elements(cell, 'a:pPr').length === 0 &&
          elements(cell, 'a:endParaRPr').length === 0,
        ))
        return selected.map((cell) => clip(value(cell)))
      })
      if (allRows.length > 20) output.truncated = true
      output.tables.push({ shapeId, rows, simpleCells, ...(output.truncated ? { truncated: true } : {}) })
      continue
    }
    const chartTags = chartNodes(frame)
    if (!chartTags.length) continue
    if (output.charts.length >= 16) { output.truncated = true; continue }
    const id = (chartTags[0]![':@'] as Node | undefined)?.['@_r:id']
    if (typeof id !== 'string') throw new Error('office_api_unsupported')
    const chart = await findChart(id)
    const allSeries = elements(chart, 'c:ser')
    const series = allSeries.slice(0, 8).map((ser) => {
      const cache = (container: string): string[] => {
        const groups = elements(ser, container)
        if (!groups.length) return []
        if (elements(groups[0]!, 'c:multiLvlStrCache').length) output.truncated = true
        const cached = elements(groups[0]!, 'c:strCache')[0] ?? elements(groups[0]!, 'c:numCache')[0]
        if (!cached) return []
        const points = tagNodes(cached, 'c:pt')
        if (points.length > 32) output.truncated = true
        return points.slice(0, 32).map((point, index) => {
          if ((point[':@'] as Node | undefined)?.['@_idx'] !== String(index)) output.truncated = true
          return clip(elements(point['c:pt'] as Node[], 'c:v')[0]?.map((item) => item['#text'] ?? '').join('') ?? '')
        })
      }
      const tx = elements(ser, 'c:tx')[0] ?? []
      const name = elements(tx, 'c:v')[0]?.map((item) => item['#text'] ?? '').join('')
      return { ...(name ? { name: clip(name) } : {}), categories: cache('c:cat'), values: cache('c:val') }
    })
    if (allSeries.length > 8) output.truncated = true
    output.charts.push({ shapeId, series, cacheOnly: true, ...(output.truncated ? { truncated: true } : {}) })
  }
  return output
}

function relationshipNodes(nodes: Node[]): Node[] {
  return tagNodes(nodes, 'Relationship')
}

function chartNodes(nodes: Node[]): Node[] {
  return tagNodes(nodes, 'c:chart')
}

function tagNodes(nodes: Node[], tag: string): Node[] {
  return nodes.flatMap((node) => [
    ...(Object.hasOwn(node, tag) ? [node] : []),
    ...Object.values(node).filter(Array.isArray).flatMap((parts) => tagNodes(parts as Node[], tag)),
  ])
}

function frameId(nodes: Node[]): string | undefined {
  for (const node of nodes) {
    if (Object.hasOwn(node, 'p:cNvPr')) {
      const id = (node[':@'] as Node | undefined)?.['@_id']
      return typeof id === 'string' ? id : undefined
    }
    for (const parts of Object.values(node).filter(Array.isArray)) {
      const id = frameId(parts as Node[])
      if (id) return id
    }
  }
  return undefined
}

/** Exact-cell evidence for reversible text edits. The digest excludes only text payloads. */
export async function inspectPowerPointTableCellPackage(
  base64: string,
  shapeId: string,
  rowIndex: number,
  columnIndex: number,
  signal?: AbortSignal,
): Promise<{ text: string; structureDigest: string }> {
  if (signal?.aborted) throw new Error('cancelled')
  if (!Number.isSafeInteger(rowIndex) || rowIndex < 0 || !Number.isSafeInteger(columnIndex) || columnIndex < 0)
    throw new Error('invalid_tool_input')
  const zip = await loadBoundedZip(base64, signal)
  const slides = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  if (slides.length !== 1) throw new Error('office_api_unsupported')
  const slide = xmlNodes(await zip.file(slides[0]!)!.async('string'))
  const frames = elements(slide, 'p:graphicFrame').filter((item) => frameId(item) === shapeId)
  if (frames.length !== 1) throw new Error('presentation_existing_target_unsupported')
  const tables = elements(frames[0]!, 'a:tbl')
  if (!tables || tables.length !== 1) throw new Error('presentation_existing_target_unsupported')
  if (tagNodes(tables[0]!, 'a:tc').some((node) => {
    const attrs = node[':@'] as Node | undefined
    return attrs && ['@_gridSpan', '@_rowSpan', '@_hMerge', '@_vMerge'].some((name) => Object.hasOwn(attrs, name))
  })) throw new Error('presentation_existing_target_unsupported')
  const rows = elements(tables[0]!, 'a:tr')
  const cell = rows[rowIndex] && elements(rows[rowIndex]!, 'a:tc')[columnIndex]
  if (!cell || elements(cell, 'a:p').length > 1 || elements(cell, 'a:r').length > 1 ||
    ['a:br', 'a:fld'].some((tag) => elements(cell, tag).length))
    throw new Error('presentation_existing_target_unsupported')
  const text = value(cell)
  if (text.length > 128) throw new Error('presentation_existing_target_unsupported')
  // Bind the exact coordinate to the whole table; only the target cell text may change.
  const tableForDigest = structuredClone(tables[0]!)
  const digestRow = elements(tableForDigest, 'a:tr')[rowIndex]
  const digestCell = digestRow && elements(digestRow, 'a:tc')[columnIndex]
  if (!digestCell) throw new Error('presentation_existing_target_unsupported')
  const blankTargetText = (nodes: Node[]) => {
    for (const node of nodes) for (const [key, entry] of Object.entries(node)) {
      if (key === 'a:t') node[key] = []
      else if (Array.isArray(entry)) blankTargetText(entry as Node[])
    }
  }
  blankTargetText(digestCell)
  const hash = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(JSON.stringify(tableForDigest)))
  if (signal?.aborted) throw new Error('cancelled')
  return { text, structureDigest: Array.from(new Uint8Array(hash), (byte) => byte.toString(16).padStart(2, '0')).join('') }
}
