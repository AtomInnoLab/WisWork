import { XMLParser, XMLValidator } from 'fast-xml-parser'
import type JSZip from 'jszip'
import type { PowerPointPageInspection } from './browser-powerpoint-adapter.js'
import {
  inspectPowerPointPictureMediaBatch,
  loadBoundedZip,
  MAX_PPTX_XML_BYTES,
} from './powerpoint-package.js'
import { inspectPowerPointComplexPagePackage } from './presentation-complex-page-package.js'
import { inspectPowerPointChartSourcesBatch } from './presentation-chart-source-package.js'
import { inspectPowerPointSourceLinksFromZip } from './presentation-source-links-package.js'

type Issue =
  | { name: string; kind: 'missing' | 'extra' | 'duplicate' }
  | { name: string; kind: 'type_changed'; sourceType: string; hostType: string }
  | { name: string; kind: 'geometry_changed' }
  | { name: string; kind: 'rotation_changed' }
type Xml = Record<string, any>
type SourceObject = {
  name: string
  shapeId: string
  type: string
  box: [number, number, number, number]
  text: string[]
  textStyles?: Array<[string, string, string, string]>
  tableCellStyles?: Array<{
    fill: string
    borders: Array<[string, string]>
    runs: Array<[string, string, string, string]>
  }>
  tableStructure?: {
    columns: number[]
    rows: number[]
    cellCounts: number[]
    merges: Array<[number, number, number, number]>
  }
  altText?: string
  crop?: [number, number, number, number]
  appearance?: [string, string, string]
  rotation: number
}
type SourceLink = { name: string; target: string; label: string; location: string }
const POINTS_PER_EMU = 72 / 914400
const types: Record<string, string[]> = {
  shape: ['TextBox', 'GeometricShape'],
  picture: ['Image'],
  table: ['Table'],
  chart: ['Chart'],
}
const parser = new XMLParser({
  ignoreAttributes: false,
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
})
const SOLID_HEX = /^[0-9A-Fa-f]{6}$/
const RELATIONSHIP_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'

async function readXml(zip: JSZip, path: string): Promise<Xml | undefined> {
  const file = zip.file(path)
  if (!file) return undefined
  const xml = await file.async('string')
  if (
    new TextEncoder().encode(xml).byteLength > MAX_PPTX_XML_BYTES ||
    /<!\s*(?:DOCTYPE|ENTITY)\b/i.test(xml) ||
    XMLValidator.validate(xml) !== true
  )
    return undefined
  return parser.parse(xml) as Xml
}

function directBackground(
  root: Xml,
  tag: 'p:sld' | 'p:sldLayout' | 'p:sldMaster',
): string | null | undefined {
  const bg = root[tag]?.['p:cSld']?.['p:bg']
  if (bg === undefined) return undefined
  const color = bg?.['p:bgPr']?.['a:solidFill']?.['a:srgbClr']?.['@_val']
  return typeof color === 'string' && SOLID_HEX.test(color) ? color.toUpperCase() : null
}

async function linkedPart(
  zip: JSZip,
  path: string,
  type: 'slideLayout' | 'slideMaster',
): Promise<string | undefined> {
  const slash = path.lastIndexOf('/')
  const rels = await readXml(zip, `${path.slice(0, slash)}/_rels/${path.slice(slash + 1)}.rels`)
  const relations = many(rels?.Relationships?.Relationship).filter(
    (item) => item['@_Type'] === `${RELATIONSHIP_BASE}${type}`,
  )
  if (relations.length !== 1 || relations[0]?.['@_TargetMode'] !== undefined) return undefined
  const target = relations[0]?.['@_Target']
  const folder = type === 'slideLayout' ? 'slideLayouts' : 'slideMasters'
  const file = type === 'slideLayout' ? 'slideLayout' : 'slideMaster'
  const match =
    typeof target === 'string' &&
    new RegExp(`^(?:\\.\\./${folder}/|/ppt/${folder}/)(${file}\\d+\\.xml)$`).exec(target)
  return match ? `ppt/${folder}/${match[1]}` : undefined
}

async function resolvedBackground(
  zip: JSZip,
  slidePath: string,
  root: Xml,
): Promise<string | undefined> {
  const slideColor = directBackground(root, 'p:sld')
  if (slideColor !== undefined) return slideColor ?? undefined
  const layoutPath = await linkedPart(zip, slidePath, 'slideLayout')
  if (!layoutPath) return undefined
  const layout = await readXml(zip, layoutPath)
  if (!layout) return undefined
  const layoutColor = directBackground(layout, 'p:sldLayout')
  if (layoutColor !== undefined) return layoutColor ?? undefined
  const masterPath = await linkedPart(zip, layoutPath, 'slideMaster')
  if (!masterPath) return undefined
  const master = await readXml(zip, masterPath)
  return master ? (directBackground(master, 'p:sldMaster') ?? undefined) : undefined
}

async function pageNotes(zip: JSZip, slidePath: string): Promise<string | undefined> {
  const slash = slidePath.lastIndexOf('/')
  const relsPath = `${slidePath.slice(0, slash)}/_rels/${slidePath.slice(slash + 1)}.rels`
  if (!zip.file(relsPath)) return ''
  const rels = await readXml(zip, relsPath)
  if (!rels) return undefined
  const matches = many(rels.Relationships?.Relationship).filter(
    (item) => item['@_Type'] === `${RELATIONSHIP_BASE}notesSlide`,
  )
  if (!matches.length) return ''
  if (matches.length !== 1 || matches[0]?.['@_TargetMode'] !== undefined) return undefined
  const target = matches[0]?.['@_Target']
  const match =
    typeof target === 'string' &&
    /^(?:\.\.\/notesSlides\/|\/ppt\/notesSlides\/)(notesSlide\d+\.xml)$/.exec(target)
  if (!match) return undefined
  const root = await readXml(zip, `ppt/notesSlides/${match[1]}`)
  if (!root) return undefined
  const shapes = many(root['p:notes']?.['p:cSld']?.['p:spTree']?.['p:sp']).filter(
    (shape) => shape['p:nvSpPr']?.['p:nvPr']?.['p:ph']?.['@_type'] === 'body',
  )
  if (shapes.length !== 1) return undefined
  const paragraphs = many(shapes[0]?.['p:txBody']?.['a:p']).map((paragraph) =>
    textRuns(paragraph).join(''),
  )
  while (paragraphs.length && paragraphs.at(-1) === '') paragraphs.pop()
  const notes = paragraphs.join('\n')
  return notes.length <= 12_000 ? notes : undefined
}
const many = (value: unknown): Xml[] =>
  value === undefined ? [] : Array.isArray(value) ? value : [value as Xml]

function textRuns(value: unknown): string[] {
  if (Array.isArray(value)) return value.flatMap(textRuns)
  if (!value || typeof value !== 'object') return []
  return Object.entries(value).flatMap(([tag, child]) =>
    tag === 'a:t'
      ? [typeof child === 'string' ? child : String((child as Xml)?.['#text'] ?? '')]
      : tag === 'a:br'
        ? ['\n']
        : tag.startsWith('@_')
          ? []
          : textRuns(child),
  )
}

function textStyles(body: Xml | undefined): Array<[string, string, string, string]> {
  return many(body?.['a:p']).flatMap((paragraph) =>
    many(paragraph['a:r']).map((run) => {
      const style = run['a:rPr']
      return [
        String(style?.['@_sz'] ?? ''),
        String(style?.['@_b'] ?? ''),
        String(style?.['a:latin']?.['@_typeface'] ?? ''),
        String(style?.['a:solidFill']?.['a:srgbClr']?.['@_val'] ?? '').toUpperCase(),
      ] as [string, string, string, string]
    }),
  )
}

/** Styles not represented by the existing RGB/width comparison cannot certify equality. */
function supportedTableBorders(cells: Xml[][]): void {
  const unavailable = (): never => {
    throw new Error('presentation_qa_structure_unavailable')
  }
  for (const row of cells)
    for (const cell of row) {
      const properties = cell['a:tcPr']
      if (properties?.['a:lnTlToBr'] !== undefined || properties?.['a:lnBlToTr'] !== undefined)
        unavailable()
      for (const side of ['a:lnL', 'a:lnR', 'a:lnT', 'a:lnB']) {
        const line = properties?.[side]
        if (line === undefined) continue
        if (
          !line ||
          typeof line !== 'object' ||
          Array.isArray(line) ||
          Object.keys(line).some(
            (key) =>
              ![
                '@_w',
                '@_cap',
                '@_cmpd',
                '@_algn',
                'a:solidFill',
                'a:noFill',
                'a:prstDash',
                'a:round',
                'a:headEnd',
                'a:tailEnd',
              ].includes(key),
          ) ||
          typeof line['@_w'] !== 'string' ||
          !/^[0-9]+$/.test(line['@_w']) ||
          Number(line['@_w']) > 91440000 ||
          (line['@_cap'] !== undefined && line['@_cap'] !== 'flat') ||
          (line['@_cmpd'] !== undefined && line['@_cmpd'] !== 'sng') ||
          (line['@_algn'] !== undefined && line['@_algn'] !== 'ctr') ||
          (line['a:prstDash'] !== undefined &&
            (Object.keys(line['a:prstDash']).join(',') !== '@_val' ||
              line['a:prstDash']['@_val'] !== 'solid')) ||
          (line['a:round'] !== undefined && line['a:round'] !== '')
        )
          unavailable()
        for (const end of ['a:headEnd', 'a:tailEnd']) {
          const marker = line[end]
          if (
            marker !== undefined &&
            (typeof marker !== 'object' ||
              Array.isArray(marker) ||
              Object.keys(marker).some((key) => !['@_type', '@_w', '@_len'].includes(key)) ||
              marker['@_type'] !== 'none' ||
              (marker['@_w'] !== undefined && marker['@_w'] !== 'med') ||
              (marker['@_len'] !== undefined && marker['@_len'] !== 'med'))
          )
            unavailable()
        }
        if (line['a:noFill'] !== undefined) {
          if (line['a:noFill'] !== '' || line['a:solidFill'] !== undefined) unavailable()
        } else {
          const fill = line['a:solidFill'],
            color = fill?.['a:srgbClr']
          if (
            !fill ||
            Object.keys(fill).join(',') !== 'a:srgbClr' ||
            !color ||
            Object.keys(color).join(',') !== '@_val' ||
            typeof color['@_val'] !== 'string' ||
            !SOLID_HEX.test(color['@_val'])
          )
            unavailable()
        }
      }
    }
}

/** Physical continuation cells occupy the grid; their spans must not be summed as columns. */
function tableStructure(element: Xml): NonNullable<SourceObject['tableStructure']> {
  const table = element['a:graphic']?.['a:graphicData']?.['a:tbl']
  const unavailable = (): never => {
    throw new Error('presentation_qa_structure_unavailable')
  }
  const integer = (value: unknown, maximum: number): number => {
    if (typeof value !== 'string' || !/^[0-9]+$/.test(value)) return unavailable()
    const number = Number(value)
    if (!Number.isSafeInteger(number) || number < 1 || number > maximum) return unavailable()
    return number
  }
  const flag = (value: unknown): boolean => {
    if (value === undefined || value === '0' || value === 'false') return false
    if (value === '1' || value === 'true') return true
    return unavailable()
  }
  if (!table || /"a:extLst":/.test(JSON.stringify(table))) return unavailable()
  const columns = many(table['a:tblGrid']?.['a:gridCol']).map((column) =>
    integer(column['@_w'], 91440000),
  )
  const rows = many(table['a:tr'])
  if (
    !columns.length ||
    columns.length > 32 ||
    !rows.length ||
    rows.length > 64 ||
    rows.length * columns.length > 2048
  )
    return unavailable()
  const heights = rows.map((row) => integer(row['@_h'], 91440000))
  const cells = rows.map((row) => many(row['a:tc']))
  if (cells.some((row) => row.length !== columns.length)) return unavailable()
  supportedTableBorders(cells)
  const occupied: Array<Array<[number, number, number, number] | undefined>> = rows.map(() =>
    Array(columns.length),
  )
  const merges: Array<[number, number, number, number]> = []
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < columns.length; c++) {
      const cell = cells[r]![c]!
      const width =
        cell['@_gridSpan'] === undefined ? 1 : integer(cell['@_gridSpan'], columns.length)
      const height = cell['@_rowSpan'] === undefined ? 1 : integer(cell['@_rowSpan'], rows.length)
      const horizontal = flag(cell['@_hMerge']),
        vertical = flag(cell['@_vMerge'])
      const owner = occupied[r]![c]
      if (owner) {
        const [startRow, startColumn, rowSpan, columnSpan] = owner
        if (
          horizontal !== c > startColumn ||
          vertical !== r > startRow ||
          (width !== 1 && (c !== startColumn || width !== columnSpan)) ||
          (height !== 1 && (r !== startRow || height !== rowSpan))
        )
          return unavailable()
        continue
      }
      if (horizontal || vertical || r + height > rows.length || c + width > columns.length)
        return unavailable()
      const rectangle: [number, number, number, number] = [r, c, height, width]
      for (let y = r; y < r + height; y++) {
        for (let x = c; x < c + width; x++) {
          if (occupied[y]![x]) return unavailable()
          occupied[y]![x] = rectangle
        }
      }
      if (height > 1 || width > 1) merges.push(rectangle)
    }
  }
  return { columns, rows: heights, cellCounts: cells.map((row) => row.length), merges }
}

function tableCellStyles(element: Xml): SourceObject['tableCellStyles'] {
  const table = element['a:graphic']?.['a:graphicData']?.['a:tbl']
  return many(table?.['a:tr']).flatMap((row) =>
    many(row['a:tc']).map((cell) => ({
      fill: String(cell['a:tcPr']?.['a:solidFill']?.['a:srgbClr']?.['@_val'] ?? '').toUpperCase(),
      borders: (['a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'] as const).map((side) => [
        String(cell['a:tcPr']?.[side]?.['@_w'] ?? ''),
        String(
          cell['a:tcPr']?.[side]?.['a:solidFill']?.['a:srgbClr']?.['@_val'] ?? '',
        ).toUpperCase(),
      ]),
      runs: textStyles(cell['a:txBody']),
    })),
  )
}

function sourceObjects(root: Xml): SourceObject[] {
  const tree = root?.['p:sld']?.['p:cSld']?.['p:spTree'] as Xml | undefined
  if (!tree) throw new Error('presentation_qa_structure_unavailable')
  if (['p:grpSp', 'p:cxnSp', 'p:contentPart', 'mc:AlternateContent'].some((tag) => tree[tag]))
    throw new Error('presentation_qa_structure_unavailable')
  const result: SourceObject[] = []
  for (const [tag, kind, nonVisual] of [
    ['p:sp', 'shape', 'p:nvSpPr'],
    ['p:pic', 'picture', 'p:nvPicPr'],
    ['p:graphicFrame', 'graphic', 'p:nvGraphicFramePr'],
  ]) {
    for (const element of many(tree[tag])) {
      const name = element[nonVisual]?.['p:cNvPr']?.['@_name']
      const shapeId = element[nonVisual]?.['p:cNvPr']?.['@_id']
      const xfrm = tag === 'p:graphicFrame' ? element['p:xfrm'] : element['p:spPr']?.['a:xfrm']
      const box = [
        xfrm?.['a:off']?.['@_x'],
        xfrm?.['a:off']?.['@_y'],
        xfrm?.['a:ext']?.['@_cx'],
        xfrm?.['a:ext']?.['@_cy'],
      ].map(Number)
      const data = element['a:graphic']?.['a:graphicData']
      const type =
        kind !== 'graphic' ? kind : data?.['a:tbl'] ? 'table' : data?.['c:chart'] ? 'chart' : ''
      if (
        typeof name !== 'string' ||
        !name ||
        typeof shapeId !== 'string' ||
        !type ||
        box.some((value) => !Number.isFinite(value))
      )
        throw new Error('presentation_qa_structure_unavailable')
      const rect = tag === 'p:pic' ? element['p:blipFill']?.['a:srcRect'] : undefined
      const crop =
        tag === 'p:pic'
          ? ['l', 't', 'r', 'b'].map((side) => Number(rect?.[`@_${side}`] ?? 0))
          : undefined
      if (crop?.some((value) => !Number.isFinite(value)))
        throw new Error('presentation_qa_structure_unavailable')
      const rotation = Number(xfrm?.['@_rot'] ?? 0)
      if (!Number.isFinite(rotation)) throw new Error('presentation_qa_structure_unavailable')
      const properties = element['p:spPr']
      const appearance: [string, string, string] = [
        String(properties?.['a:prstGeom']?.['@_prst'] ?? ''),
        String(properties?.['a:solidFill']?.['a:srgbClr']?.['@_val'] ?? '').toUpperCase(),
        String(properties?.['a:ln']?.['a:solidFill']?.['a:srgbClr']?.['@_val'] ?? '').toUpperCase(),
      ]
      result.push({
        name,
        shapeId,
        type,
        box: box as SourceObject['box'],
        text: textRuns(element),
        ...(tag === 'p:sp' ? { textStyles: textStyles(element['p:txBody']) } : {}),
        ...(type === 'table'
          ? { tableCellStyles: tableCellStyles(element), tableStructure: tableStructure(element) }
          : {}),
        rotation,
        ...(tag === 'p:sp' ? { appearance } : {}),
        ...(tag === 'p:pic'
          ? {
              altText: String(element[nonVisual]?.['p:cNvPr']?.['@_descr'] ?? ''),
              crop: crop as [number, number, number, number],
            }
          : {}),
      })
    }
  }
  return result
}

async function readPage(
  base64: string,
  index: number,
  maxBytes: number,
): Promise<{
  objects: SourceObject[]
  backgroundColor?: string
  notesText?: string
  links?: SourceLink[]
}> {
  try {
    const zip = await loadBoundedZip(base64, undefined, true, maxBytes)
    const slidePath = `ppt/slides/slide${index + 1}.xml`
    const root = await readXml(zip, slidePath)
    if (!root) throw new Error('invalid slide XML')
    const objects = sourceObjects(root)
    if (!objects.length || objects.length > 100) throw new Error('invalid slide objects')
    const backgroundColor = await resolvedBackground(zip, slidePath, root)
    const notesText = await pageNotes(zip, slidePath)
    const extractedLinks = await inspectPowerPointSourceLinksFromZip(zip, slidePath).catch(
      () => undefined,
    )
    const links = extractedLinks?.links.map((link) => {
      const name = objects.find((item) => item.shapeId === link.packageShapeId)?.name
      if (!name) throw new Error('unmapped source link')
      return { name, target: link.target, label: link.label, location: link.location }
    })
    return {
      objects,
      ...(backgroundColor ? { backgroundColor } : {}),
      ...(notesText !== undefined ? { notesText } : {}),
      ...(links !== undefined ? { links } : {}),
    }
  } catch {
    throw new Error('presentation_qa_structure_unavailable')
  }
}

/** Compare source PPTX objects with the exact imported Office page; no host mutation. */
export async function comparePresentationPageStructure(
  sourceBase64: string,
  sourceIndex: number,
  host: PowerPointPageInspection,
  hostBase64?: string,
): Promise<{
  status: 'passed' | 'warning' | 'incomplete'
  structureStatus: 'passed' | 'warning' | 'incomplete'
  sourceCount: number
  hostCount: number
  issues: Issue[]
  readbackConsistent: boolean
  content: {
    status: 'passed' | 'warning' | 'incomplete'
    changed: string[]
    cacheChanged: string[]
    chartTypeChanged: string[]
    chartStyleChanged: string[]
    chartSourceChanged: string[]
    chartFormulaChanged: string[]
    chartSourceVerificationRegressed: string[]
    chartSourceUnreadable: string[]
    workbookBytesChanged: string[]
    workbookDataChanged: string[]
    backgroundChanged: boolean
    backgroundUnchecked: boolean
    notesChanged: boolean
    notesUnchecked: boolean
    sourceLinkChanged: string[]
    sourceLinksUnchecked: boolean
    mediaChanged: string[]
    mediaChecked: string[]
    mediaUnchecked: string[]
    altTextChanged: string[]
    cropChanged: string[]
    appearanceChanged: string[]
    textStyleChanged: string[]
    tableStyleChanged: string[]
    tableStructureChanged: string[]
    unchecked: string[]
  }
}> {
  if (
    typeof sourceBase64 !== 'string' ||
    sourceBase64.length > Math.ceil((10 * 1024 * 1024) / 3) * 4 ||
    !Number.isSafeInteger(sourceIndex) ||
    sourceIndex < 0 ||
    sourceIndex >= 32 ||
    !host ||
    !Array.isArray(host.shapes) ||
    host.shapes.length > 100
  )
    throw new Error('presentation_qa_structure_unavailable')
  const sourcePage = await readPage(sourceBase64, sourceIndex, 10 * 1024 * 1024)
  const source = sourcePage.objects
  const issues: Issue[] = []
  const hostByName = new Map<string, PowerPointPageInspection['shapes']>()
  for (const shape of host.shapes) {
    const named = hostByName.get(shape.name) ?? []
    named.push(shape)
    hostByName.set(shape.name, named)
  }
  const sourceNames = new Set<string>()
  for (const element of source) {
    if (sourceNames.has(element.name)) {
      issues.push({ name: element.name, kind: 'duplicate' })
      continue
    }
    sourceNames.add(element.name)
    const matches = hostByName.get(element.name) ?? []
    if (!matches.length) {
      issues.push({ name: element.name, kind: 'missing' })
      continue
    }
    if (matches.length !== 1) {
      issues.push({ name: element.name, kind: 'duplicate' })
      continue
    }
    const shape = matches[0]!
    if (!types[element.type]!.includes(shape.type)) {
      issues.push({
        name: element.name,
        kind: 'type_changed',
        sourceType: element.type,
        hostType: shape.type,
      })
      continue
    }
    if (
      [shape.left, shape.top, shape.width, shape.height].some(
        (value, index) =>
          !Number.isFinite(value) || Math.abs(value - element.box[index]! * POINTS_PER_EMU) > 1.5,
      )
    )
      issues.push({ name: element.name, kind: 'geometry_changed' })
  }
  for (const name of hostByName.keys())
    if (!sourceNames.has(name)) issues.push({ name, kind: 'extra' })
  const exportedPage = hostBase64 ? await readPage(hostBase64, 0, 8 * 1024 * 1024) : undefined
  const exported = exportedPage?.objects ?? []
  const readbackConsistent =
    !hostBase64 ||
    (exported.length === host.shapes.length &&
      exported.every((element) => {
        const matches = hostByName.get(element.name) ?? []
        if (matches.length !== 1 || !types[element.type]!.includes(matches[0]!.type)) return false
        return [matches[0]!.left, matches[0]!.top, matches[0]!.width, matches[0]!.height].every(
          (value, index) =>
            Number.isFinite(value) && Math.abs(value - element.box[index]! * POINTS_PER_EMU) <= 1.5,
        )
      }))
  const exportedByName = new Map(exported.map((element) => [element.name, element]))
  const backgroundUnchecked =
    !hostBase64 ||
    !readbackConsistent ||
    !sourcePage.backgroundColor ||
    !exportedPage?.backgroundColor
  const backgroundChanged =
    !backgroundUnchecked && sourcePage.backgroundColor !== exportedPage?.backgroundColor
  const notesUnchecked =
    !hostBase64 ||
    !readbackConsistent ||
    sourcePage.notesText === undefined ||
    exportedPage?.notesText === undefined
  const notesChanged = !notesUnchecked && sourcePage.notesText !== exportedPage?.notesText
  const sourceLinksUnchecked =
    !hostBase64 || !readbackConsistent || !sourcePage.links || !exportedPage?.links
  const sourceLinkChanged: string[] = []
  if (!sourceLinksUnchecked) {
    const names = new Set([
      ...sourcePage.links!.map((link) => link.name),
      ...exportedPage!.links!.map((link) => link.name),
    ])
    for (const name of names) {
      const linksFor = (links: SourceLink[]) =>
        links
          .filter((link) => link.name === name)
          .map((link) => JSON.stringify([link.location, link.label, link.target]))
          .sort()
      if (
        JSON.stringify(linksFor(sourcePage.links!)) !==
        JSON.stringify(linksFor(exportedPage!.links!))
      )
        sourceLinkChanged.push(name)
    }
  }
  if (hostBase64 && readbackConsistent)
    for (const element of source) {
      const actual = exportedByName.get(element.name)
      if (actual?.type === element.type && actual.rotation !== element.rotation)
        issues.push({ name: element.name, kind: 'rotation_changed' })
    }
  const changed: string[] = [],
    cacheChanged: string[] = [],
    chartTypeChanged: string[] = [],
    chartStyleChanged: string[] = [],
    chartSourceChanged: string[] = [],
    chartFormulaChanged: string[] = [],
    chartSourceVerificationRegressed: string[] = [],
    chartSourceUnreadable: string[] = [],
    workbookBytesChanged: string[] = [],
    workbookDataChanged: string[] = [],
    workbookUnverifiedChanged: string[] = [],
    mediaChanged: string[] = [],
    mediaChecked: string[] = [],
    mediaUnchecked: string[] = [],
    altTextChanged: string[] = [],
    cropChanged: string[] = [],
    appearanceChanged: string[] = [],
    textStyleChanged: string[] = [],
    tableStyleChanged: string[] = [],
    tableStructureChanged: string[] = [],
    unchecked: string[] = []
  if (hostBase64 && readbackConsistent) {
    for (const element of source.filter((item) => item.type === 'shape')) {
      const actual = exportedByName.get(element.name)
      if (
        actual?.type === 'shape' &&
        JSON.stringify(actual.appearance) !== JSON.stringify(element.appearance)
      )
        appearanceChanged.push(element.name)
      if (
        actual?.type === 'shape' &&
        element.textStyles?.length &&
        JSON.stringify(actual.textStyles) !== JSON.stringify(element.textStyles)
      )
        textStyleChanged.push(element.name)
    }
    for (const element of source.filter((item) => item.type === 'table')) {
      const actual = exportedByName.get(element.name)
      if (
        actual?.type === 'table' &&
        JSON.stringify(actual.tableStructure) !== JSON.stringify(element.tableStructure)
      )
        tableStructureChanged.push(element.name)
      if (
        actual?.type === 'table' &&
        JSON.stringify(actual.tableCellStyles) !== JSON.stringify(element.tableCellStyles)
      )
        tableStyleChanged.push(element.name)
    }
    const pictures = source.filter((element) => element.type === 'picture')
    for (const element of pictures) {
      const actual = exportedByName.get(element.name)
      if (!actual || actual.type !== 'picture') continue
      if (element.altText !== actual.altText) altTextChanged.push(element.name)
      if (JSON.stringify(element.crop) !== JSON.stringify(actual.crop))
        cropChanged.push(element.name)
    }
    if (pictures.length) {
      try {
        const hostPictures = pictures.flatMap((element) => {
          const hostElement = exportedByName.get(element.name)
          return hostElement?.type === 'picture' ? [hostElement] : []
        })
        const [before, after] = await Promise.all([
          inspectPowerPointPictureMediaBatch(
            sourceBase64,
            pictures.map((element) => element.shapeId),
            undefined,
            { slideIndex: sourceIndex, maxBytes: 10 * 1024 * 1024 },
          ),
          inspectPowerPointPictureMediaBatch(
            hostBase64,
            hostPictures.map((element) => element.shapeId),
          ),
        ])
        for (const element of pictures) {
          const hostElement = exportedByName.get(element.name)
          if (!hostElement || hostElement.type !== 'picture') continue
          const original = before.mediaDigests[element.shapeId]
          const current = after.mediaDigests[hostElement.shapeId]
          if (!original || !current) continue
          mediaChecked.push(element.name)
          if (original !== current) mediaChanged.push(element.name)
        }
      } catch {
        // Unsupported package structure remains unchecked; no partial pass is claimed.
      }
    }
  }
  if (hostBase64 && readbackConsistent && source.some((element) => element.type === 'chart')) {
    try {
      const [before, after] = await Promise.all([
        inspectPowerPointComplexPagePackage(sourceBase64, undefined, {
          slideIndex: sourceIndex,
          includeVisibleStyle: true,
          maxBytes: 10 * 1024 * 1024,
        }),
        inspectPowerPointComplexPagePackage(hostBase64, undefined, { includeVisibleStyle: true }),
      ])
      for (const element of source.filter((item) => item.type === 'chart')) {
        const hostElement = exportedByName.get(element.name)
        const original = before.charts.find((chart) => chart.shapeId === element.shapeId)
        const current = after.charts.find((chart) => chart.shapeId === hostElement?.shapeId)
        if (!original || !current) continue
        const effectiveStyle = (style: Record<string, string>, key: string): string | undefined => {
          if (style[key] !== undefined) return style[key]
          const point = /^(series\.\d+)\.point\.\d+\.(marker\.)?(.+)$/.exec(key)
          if (point) {
            const parent = `${point[1]}${point[2] ? '.marker' : ''}`
            const role = key.slice(0, key.length - point[3]!.length - 1)
            if (Object.keys(style).some((field) => field.startsWith(`${role}.`))) return undefined
            return style[`${parent}.${point[3]}`]
          }
          const label = /^(series\.\d+\.labels)(?:\.point\.\d+)?\.(.+)$/.exec(key)
          if (label) {
            const role = key.slice(0, key.length - label[2]!.length - 1)
            if (Object.keys(style).some((field) => field.startsWith(`${role}.`))) return undefined
            return style[`${label[1]}.${label[2]}`] ?? style[`labels.${label[2]}`]
          }
          return undefined
        }
        const explicitStyleChanged =
          original.visibleStyle &&
          current.visibleStyle &&
          [
            ...new Set([
              ...Object.keys(original.visibleStyle),
              ...Object.keys(current.visibleStyle),
            ]),
          ]
            .filter((key) => !key.endsWith('.$explicit'))
            .some((key) => {
              const beforeValue = effectiveStyle(original.visibleStyle!, key)
              const afterValue = effectiveStyle(current.visibleStyle!, key)
              return (
                beforeValue !== undefined && afterValue !== undefined && beforeValue !== afterValue
              )
            })
        if (explicitStyleChanged) chartStyleChanged.push(element.name)
        // Visible styles have an independent bounded projection; a preceding table/cache text
        // budget must not hide known style drift. Truncated cache evidence remains unverified.
        if (original.truncated || current.truncated) continue
        if (JSON.stringify(original.plotTypes) !== JSON.stringify(current.plotTypes))
          chartTypeChanged.push(element.name)
        const normalizeOptions = (value: typeof original.visualOptions) => ({
          ...value,
          valueLabels: value.valueLabels.map((label) =>
            label === 'true' ? '1' : label === 'false' ? '0' : label,
          ),
        })
        if (
          JSON.stringify(normalizeOptions(original.visualOptions)) !==
            JSON.stringify(normalizeOptions(current.visualOptions)) &&
          !chartStyleChanged.includes(element.name)
        )
          chartStyleChanged.push(element.name)
        if (
          (original.series.length || current.series.length) &&
          JSON.stringify(original.series) !== JSON.stringify(current.series)
        )
          cacheChanged.push(element.name)
      }
    } catch {
      // Unsupported chart packages stay unchecked; text and geometry still report.
    }
    const charts = source.filter((item) => item.type === 'chart')
    try {
      const [before, after] = await Promise.all([
        inspectPowerPointChartSourcesBatch(
          sourceBase64,
          charts.map((item) => item.shapeId),
          undefined,
          {
            slideIndex: sourceIndex,
            maxBytes: 10 * 1024 * 1024,
            allowAbsoluteChartTarget: true,
            includeWorkbookContentDigest: true,
          },
        ),
        inspectPowerPointChartSourcesBatch(
          hostBase64,
          charts.flatMap((item) => {
            const actual = exportedByName.get(item.name)
            return actual?.type === 'chart' ? [actual.shapeId] : []
          }),
          undefined,
          { allowAbsoluteChartTarget: true, includeWorkbookContentDigest: true },
        ),
      ])
      for (const element of charts) {
        const hostElement = exportedByName.get(element.name)
        if (!hostElement || hostElement.type !== 'chart') continue
        const original = before.reports[element.shapeId]
        const current = after.reports[hostElement.shapeId]
        if (original && !current) {
          chartSourceUnreadable.push(element.name)
          continue
        }
        if (!original || !current) continue
        if (
          original.sourceKind !== current.sourceKind ||
          (original.sourceKind === 'external_link' &&
            current.sourceKind === 'external_link' &&
            original.externalTargetDigest !== current.externalTargetDigest)
        )
          chartSourceChanged.push(element.name)
        if (
          JSON.stringify(original.formulaReferences) !== JSON.stringify(current.formulaReferences)
        )
          chartFormulaChanged.push(element.name)
        if (original.verification === 'matches' && current.verification !== 'matches')
          chartSourceVerificationRegressed.push(element.name)
        if (
          original?.sourceKind === 'embedded_xlsx' &&
          current?.sourceKind === 'embedded_xlsx' &&
          original.workbookContentDigest &&
          current.workbookContentDigest &&
          original.workbookContentDigest !== current.workbookContentDigest
        ) {
          workbookBytesChanged.push(element.name)
          if (!original.workbookDataDigest || !current.workbookDataDigest)
            workbookUnverifiedChanged.push(element.name)
        }
        if (
          original?.sourceKind === 'embedded_xlsx' &&
          current?.sourceKind === 'embedded_xlsx' &&
          original.workbookDataDigest &&
          current.workbookDataDigest &&
          original.workbookDataDigest !== current.workbookDataDigest
        )
          workbookDataChanged.push(element.name)
      }
    } catch {
      // Unsupported packages remain unchecked; cache evidence still reports.
    }
  }
  for (const element of source) {
    if (element.type === 'picture' && !mediaChecked.includes(element.name))
      mediaUnchecked.push(element.name)
    if (element.type !== 'shape' && element.type !== 'table') {
      unchecked.push(element.name)
      continue
    }
    if (!hostBase64 || !readbackConsistent) {
      unchecked.push(element.name)
      continue
    }
    const actual = exportedByName.get(element.name)
    if (
      !actual ||
      actual.type !== element.type ||
      JSON.stringify(actual.text) !== JSON.stringify(element.text)
    )
      changed.push(element.name)
  }
  const content = {
    status: (changed.length ||
    cacheChanged.length ||
    backgroundChanged ||
    notesChanged ||
    sourceLinkChanged.length ||
    chartTypeChanged.length ||
    chartStyleChanged.length ||
    chartSourceChanged.length ||
    chartFormulaChanged.length ||
    chartSourceVerificationRegressed.length ||
    chartSourceUnreadable.length ||
    workbookDataChanged.length ||
    workbookUnverifiedChanged.length ||
    mediaChanged.length ||
    altTextChanged.length ||
    cropChanged.length ||
    appearanceChanged.length ||
    textStyleChanged.length ||
    tableStyleChanged.length ||
    tableStructureChanged.length
      ? 'warning'
      : unchecked.length || backgroundUnchecked || notesUnchecked || sourceLinksUnchecked
        ? 'incomplete'
        : 'passed') as 'passed' | 'warning' | 'incomplete',
    changed,
    backgroundChanged,
    backgroundUnchecked,
    notesChanged,
    notesUnchecked,
    sourceLinkChanged,
    sourceLinksUnchecked,
    cacheChanged,
    chartTypeChanged,
    chartStyleChanged,
    chartSourceChanged,
    chartFormulaChanged,
    chartSourceVerificationRegressed,
    chartSourceUnreadable,
    workbookBytesChanged,
    workbookDataChanged,
    mediaChanged,
    mediaChecked,
    mediaUnchecked,
    altTextChanged,
    cropChanged,
    appearanceChanged,
    textStyleChanged,
    tableStyleChanged,
    tableStructureChanged,
    unchecked,
  }
  const structureStatus =
    host.shapesTruncated || !readbackConsistent
      ? 'incomplete'
      : issues.length
        ? 'warning'
        : 'passed'
  return {
    status:
      structureStatus === 'incomplete' || content.status === 'incomplete'
        ? 'incomplete'
        : structureStatus === 'warning' || content.status === 'warning'
          ? 'warning'
          : 'passed',
    structureStatus,
    sourceCount: source.length,
    hostCount: host.shapes.length,
    issues,
    readbackConsistent,
    content,
  }
}

// Package observers reuse these checks after their own bounded XML read.
export {
  tableStructure as inspectNativeTableStructure,
  tableCellStyles as inspectNativeTableCellStyles,
}
