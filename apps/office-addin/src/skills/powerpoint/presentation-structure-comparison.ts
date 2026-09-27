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
  altText?: string
  crop?: [number, number, number, number]
  appearance?: [string, string, string]
  rotation: number
}
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
const many = (value: unknown): Xml[] =>
  value === undefined ? [] : Array.isArray(value) ? value : [value as Xml]

function textRuns(value: unknown): string[] {
  if (Array.isArray(value)) return value.flatMap(textRuns)
  if (!value || typeof value !== 'object') return []
  return Object.entries(value).flatMap(([tag, child]) =>
    tag === 'a:t' ? [String(child)] : tag.startsWith('@_') ? [] : textRuns(child),
  )
}

function textStyles(element: Xml): Array<[string, string, string, string]> {
  return many(element['p:txBody']?.['a:p']).flatMap((paragraph) =>
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
        ...(tag === 'p:sp' ? { textStyles: textStyles(element) } : {}),
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
): Promise<{ objects: SourceObject[]; backgroundColor?: string }> {
  try {
    const zip = await loadBoundedZip(base64, undefined, true, maxBytes)
    const slidePath = `ppt/slides/slide${index + 1}.xml`
    const root = await readXml(zip, slidePath)
    if (!root) throw new Error('invalid slide XML')
    const objects = sourceObjects(root)
    if (!objects.length || objects.length > 100) throw new Error('invalid slide objects')
    const backgroundColor = await resolvedBackground(zip, slidePath, root)
    return { objects, ...(backgroundColor ? { backgroundColor } : {}) }
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
    workbookBytesChanged: string[]
    backgroundChanged: boolean
    backgroundUnchecked: boolean
    mediaChanged: string[]
    mediaChecked: string[]
    mediaUnchecked: string[]
    altTextChanged: string[]
    cropChanged: string[]
    appearanceChanged: string[]
    textStyleChanged: string[]
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
  if (hostBase64 && readbackConsistent)
    for (const element of source) {
      const actual = exportedByName.get(element.name)
      if (actual?.type === element.type && actual.rotation !== element.rotation)
        issues.push({ name: element.name, kind: 'rotation_changed' })
    }
  const changed: string[] = [],
    cacheChanged: string[] = [],
    chartTypeChanged: string[] = [],
    workbookBytesChanged: string[] = [],
    mediaChanged: string[] = [],
    mediaChecked: string[] = [],
    mediaUnchecked: string[] = [],
    altTextChanged: string[] = [],
    cropChanged: string[] = [],
    appearanceChanged: string[] = [],
    textStyleChanged: string[] = [],
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
          maxBytes: 10 * 1024 * 1024,
        }),
        inspectPowerPointComplexPagePackage(hostBase64),
      ])
      for (const element of source.filter((item) => item.type === 'chart')) {
        const hostElement = exportedByName.get(element.name)
        const original = before.charts.find((chart) => chart.shapeId === element.shapeId)
        const current = after.charts.find((chart) => chart.shapeId === hostElement?.shapeId)
        if (!original || !current || original.truncated || current.truncated) continue
        if (JSON.stringify(original.plotTypes) !== JSON.stringify(current.plotTypes))
          chartTypeChanged.push(element.name)
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
        if (
          original?.sourceKind === 'embedded_xlsx' &&
          current?.sourceKind === 'embedded_xlsx' &&
          original.workbookContentDigest &&
          current.workbookContentDigest &&
          original.workbookContentDigest !== current.workbookContentDigest
        )
          workbookBytesChanged.push(element.name)
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
    chartTypeChanged.length ||
    workbookBytesChanged.length ||
    mediaChanged.length ||
    altTextChanged.length ||
    cropChanged.length ||
    appearanceChanged.length ||
    textStyleChanged.length
      ? 'warning'
      : unchecked.length || backgroundUnchecked
        ? 'incomplete'
        : 'passed') as 'passed' | 'warning' | 'incomplete',
    changed,
    backgroundChanged,
    backgroundUnchecked,
    cacheChanged,
    chartTypeChanged,
    workbookBytesChanged,
    mediaChanged,
    mediaChecked,
    mediaUnchecked,
    altTextChanged,
    cropChanged,
    appearanceChanged,
    textStyleChanged,
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
