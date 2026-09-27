import { XMLParser } from 'fast-xml-parser'
import type { PowerPointPageInspection } from './browser-powerpoint-adapter.js'
import {
  inspectPowerPointPictureMediaBatch,
  loadBoundedZip,
  MAX_PPTX_XML_BYTES,
} from './powerpoint-package.js'
import { inspectPowerPointComplexPagePackage } from './presentation-complex-page-package.js'
import { inspectPowerPointChartSourcePackage } from './presentation-chart-source-package.js'

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

function sourceObjects(xml: string): SourceObject[] {
  const tree = parser.parse(xml)?.['p:sld']?.['p:cSld']?.['p:spTree'] as Xml | undefined
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

async function readObjects(
  base64: string,
  index: number,
  maxBytes: number,
): Promise<SourceObject[]> {
  try {
    const zip = await loadBoundedZip(base64, undefined, true, maxBytes)
    const file = zip.file(`ppt/slides/slide${index + 1}.xml`)
    if (!file) throw new Error('missing slide')
    const xml = await file.async('string')
    if (
      new TextEncoder().encode(xml).byteLength > MAX_PPTX_XML_BYTES ||
      /<!\s*(?:DOCTYPE|ENTITY)\b/i.test(xml)
    )
      throw new Error('invalid slide XML')
    const objects = sourceObjects(xml)
    if (!objects.length || objects.length > 100) throw new Error('invalid slide objects')
    return objects
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
  const source = await readObjects(sourceBase64, sourceIndex, 10 * 1024 * 1024)
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
  const exported = hostBase64 ? await readObjects(hostBase64, 0, 8 * 1024 * 1024) : []
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
    for (const element of source.filter((item) => item.type === 'chart')) {
      const hostElement = exportedByName.get(element.name)
      if (!hostElement || hostElement.type !== 'chart') continue
      try {
        const [before, after] = await Promise.all([
          inspectPowerPointChartSourcePackage(sourceBase64, element.shapeId, undefined, {
            slideIndex: sourceIndex,
            maxBytes: 10 * 1024 * 1024,
            allowAbsoluteChartTarget: true,
            includeWorkbookContentDigest: true,
          }),
          inspectPowerPointChartSourcePackage(hostBase64, hostElement.shapeId, undefined, {
            allowAbsoluteChartTarget: true,
            includeWorkbookContentDigest: true,
          }),
        ])
        if (
          before.sourceKind === 'embedded_xlsx' &&
          after.sourceKind === 'embedded_xlsx' &&
          before.workbookContentDigest &&
          after.workbookContentDigest &&
          before.workbookContentDigest !== after.workbookContentDigest
        )
          workbookBytesChanged.push(element.name)
      } catch {
        // Unsupported workbooks remain unchecked; cache evidence still reports.
      }
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
    chartTypeChanged.length ||
    workbookBytesChanged.length ||
    mediaChanged.length ||
    altTextChanged.length ||
    cropChanged.length ||
    appearanceChanged.length ||
    textStyleChanged.length
      ? 'warning'
      : unchecked.length
        ? 'incomplete'
        : 'passed') as 'passed' | 'warning' | 'incomplete',
    changed,
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
