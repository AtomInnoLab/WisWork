import { XMLParser } from 'fast-xml-parser'
import type { PowerPointPageInspection } from './browser-powerpoint-adapter.js'
import { loadBoundedZip, MAX_PPTX_XML_BYTES } from './powerpoint-package.js'

type Issue =
  | { name: string; kind: 'missing' | 'extra' | 'duplicate' }
  | { name: string; kind: 'type_changed'; sourceType: string; hostType: string }
  | { name: string; kind: 'geometry_changed' }
type Xml = Record<string, any>
type SourceObject = { name: string; type: string; box: [number, number, number, number] }
const POINTS_PER_EMU = 72 / 914400
const types: Record<string, string[]> = {
  shape: ['TextBox', 'GeometricShape'],
  picture: ['Image'],
  table: ['Table'],
  chart: ['Chart'],
}
const parser = new XMLParser({ ignoreAttributes: false, parseAttributeValue: false })
const many = (value: unknown): Xml[] =>
  value === undefined ? [] : Array.isArray(value) ? value : [value as Xml]

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
        !type ||
        box.some((value) => !Number.isFinite(value))
      )
        throw new Error('presentation_qa_structure_unavailable')
      result.push({ name, type, box: box as SourceObject['box'] })
    }
  }
  return result
}

/** Compare source PPTX objects with the exact imported Office page; no host mutation. */
export async function comparePresentationPageStructure(
  sourceBase64: string,
  sourceIndex: number,
  host: PowerPointPageInspection,
): Promise<{
  status: 'passed' | 'warning' | 'incomplete'
  sourceCount: number
  hostCount: number
  issues: Issue[]
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
  let source: SourceObject[]
  try {
    const zip = await loadBoundedZip(sourceBase64, undefined, true, 10 * 1024 * 1024)
    const file = zip.file(`ppt/slides/slide${sourceIndex + 1}.xml`)
    if (!file) throw new Error('missing slide')
    const xml = await file.async('string')
    if (
      new TextEncoder().encode(xml).byteLength > MAX_PPTX_XML_BYTES ||
      /<!\s*(?:DOCTYPE|ENTITY)\b/i.test(xml)
    )
      throw new Error('invalid slide XML')
    source = sourceObjects(xml)
  } catch {
    throw new Error('presentation_qa_structure_unavailable')
  }
  if (!source.length || source.length > 100)
    throw new Error('presentation_qa_structure_unavailable')
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
  return {
    status: host.shapesTruncated ? 'incomplete' : issues.length ? 'warning' : 'passed',
    sourceCount: source.length,
    hostCount: host.shapes.length,
    issues,
  }
}
