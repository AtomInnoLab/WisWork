import { XMLParser } from 'fast-xml-parser'
import type { PowerPointPageInspection } from './browser-powerpoint-adapter.js'
import { loadBoundedZip, MAX_PPTX_XML_BYTES } from './powerpoint-package.js'

type Issue =
  | { name: string; kind: 'missing' | 'extra' | 'duplicate' }
  | { name: string; kind: 'type_changed'; sourceType: string; hostType: string }
  | { name: string; kind: 'geometry_changed' }
type Xml = Record<string, any>
type SourceObject = {
  name: string
  type: string
  box: [number, number, number, number]
  text: string[]
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
      result.push({ name, type, box: box as SourceObject['box'], text: textRuns(element) })
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
  content: { status: 'passed' | 'warning' | 'incomplete'; changed: string[]; unchecked: string[] }
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
  const changed: string[] = [],
    unchecked: string[] = []
  for (const element of source) {
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
    status: (changed.length ? 'warning' : unchecked.length ? 'incomplete' : 'passed') as
      'passed' | 'warning' | 'incomplete',
    changed,
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
