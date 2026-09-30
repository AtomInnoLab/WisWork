import { XMLParser, XMLValidator } from 'fast-xml-parser'
import { loadBoundedZip, MAX_PPTX_XML_BYTES } from './powerpoint-package.js'

type Node = Record<string, unknown>
const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  parseTagValue: false,
  trimValues: false,
})

function nodes(tree: Node[], tag: string): Node[] {
  return tree.flatMap((node) =>
    Object.entries(node).flatMap(([key, value]) =>
      Array.isArray(value) ? [...(key === tag ? [node] : []), ...nodes(value as Node[], tag)] : [],
    ),
  )
}

function stable(value: unknown): unknown {
  if (Array.isArray(value))
    return value
      .filter(
        (part) =>
          !(
            part &&
            typeof part === 'object' &&
            Object.keys(part).length === 1 &&
            typeof (part as Node)['#text'] === 'string' &&
            /^\s*$/.test((part as Node)['#text'] as string)
          ),
      )
      .map(stable)
  if (!value || typeof value !== 'object') return value
  return Object.fromEntries(
    Object.entries(value as Node)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([key, child]) => [key, stable(child)]),
  )
}

async function hash(value: Uint8Array): Promise<string> {
  const bytes = new Uint8Array(await crypto.subtle.digest('SHA-256', Uint8Array.from(value)))
  return Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0')).join('')
}

/** Exact frame plus the internal parts reached through this chart's relationships. */
export async function inspectPowerPointChartFingerprints(
  base64: string,
  shapeIds: string[],
  signal?: AbortSignal,
): Promise<Record<string, string>> {
  if (signal?.aborted) throw Error('cancelled')
  if (
    shapeIds.length > 100 ||
    new Set(shapeIds).size !== shapeIds.length ||
    shapeIds.some((id) => !/^\d{1,10}$/.test(id))
  )
    throw Error('invalid_tool_input')
  const zip = await loadBoundedZip(base64, signal, true, 8 * 1024 * 1024)
  const slides = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  if (slides.length !== 1) throw Error('office_api_unsupported')
  const source = await zip.file(slides[0]!)!.async('string')
  if (
    new TextEncoder().encode(source).byteLength > MAX_PPTX_XML_BYTES ||
    /<!\s*(?:DOCTYPE|ENTITY)\b/i.test(source) ||
    XMLValidator.validate(source) !== true
  )
    throw Error('office_api_unsupported')
  const slide = parser.parse(source) as Node[]
  const parseXml = (value: string) => {
    if (
      new TextEncoder().encode(value).byteLength > MAX_PPTX_XML_BYTES ||
      /<!\s*(?:DOCTYPE|ENTITY)\b/i.test(value) ||
      XMLValidator.validate(value) !== true
    )
      throw Error('office_api_unsupported')
    return parser.parse(value) as Node[]
  }
  const relationships = async (part: string) => {
    const slash = part.lastIndexOf('/')
    const path = `${part.slice(0, slash)}/_rels/${part.slice(slash + 1)}.rels`
    const file = zip.file(path)
    if (!file) return []
    return nodes(parseXml(await file.async('string')), 'Relationship').map((entry) => {
      const attributes = entry[':@'] as Node | undefined
      const id = attributes?.['@_Id'],
        target = attributes?.['@_Target']
      if (
        typeof id !== 'string' ||
        typeof target !== 'string' ||
        attributes?.['@_TargetMode'] === 'External'
      )
        throw Error('office_api_unsupported')
      const result = target.startsWith('/') ? [] : part.slice(0, slash).split('/')
      for (const segment of target.split('/')) {
        if (segment === '..') result.pop()
        else if (segment !== '.' && segment !== '') result.push(segment)
      }
      const resolved = result.join('/')
      if (!resolved.startsWith('ppt/') || !zip.file(resolved)) throw Error('office_api_unsupported')
      return { id, path: resolved }
    })
  }
  const slideRelationships = await relationships(slides[0]!)
  const partsFor = async (chartPath: string) => {
    const visited = new Set<string>(),
      pending = [chartPath]
    while (pending.length) {
      if (signal?.aborted) throw Error('cancelled')
      const path = pending.pop()!
      if (visited.has(path)) continue
      visited.add(path)
      if (visited.size > 100) throw Error('office_api_unsupported')
      for (const related of await relationships(path)) pending.push(related.path)
    }
    return Promise.all(
      [...visited]
        .sort()
        .map(
          async (path) => [path, await hash(await zip.file(path)!.async('uint8array'))] as const,
        ),
    )
  }
  const frames = nodes(slide, 'p:graphicFrame').filter(
    (frame) => nodes(frame['p:graphicFrame'] as Node[], 'c:chart').length,
  )
  const fingerprints: Record<string, string> = Object.create(null)
  for (const frame of frames) {
    const id = (
      nodes(frame['p:graphicFrame'] as Node[], 'p:cNvPr')[0]?.[':@'] as Node | undefined
    )?.['@_id']
    if (typeof id !== 'string' || !/^\d{1,10}$/.test(id) || Object.hasOwn(fingerprints, id))
      throw Error('office_api_unsupported')
    if (shapeIds.includes(id)) {
      const references = nodes(frame['p:graphicFrame'] as Node[], 'c:chart')
      const relationId = (references[0]?.[':@'] as Node | undefined)?.['@_r:id']
      const related = slideRelationships.filter((item) => item.id === relationId)
      if (
        references.length !== 1 ||
        related.length !== 1 ||
        !/^ppt\/charts\/chart\d+\.xml$/.test(related[0]!.path)
      )
        throw Error('office_api_unsupported')
      fingerprints[id] = await hash(
        new TextEncoder().encode(
          JSON.stringify([stable(frame), related[0]!.path, await partsFor(related[0]!.path)]),
        ),
      )
    }
  }
  if (Object.keys(fingerprints).length !== shapeIds.length) throw Error('office_api_unsupported')
  if (signal?.aborted) throw Error('cancelled')
  return fingerprints
}
