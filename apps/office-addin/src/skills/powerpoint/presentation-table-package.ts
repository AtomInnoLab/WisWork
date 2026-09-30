import { XMLParser, XMLValidator } from 'fast-xml-parser'
import { loadBoundedZip, MAX_PPTX_PAGE_SHAPES, MAX_PPTX_XML_BYTES } from './powerpoint-package.js'

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

function parse(source: string): Node[] {
  if (
    new TextEncoder().encode(source).byteLength > MAX_PPTX_XML_BYTES ||
    /<!\s*(?:DOCTYPE|ENTITY)\b/i.test(source) ||
    XMLValidator.validate(source) !== true
  )
    throw Error('office_api_unsupported')
  return parser.parse(source) as Node[]
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

async function digest(value: unknown): Promise<string> {
  const bytes = new TextEncoder().encode(JSON.stringify(stable(value)))
  const hash = new Uint8Array(await crypto.subtle.digest('SHA-256', bytes))
  return Array.from(hash, (byte) => byte.toString(16).padStart(2, '0')).join('')
}

/** Exact table-frame XML and shared table-style part for one bounded page package. */
export async function inspectPowerPointTableFingerprints(
  base64: string,
  shapeIds: string[],
  signal?: AbortSignal,
): Promise<Record<string, string>> {
  if (signal?.aborted) throw Error('cancelled')
  if (
    shapeIds.length > MAX_PPTX_PAGE_SHAPES ||
    new Set(shapeIds).size !== shapeIds.length ||
    shapeIds.some((id) => !/^\d{1,10}$/.test(id))
  )
    throw Error('invalid_tool_input')
  const zip = await loadBoundedZip(base64, signal, true, 8 * 1024 * 1024)
  const paths = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  if (paths.length !== 1) throw Error('office_api_unsupported')
  const slide = parse(await zip.file(paths[0]!)!.async('string'))
  const tableStyles = zip.file('ppt/tableStyles.xml')
    ? parse(await zip.file('ppt/tableStyles.xml')!.async('string'))
    : null
  const frames = nodes(slide, 'p:graphicFrame').filter(
    (frame) => nodes(frame['p:graphicFrame'] as Node[], 'a:tbl').length,
  )
  const fingerprints: Record<string, string> = Object.create(null)
  for (const frame of frames) {
    const id = (
      nodes(frame['p:graphicFrame'] as Node[], 'p:cNvPr')[0]?.[':@'] as Node | undefined
    )?.['@_id']
    if (typeof id !== 'string' || !/^\d{1,10}$/.test(id) || Object.hasOwn(fingerprints, id))
      throw Error('office_api_unsupported')
    if (shapeIds.includes(id)) fingerprints[id] = await digest([frame, tableStyles])
  }
  if (Object.keys(fingerprints).length !== shapeIds.length) throw Error('office_api_unsupported')
  if (signal?.aborted) throw Error('cancelled')
  return fingerprints
}
