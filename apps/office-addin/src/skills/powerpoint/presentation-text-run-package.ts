import { XMLParser, XMLValidator } from 'fast-xml-parser'
import { loadBoundedZip, MAX_PPTX_XML_BYTES } from './powerpoint-package.js'

type Node = Record<string, unknown>
const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  parseTagValue: false,
  trimValues: false,
})

function xml(source: string): Node[] {
  if (
    new TextEncoder().encode(source).byteLength > MAX_PPTX_XML_BYTES ||
    /<!\s*(?:DOCTYPE|ENTITY)\b/i.test(source) ||
    XMLValidator.validate(source) !== true
  )
    throw new Error('office_api_unsupported')
  return parser.parse(source) as Node[]
}
function tags(tree: Node[], tag: string): Node[] {
  return tree.flatMap((node) =>
    Object.entries(node).flatMap(([key, value]) =>
      Array.isArray(value) ? [...(key === tag ? [node] : []), ...tags(value as Node[], tag)] : [],
    ),
  )
}
function attr(node: Node | undefined, name: string): string | undefined {
  const value = (node?.[':@'] as Node | undefined)?.[`@_${name}`]
  return typeof value === 'string' ? value : undefined
}
function text(node: Node): string {
  const children = node['a:t'] as Node[] | undefined
  if (
    !children ||
    children.some((part) => typeof part['#text'] !== 'string' || Object.keys(part).length !== 1)
  )
    throw new Error('presentation_existing_target_unsupported')
  return children.map((part) => part['#text'] as string).join('')
}
function stable(value: unknown, preserveText = false): unknown {
  if (Array.isArray(value))
    return value
      .filter(
        (part) =>
          preserveText ||
          !(
            part &&
            typeof part === 'object' &&
            Object.keys(part).length === 1 &&
            typeof (part as Node)['#text'] === 'string' &&
            /^\s*$/.test((part as Node)['#text'] as string)
          ),
      )
      .map((part) => stable(part, preserveText))
  if (!value || typeof value !== 'object') return value
  return Object.fromEntries(
    Object.entries(value as Node)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([key, child]) => [key, stable(child, preserveText || key === 'a:t')]),
  )
}

/** Hashes one exact native shape with only the chosen run's text payload excluded. */
export async function inspectPowerPointTextRunPackage(
  base64: string,
  shapeId: string,
  fullText: string,
  start: number,
  length: number,
  signal?: AbortSignal,
): Promise<{ structureDigest: string }> {
  if (signal?.aborted) throw new Error('cancelled')
  if (
    typeof shapeId !== 'string' ||
    !shapeId ||
    shapeId.length > 256 ||
    typeof fullText !== 'string' ||
    fullText.length > 12000 ||
    !Number.isSafeInteger(start) ||
    !Number.isSafeInteger(length) ||
    start < 0 ||
    length < 1 ||
    length > 128 ||
    start + length > fullText.length ||
    /[\r\n\uD800-\uDFFF]/.test(fullText)
  )
    throw new Error('invalid_tool_input')
  const zip = await loadBoundedZip(base64, signal, true, 8 * 1024 * 1024)
  const slides = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  if (slides.length !== 1) throw new Error('office_api_unsupported')
  const slide = xml(await zip.file(slides[0]!)!.async('string'))
  const shapes = tags(slide, 'p:sp').filter(
    (shape) => attr(tags(shape['p:sp'] as Node[], 'p:cNvPr')[0], 'id') === shapeId,
  )
  if (shapes.length !== 1) throw new Error('presentation_existing_target_unsupported')
  const shape = shapes[0]!
  const body = tags(shape['p:sp'] as Node[], 'p:txBody')
  if (
    body.length !== 1 ||
    tags(body[0]!['p:txBody'] as Node[], 'a:p').length !== 1 ||
    ['a:fld', 'a:br', 'a:tab', 'a:hlinkClick', 'a:hlinkMouseOver'].some(
      (tag) => tags(shape['p:sp'] as Node[], tag).length,
    )
  )
    throw new Error('presentation_existing_target_unsupported')
  const runs = tags(body[0]!['p:txBody'] as Node[], 'a:r')
  if (!runs.length || runs.length > 2000)
    throw new Error('presentation_existing_target_unsupported')
  let offset = 0
  let target: Node | undefined
  for (const run of runs) {
    const textNodes = tags(run['a:r'] as Node[], 'a:t')
    if (textNodes.length !== 1) throw new Error('presentation_existing_target_unsupported')
    const value = text(textNodes[0]!)
    if (start >= offset && start + length <= offset + value.length) {
      if (target) throw new Error('presentation_existing_target_unsupported')
      target = textNodes[0]
    }
    offset += value.length
  }
  if (
    !target ||
    offset !== fullText.length ||
    runs.map((run) => text(tags(run['a:r'] as Node[], 'a:t')[0]!)).join('') !== fullText
  )
    throw new Error('presentation_existing_target_unsupported')
  target['a:t'] = [{ '#text': '__WISWORK_TARGET_RUN_TEXT__' }]
  const bytes = new TextEncoder().encode(JSON.stringify(stable(shape)))
  const hash = new Uint8Array(await crypto.subtle.digest('SHA-256', bytes))
  if (signal?.aborted) throw new Error('cancelled')
  return {
    structureDigest: Array.from(hash, (byte) => byte.toString(16).padStart(2, '0')).join(''),
  }
}
