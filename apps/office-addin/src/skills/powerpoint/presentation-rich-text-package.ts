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
function contents(tree: Node[]): string {
  return tree
    .map((node) =>
      typeof node['#text'] === 'string'
        ? node['#text']
        : Object.entries(node)
            .filter(([, value]) => Array.isArray(value))
            .map(([, value]) => contents(value as Node[]))
            .join(''),
    )
    .join('')
}
function directFont(run: Node[]): Record<string, string | number | boolean> {
  const properties =
    tags(run, 'a:rPr')[0] ?? tags(run, 'a:defRPr')[0] ?? tags(run, 'a:endParaRPr')[0]
  if (!properties) return {}
  const propertyTree = (properties['a:rPr'] ??
    properties['a:defRPr'] ??
    properties['a:endParaRPr']) as Node[]
  const size = attr(properties, 'sz')
  const fill = tags(propertyTree, 'a:solidFill')[0]
  const fillTree = (fill?.['a:solidFill'] as Node[] | undefined) ?? []
  const color = attr(tags(fillTree, 'a:srgbClr')[0], 'val')
  const themeColor = attr(tags(fillTree, 'a:schemeClr')[0], 'val')
  const typeface = attr(tags(propertyTree, 'a:latin')[0], 'typeface')
  return {
    ...(size && /^\d+$/.test(size) ? { sizePt: Number(size) / 100 } : {}),
    ...(color && /^[0-9a-fA-F]{6}$/.test(color) ? { color: `#${color.toUpperCase()}` } : {}),
    ...(themeColor ? { themeColor: themeColor.slice(0, 64) } : {}),
    ...(typeface ? { typeface: typeface.slice(0, 256) } : {}),
    ...(attr(properties, 'b') !== undefined
      ? { bold: ['1', 'true'].includes(attr(properties, 'b')!) }
      : {}),
    ...(attr(properties, 'i') !== undefined
      ? { italic: ['1', 'true'].includes(attr(properties, 'i')!) }
      : {}),
    ...(attr(properties, 'u') ? { underline: attr(properties, 'u')! } : {}),
  }
}

export interface RichTextRun {
  text: string
  directFont: Record<string, string | number | boolean>
}
export interface RichTextShape {
  packageShapeId: string
  name: string
  paragraphs: Array<{
    alignment?: string
    paragraphDefaultFont?: Record<string, string | number | boolean>
    endParagraphFont?: Record<string, string | number | boolean>
    runs: RichTextRun[]
  }>
}

/** Direct formatting only; theme, layout and master inheritance remain unresolved. */
export async function inspectPowerPointRichText(
  base64: string,
  signal?: AbortSignal,
): Promise<{ shapes: RichTextShape[]; inheritanceResolved: false }> {
  const zip = await loadBoundedZip(base64, signal, true, 8 * 1024 * 1024)
  const slides = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  if (slides.length !== 1) throw new Error('office_api_unsupported')
  const slide = xml(await zip.file(slides[0]!)!.async('string'))
  const shapes: RichTextShape[] = []
  let runCount = 0
  let textLength = 0
  for (const shape of tags(slide, 'p:sp')) {
    const body = tags(shape['p:sp'] as Node[], 'p:txBody')[0]
    if (!body) continue
    const identity = tags(shape['p:sp'] as Node[], 'p:cNvPr')[0]
    const packageShapeId = attr(identity, 'id')
    if (!packageShapeId || !/^\d+$/.test(packageShapeId)) throw new Error('office_api_unsupported')
    const paragraphs = tags(body['p:txBody'] as Node[], 'a:p').map((paragraph) => {
      const children = paragraph['a:p'] as Node[]
      const paragraphProperties = tags(children, 'a:pPr')[0]
      const alignment = attr(paragraphProperties, 'algn')
      const paragraphDefaultFont = paragraphProperties
        ? directFont(paragraphProperties['a:pPr'] as Node[])
        : {}
      const endParagraphFont = directFont(tags(children, 'a:endParaRPr'))
      const runs: RichTextRun[] = []
      for (const child of children) {
        const entry = child['a:r'] ?? child['a:fld']
        if (Array.isArray(entry)) {
          const text = tags(entry as Node[], 'a:t')
            .map((node) => contents(node['a:t'] as Node[]))
            .join('')
          runs.push({ text, directFont: directFont(entry as Node[]) })
        } else if (Array.isArray(child['a:br'])) {
          runs.push({ text: '\n', directFont: {} })
        } else if (Array.isArray(child['a:tab'])) {
          runs.push({ text: '\t', directFont: {} })
        }
      }
      runCount += runs.length
      textLength += runs.reduce((total, run) => total + run.text.length, 0)
      return {
        ...(alignment ? { alignment } : {}),
        ...(Object.keys(paragraphDefaultFont).length ? { paragraphDefaultFont } : {}),
        ...(Object.keys(endParagraphFont).length ? { endParagraphFont } : {}),
        runs,
      }
    })
    shapes.push({ packageShapeId, name: (attr(identity, 'name') ?? '').slice(0, 256), paragraphs })
    if (shapes.length > 100 || runCount > 2000 || textLength > 120_000)
      throw new Error('office_api_unsupported')
  }
  if (signal?.aborted) throw new Error('cancelled')
  return { shapes, inheritanceResolved: false }
}
