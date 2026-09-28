import { XMLParser, XMLValidator } from 'fast-xml-parser'
import type JSZip from 'jszip'
import { loadBoundedZip, MAX_PPTX_XML_BYTES } from './powerpoint-package.js'

type Node = Record<string, unknown>
const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  parseTagValue: false,
  trimValues: false,
})
const HYPERLINK_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'

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
function attr(node: Node, name: string): string | undefined {
  const value = (node[':@'] as Node | undefined)?.[`@_${name}`]
  return typeof value === 'string' ? value : undefined
}
function text(tree: Node[]): string {
  return tags(tree, 'a:t')
    .map((node) => (node['a:t'] as Node[]).map((part) => part['#text'] ?? '').join(''))
    .join('')
}

export interface PresentationSourceLink {
  packageShapeId: string
  target: string
  label: string
  location: 'text_run' | 'shape_action'
  sourceVerified: false
}

/** Extract only explicit external links in a single exported slide. Targets are never fetched. */
export async function inspectPowerPointSourceLinks(
  base64: string,
  signal?: AbortSignal,
): Promise<{ status: 'read' | 'not_present'; links: PresentationSourceLink[] }> {
  const zip = await loadBoundedZip(base64, signal, true, 8 * 1024 * 1024)
  const slides = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  if (slides.length !== 1) throw new Error('office_api_unsupported')
  return inspectPowerPointSourceLinksFromZip(zip, slides[0]!, signal)
}

/** Reuses the same bounded parser for one slide selected from a multi-slide source package. */
export async function inspectPowerPointSourceLinksFromZip(
  zip: JSZip,
  slidePath: string,
  signal?: AbortSignal,
): Promise<{ status: 'read' | 'not_present'; links: PresentationSourceLink[] }> {
  if (!/^ppt\/slides\/slide[1-9]\d*\.xml$/.test(slidePath) || !zip.file(slidePath))
    throw new Error('office_api_unsupported')
  const slide = xml(await zip.file(slidePath)!.async('string'))
  const relPath = slidePath.replace('/slides/', '/slides/_rels/') + '.rels'
  const relFile = zip.file(relPath)
  const rels = relFile ? tags(xml(await relFile.async('string')), 'Relationship') : []
  const targets = new Map<string, string>()
  const ids = new Set<string>()
  for (const rel of rels) {
    const id = attr(rel, 'Id')
    if (!id || ids.has(id)) throw new Error('office_api_unsupported')
    ids.add(id)
    if (attr(rel, 'Type') !== HYPERLINK_REL) continue
    const target = attr(rel, 'Target')
    if (
      attr(rel, 'TargetMode') !== 'External' ||
      !target ||
      target.length > 4096 ||
      !/^https?:\/\//i.test(target) ||
      Array.from(target).some((character) => {
        const code = character.charCodeAt(0)
        return code < 32 || code === 127
      })
    )
      continue
    targets.set(id, target)
  }
  const links: PresentationSourceLink[] = []
  for (const shape of tags(slide, 'p:sp')) {
    const shapeId = attr(tags(shape['p:sp'] as Node[], 'p:cNvPr')[0] ?? {}, 'id')
    if (!shapeId || !/^\d+$/.test(shapeId)) throw new Error('office_api_unsupported')
    const shapeText = text(shape['p:sp'] as Node[]).slice(0, 500)
    for (const run of tags(shape['p:sp'] as Node[], 'a:r')) {
      const label = text(run['a:r'] as Node[]).slice(0, 500)
      for (const click of tags(run['a:r'] as Node[], 'a:hlinkClick')) {
        const target = targets.get(attr(click, 'r:id') ?? '')
        if (target)
          links.push({
            packageShapeId: shapeId,
            target,
            label,
            location: 'text_run',
            sourceVerified: false,
          })
      }
    }
    for (const click of tags(shape['p:sp'] as Node[], 'p:cNvPr').flatMap((node) =>
      tags(node['p:cNvPr'] as Node[], 'a:hlinkClick'),
    )) {
      const target = targets.get(attr(click, 'r:id') ?? '')
      if (
        target &&
        !links.some((link) => link.packageShapeId === shapeId && link.target === target)
      )
        links.push({
          packageShapeId: shapeId,
          target,
          label: shapeText,
          location: 'shape_action',
          sourceVerified: false,
        })
    }
    if (links.length > 64) throw new Error('office_api_unsupported')
  }
  if (signal?.aborted) throw new Error('cancelled')
  return { status: links.length ? 'read' : 'not_present', links }
}
