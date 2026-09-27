import { XMLParser, XMLValidator } from 'fast-xml-parser'
import { loadBoundedZip, MAX_PPTX_XML_BYTES } from './powerpoint-package.js'

type Node = Record<string, unknown>
const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  parseTagValue: false,
  trimValues: false,
})
const NOTES_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'

function tags(tree: Node[], tag: string): Node[] {
  return tree.flatMap((node) =>
    Object.entries(node).flatMap(([key, value]) =>
      Array.isArray(value) ? [...(key === tag ? [node] : []), ...tags(value as Node[], tag)] : [],
    ),
  )
}
function attribute(node: Node, name: string): string | undefined {
  const value = (node[':@'] as Node | undefined)?.[`@_${name}`]
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
function paragraph(tree: Node[]): string {
  return tree
    .map((node) =>
      Object.entries(node)
        .map(([tag, value]) =>
          tag === 'a:t' && Array.isArray(value)
            ? contents(value as Node[])
            : tag === 'a:br'
              ? '\n'
              : Array.isArray(value)
                ? paragraph(value as Node[])
                : '',
        )
        .join(''),
    )
    .join('')
}
function parseXml(source: string): Node[] {
  if (
    new TextEncoder().encode(source).byteLength > MAX_PPTX_XML_BYTES ||
    /<!\s*(?:DOCTYPE|ENTITY)\b/i.test(source) ||
    XMLValidator.validate(source) !== true
  )
    throw new Error('office_api_unsupported')
  return parser.parse(source) as Node[]
}

/** Read the body placeholder only; footer, page number and slide image are excluded. */
export async function inspectPowerPointPageNotes(
  base64: string,
  signal?: AbortSignal,
): Promise<{ status: 'read' | 'not_present'; text: string }> {
  const zip = await loadBoundedZip(base64, signal, true, 8 * 1024 * 1024)
  const slides = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  if (slides.length !== 1) throw new Error('office_api_unsupported')
  const relPath = slides[0]!.replace('/slides/', '/slides/_rels/') + '.rels'
  const relFile = zip.file(relPath)
  if (!relFile) return { status: 'not_present', text: '' }
  const rels = tags(parseXml(await relFile.async('string')), 'Relationship')
  const ids = rels.map((rel) => attribute(rel, 'Id'))
  if (ids.some((id) => !id) || new Set(ids).size !== ids.length)
    throw new Error('office_api_unsupported')
  const matches = rels.filter((rel) => attribute(rel, 'Type') === NOTES_REL)
  if (!matches.length) return { status: 'not_present', text: '' }
  if (matches.length !== 1 || attribute(matches[0]!, 'TargetMode') !== undefined)
    throw new Error('office_api_unsupported')
  const target = attribute(matches[0]!, 'Target')
  const match =
    target && /^(?:\.\.\/notesSlides\/|\/ppt\/notesSlides\/)(notesSlide\d+\.xml)$/.exec(target)
  if (!match) throw new Error('office_api_unsupported')
  const noteFile = zip.file(`ppt/notesSlides/${match[1]}`)
  if (!noteFile) throw new Error('office_api_unsupported')
  const notes = parseXml(await noteFile.async('string'))
  const bodies = tags(notes, 'p:sp').filter((shape) =>
    tags(shape['p:sp'] as Node[], 'p:ph').some(
      (placeholder) => attribute(placeholder, 'type') === 'body',
    ),
  )
  if (bodies.length !== 1) throw new Error('office_api_unsupported')
  const textBodies = tags(bodies[0]!['p:sp'] as Node[], 'p:txBody')
  if (textBodies.length !== 1) throw new Error('office_api_unsupported')
  const paragraphs = tags(textBodies[0]!['p:txBody'] as Node[], 'a:p').map((item) =>
    paragraph(item['a:p'] as Node[]),
  )
  while (paragraphs.length && paragraphs[paragraphs.length - 1] === '') paragraphs.pop()
  const text = paragraphs.join('\n')
  if (text.length > 12_000) throw new Error('office_api_unsupported')
  if (signal?.aborted) throw new Error('cancelled')
  return { status: 'read', text }
}
