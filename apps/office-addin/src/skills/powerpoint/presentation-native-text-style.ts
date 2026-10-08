import { XMLParser, XMLValidator } from 'fast-xml-parser'
import { loadBoundedZip, MAX_PPTX_XML_BYTES } from './powerpoint-package.js'
import type { NativeModifyOperation } from './presentation-existing-batch.js'

type Node = Record<string, unknown>
type Style = Extract<NativeModifyOperation, { op: 'set_shape_text_style' }>
const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  parseTagValue: false,
  trimValues: false,
})
const tags = (nodes: Node[], tag: string): Node[] =>
  nodes.flatMap((node) =>
    Object.entries(node).flatMap(([key, value]) =>
      Array.isArray(value) ? [...(key === tag ? [node] : []), ...tags(value as Node[], tag)] : [],
    ),
  )

/** Prove that the host changed only requested font fields, including XML not exposed by Office.js. */
export async function verifyNativeTextStylePackage(
  before: string,
  after: string,
  shapeId: string,
  style: Style,
): Promise<void> {
  const canonical = async (base64: string) => {
    const zip = await loadBoundedZip(base64, undefined, true, 8 * 1024 * 1024)
    const slides = Object.keys(zip.files).filter((path) =>
      /^ppt\/slides\/slide\d+\.xml$/.test(path),
    )
    if (slides.length !== 1) throw Error('office_api_unsupported')
    const source = await zip.file(slides[0]!)!.async('string')
    if (
      new TextEncoder().encode(source).byteLength > MAX_PPTX_XML_BYTES ||
      /<!\s*(?:DOCTYPE|ENTITY)\b/i.test(source) ||
      XMLValidator.validate(source) !== true
    )
      throw Error('office_api_unsupported')
    const tree = parser.parse(source) as Node[]
    const shapes = tags(tree, 'p:sp').filter((shape) =>
      tags(shape['p:sp'] as Node[], 'p:cNvPr').some(
        (id) => (id[':@'] as Node | undefined)?.['@_id'] === shapeId,
      ),
    )
    if (shapes.length !== 1) throw Error('office_api_unsupported')
    const normalize = (nodes: Node[]): Node[] =>
      nodes.flatMap((node) => {
        const copy: Node = {}
        for (const [key, value] of Object.entries(node))
          copy[key] = Array.isArray(value) ? normalize(value as Node[]) : value
        const tag = ['a:rPr', 'a:defRPr', 'a:endParaRPr'].find((name) => Object.hasOwn(copy, name))
        if (tag) {
          const attrs = { ...(copy[':@'] as Node | undefined) }
          if (style.fontSize !== undefined) delete attrs['@_sz']
          if (style.bold !== undefined) delete attrs['@_b']
          if (style.italic !== undefined) delete attrs['@_i']
          copy[tag] = (copy[tag] as Node[]).filter(
            (child) =>
              !(
                style.color !== undefined &&
                ['a:solidFill', 'a:gradFill', 'a:noFill'].some((name) => Object.hasOwn(child, name))
              ) && !(style.fontFamily !== undefined && Object.hasOwn(child, 'a:latin')),
          )
          if (Object.keys(attrs).length) copy[':@'] = attrs
          else delete copy[':@']
          if (!(copy[tag] as Node[]).length && !copy[':@']) return []
        }
        return [copy]
      })
    const stable = (value: unknown): unknown =>
      Array.isArray(value)
        ? value.map(stable)
        : value && typeof value === 'object'
          ? Object.fromEntries(
              Object.entries(value)
                .sort(([a], [b]) => a.localeCompare(b))
                .map(([key, entry]) => [key, stable(entry)]),
            )
          : value
    const resources = await Promise.all(
      Object.keys(zip.files)
        .filter((path) => path.startsWith('ppt/') && path !== slides[0] && !zip.files[path]!.dir)
        .sort()
        .map(async (path) => {
          const bytes = await zip.file(path)!.async('uint8array')
          const digest = await crypto.subtle.digest('SHA-256', new Uint8Array(bytes).buffer)
          return [
            path,
            Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join(
              '',
            ),
          ]
        }),
    )
    const normalizedShape = normalize(shapes)[0]!
    for (const key of Object.keys(shapes[0]!)) delete shapes[0]![key]
    Object.assign(shapes[0]!, normalizedShape)
    return JSON.stringify([stable(tree), resources])
  }
  if ((await canonical(before)) !== (await canonical(after))) throw Error('office_verify_failed')
}
