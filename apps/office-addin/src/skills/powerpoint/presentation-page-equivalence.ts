import { XMLParser, XMLValidator } from 'fast-xml-parser'
import { loadBoundedZip, presentationPackageDigest } from './powerpoint-package.js'

const ordered = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
})
const parser = new XMLParser({ ignoreAttributes: false, parseAttributeValue: false })
const sha256 = async (bytes: Uint8Array): Promise<string> =>
  Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', new Uint8Array(bytes))), (byte) =>
    byte.toString(16).padStart(2, '0'),
  ).join('')
const canonical = (value: unknown, preserveText = false): unknown => {
  if (Array.isArray(value))
    return value
      .filter(
        (item) =>
          preserveText ||
          !(
            item &&
            typeof item === 'object' &&
            Object.keys(item).length === 1 &&
            typeof item['#text'] === 'string' &&
            /^\s*$/.test(item['#text'])
          ),
      )
      .map((item) => canonical(item, preserveText))
  if (!value || typeof value !== 'object') return value
  const record = value as Record<string, unknown>
  return Object.fromEntries(
    Object.entries(record)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([key, child]) => [
        key,
        canonical(
          child,
          preserveText ||
            ['a:t', 'a:instrText'].includes(key) ||
            (record[':@'] as Record<string, unknown> | undefined)?.['@_xml:space'] === 'preserve',
        ),
      ]),
  )
}
const xml = (value: string): boolean =>
  new TextEncoder().encode(value).byteLength <= 512 * 1024 &&
  !/<!\s*(?:DOCTYPE|ENTITY)\b/i.test(value) &&
  XMLValidator.validate(value) === true

/** A conservative fallback for native text/shape/picture pages reserialized by PowerPoint. */
async function simplePage(base64: string): Promise<string | undefined> {
  const zip = await loadBoundedZip(base64)
  if (!zip.file('[Content_Types].xml') || !zip.file('ppt/presentation.xml')) return undefined
  const presentationXml = await zip.file('ppt/presentation.xml')!.async('string')
  if (!xml(presentationXml)) return undefined
  const slideSize = parser.parse(presentationXml)?.['p:presentation']?.['p:sldSz']
  if (
    !/^[1-9][0-9]{0,9}$/.test(slideSize?.['@_cx'] ?? '') ||
    !/^[1-9][0-9]{0,9}$/.test(slideSize?.['@_cy'] ?? '')
  )
    return undefined
  const slides = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  if (slides.length !== 1) return undefined
  const slidePath = slides[0]!
  const source = await zip.file(slidePath)!.async('string')
  const relPath = slidePath.replace(/\/([^/]+)$/, '/_rels/$1.rels')
  const relXml = zip.file(relPath) ? await zip.file(relPath)!.async('string') : ''
  if (!xml(source) || (relXml && !xml(relXml))) return undefined
  const relRoot = relXml ? parser.parse(relXml).Relationships : undefined
  const items = relRoot?.Relationship
  const relationships = Array.isArray(items) ? items : items ? [items] : []
  const images = new Map<string, string>()
  const relationIds = new Set<string>()
  for (const rel of relationships) {
    const id = rel?.['@_Id'],
      target = rel?.['@_Target'],
      type = rel?.['@_Type']
    if (typeof id !== 'string' || typeof target !== 'string' || typeof type !== 'string')
      return undefined
    if (relationIds.has(id)) return undefined
    relationIds.add(id)
    if (!type.endsWith('/image')) continue
    if (rel['@_TargetMode'] !== undefined || images.has(id)) return undefined
    const parts = slidePath.slice(0, slidePath.lastIndexOf('/')).split('/')
    for (const segment of target.split('/')) {
      if (segment === '..') parts.pop()
      else if (segment && segment !== '.') parts.push(segment)
    }
    const path = parts.join('/')
    if (!/^ppt\/media\/[A-Za-z0-9_.-]+\.(?:png|jpe?g)$/i.test(path) || !zip.file(path))
      return undefined
    images.set(
      id,
      `${path.slice(path.lastIndexOf('.')).toLowerCase()}:${await sha256(await zip.file(path)!.async('uint8array'))}`,
    )
  }
  const tree = ordered.parse(source) as Record<string, unknown>[]
  let pictures = 0,
    blips = 0
  const shapeIds = new Set<string>()
  const visit = (nodes: unknown): boolean => {
    if (!Array.isArray(nodes)) return true
    for (const node of nodes) {
      if (!node || typeof node !== 'object') return false
      for (const [key, child] of Object.entries(node)) {
        if (['p:graphicFrame', 'p:grpSp', 'p:cxnSp', 'p:oleObj'].includes(key)) return false
        if (key === 'p:pic') pictures++
        if (key === 'p:cNvPr') {
          const attrs = node[':@'] as Record<string, unknown> | undefined
          const id = attrs?.['@_id']
          if (
            !attrs ||
            typeof id !== 'string' ||
            !/^[1-9][0-9]{0,9}$/.test(id) ||
            Number(id) > 4294967295 ||
            shapeIds.has(id)
          )
            return false
          shapeIds.add(id)
          delete attrs['@_id']
        }
        if (key === 'a:blip') {
          const attrs = node[':@'] as Record<string, unknown> | undefined
          const id = attrs?.['@_r:embed']
          if (!attrs || typeof id !== 'string' || !images.has(id) || attrs['@_r:link']) return false
          attrs['@_r:embed'] = images.get(id)
          blips++
        } else if (key === ':@') {
          if (
            Object.keys(child as object).some(
              (attr) => attr.startsWith('@_r:') && !(attr === '@_r:embed' && 'a:blip' in node),
            )
          )
            return false
        }
        if (key !== ':@' && !visit(child)) return false
      }
    }
    return true
  }
  if (!visit(tree) || pictures !== blips || !/<p:bg(?:\s|>)/.test(source)) return undefined
  const support: [string, string][] = []
  for (const path of Object.keys(zip.files).sort()) {
    if (
      /^(?:ppt\/(?:slideLayouts|slideMasters|theme|notesSlides)\/|ppt\/media\/)/.test(path) &&
      !zip.files[path]!.dir
    )
      support.push([
        path.startsWith('ppt/media/') ? path.slice(path.lastIndexOf('.')) : path,
        await sha256(await zip.file(path)!.async('uint8array')),
      ])
  }
  support.sort(([a, x], [b, y]) => a.localeCompare(b) || x.localeCompare(y))
  return JSON.stringify([slideSize['@_cx'], slideSize['@_cy'], support, canonical(tree)])
}

export async function equivalentNativePresentationPage(
  source: string,
  exported: string,
): Promise<boolean> {
  if ((await presentationPackageDigest(source)) === (await presentationPackageDigest(exported)))
    return true
  const expected = await simplePage(source)
  return expected !== undefined && expected === (await simplePage(exported))
}
