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
  const themeColorModified = (
    (tags(fillTree, 'a:schemeClr')[0]?.['a:schemeClr'] as Node[] | undefined) ?? []
  ).some((node) => Object.keys(node).some((key) => key !== ':@' && key !== '#text'))
  const typeface = attr(tags(propertyTree, 'a:latin')[0], 'typeface')
  return {
    ...(size && /^\d+$/.test(size) ? { sizePt: Number(size) / 100 } : {}),
    ...(color && /^[0-9a-fA-F]{6}$/.test(color) ? { color: `#${color.toUpperCase()}` } : {}),
    ...(themeColor ? { themeColor: themeColor.slice(0, 64) } : {}),
    ...(themeColorModified ? { themeColorModified: true } : {}),
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
function knownFont(...layers: Array<Record<string, string | number | boolean>>) {
  const result: Record<string, string | number | boolean> = {}
  for (const layer of layers) {
    if ('color' in layer) {
      delete result.themeColor
      delete result.themeColorModified
    }
    if ('themeColor' in layer) {
      delete result.color
      delete result.themeColorModified
    }
    Object.assign(result, layer)
  }
  return result
}

export interface RichTextRun {
  text: string
  directFont: Record<string, string | number | boolean>
  /** Known local formatting only; theme and master/layout inheritance remain unresolved. */
  knownFont: Record<string, string | number | boolean>
  /** Exact theme RGB only when the page's linked layout/master/theme chain is unambiguous. */
  resolvedThemeColor?: string
  /** Exact Latin theme typeface for +mj-lt or +mn-lt only. */
  resolvedThemeTypeface?: string
}
export interface RichTextShape {
  packageShapeId: string
  name: string
  paragraphs: Array<{
    alignment?: string
    knownAlignment?: string
    listStyleFont?: Record<string, string | number | boolean>
    paragraphDefaultFont?: Record<string, string | number | boolean>
    endParagraphFont?: Record<string, string | number | boolean>
    runs: RichTextRun[]
  }>
}

function stableShape(value: unknown, preserveText = false): unknown {
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
      .map((part) => stableShape(part, preserveText))
  if (!value || typeof value !== 'object') return value
  return Object.fromEntries(
    Object.entries(value as Node)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([key, child]) => [key, stableShape(child, preserveText || key === 'a:t')]),
  )
}

async function shapeHash(shape: Node, relationships: unknown): Promise<string> {
  const data = new TextEncoder().encode(JSON.stringify([stableShape(shape), relationships]))
  const digest = new Uint8Array(await crypto.subtle.digest('SHA-256', data))
  return Array.from(digest, (byte) => byte.toString(16).padStart(2, '0')).join('')
}

/** Exact XML guard for top-level native shapes and their referenced resources. */
export async function inspectPowerPointTextShapeFingerprints(
  base64: string,
  shapeIds: string[],
  signal?: AbortSignal,
  allowNoTextBody = false,
): Promise<Record<string, { exact: string; content: string; formatting: string }>> {
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
  const slide = xml(await zip.file(slides[0]!)!.async('string'))
  const slidePath = slides[0]!
  const slash = slidePath.lastIndexOf('/')
  const relationshipPart = `${slidePath.slice(0, slash)}/_rels/${slidePath.slice(slash + 1)}.rels`
  const relationshipFile = zip.file(relationshipPart)
  const relationships = relationshipFile
    ? tags(xml(await relationshipFile.async('string')), 'Relationship')
    : []
  const resourceDigests = new Map<string, string>()
  const referenced = (shape: Node): string[] => {
    const ids = new Set<string>()
    const visit = (value: unknown): void => {
      if (Array.isArray(value)) {
        value.forEach(visit)
        return
      }
      if (!value || typeof value !== 'object') return
      for (const [key, child] of Object.entries(value as Node)) {
        if (
          key.startsWith('@_r:') &&
          typeof child === 'string' &&
          /^rId[A-Za-z0-9_-]+$/.test(child)
        )
          ids.add(child)
        else visit(child)
      }
    }
    visit(shape)
    return [...ids].sort()
  }
  const found: Record<string, { exact: string; content: string; formatting: string }> =
    Object.create(null)
  const shapeTree = tags(slide, 'p:spTree')[0]?.['p:spTree'] as Node[] | undefined
  if (!shapeTree) throw Error('office_api_unsupported')
  for (const original of shapeTree) {
    const kind = ['p:sp', 'p:cxnSp', 'p:grpSp', 'p:graphicFrame', 'p:pic', 'p:contentPart'].find(
      (key) => Array.isArray(original[key]),
    )
    if (!kind) continue
    const children = original[kind] as Node[]
    const id = attr(tags(children, 'p:cNvPr')[0], 'id')
    if (!id || !shapeIds.includes(id)) continue
    if (
      Object.hasOwn(found, id) ||
      (!allowNoTextBody && (kind !== 'p:sp' || !tags(children, 'p:txBody').length))
    )
      throw Error('office_api_unsupported')
    const linked = []
    for (const relationId of referenced(original)) {
      const matched = relationships.filter(
        (entry) => (entry[':@'] as Node | undefined)?.['@_Id'] === relationId,
      )
      if (matched.length !== 1) throw Error('office_api_unsupported')
      const relation = matched[0]!
      let resourceDigest: string | undefined
      if (attr(relation, 'TargetMode') !== 'External') {
        const path = relatedPath(slidePath, attr(relation, 'Target') ?? '')
        const resource = path && zip.file(path)
        if (!resource) throw Error('office_api_unsupported')
        resourceDigest = resourceDigests.get(path!)
        if (!resourceDigest) {
          const bytes = await resource.async('uint8array')
          const digest = new Uint8Array(
            await crypto.subtle.digest('SHA-256', new Uint8Array(bytes)),
          )
          resourceDigest = Array.from(digest, (byte) => byte.toString(16).padStart(2, '0')).join('')
          resourceDigests.set(path!, resourceDigest)
        }
      }
      linked.push([relationId, stableShape(relation), resourceDigest])
    }
    const exact = await shapeHash(original, linked)
    const shape = structuredClone(original)
    const propertyTag = kind === 'p:grpSp' ? 'p:grpSpPr' : 'p:spPr'
    for (const properties of (shape[kind] as Node[]).filter((node) =>
      Object.hasOwn(node, propertyTag),
    ))
      for (const transform of tags(properties[propertyTag] as Node[], 'a:xfrm'))
        transform['a:xfrm'] = (transform['a:xfrm'] as Node[]).filter(
          (part) => !Object.hasOwn(part, 'a:off') && !Object.hasOwn(part, 'a:ext'),
        )
    const content = await shapeHash(shape, linked)
    const formattingShape = structuredClone(original)
    for (const text of kind === 'p:sp' ? tags(formattingShape[kind] as Node[], 'a:t') : []) {
      text['a:t'] = [{ '#text': '' }]
      const attributes = text[':@'] as Node | undefined
      if (attributes) delete attributes['@_xml:space']
    }
    found[id] = { exact, content, formatting: await shapeHash(formattingShape, linked) }
  }
  if (Object.keys(found).length !== shapeIds.length) throw Error('office_api_unsupported')
  if (signal?.aborted) throw Error('cancelled')
  return found
}

type Zip = Awaited<ReturnType<typeof loadBoundedZip>>
function relationshipPath(part: string): string {
  const index = part.lastIndexOf('/')
  return `${part.slice(0, index)}/_rels/${part.slice(index + 1)}.rels`
}
function relatedPath(source: string, target: string): string | undefined {
  if (
    !target ||
    target.startsWith('/') ||
    target.includes('\\') ||
    target.includes('?') ||
    target.includes('#')
  )
    return
  const parts = source.split('/').slice(0, -1)
  for (const segment of target.split('/')) {
    if (segment === '..') parts.pop()
    else if (segment !== '.' && segment !== '') parts.push(segment)
    else if (segment === '') return
    if (!parts.length) return
  }
  const path = parts.join('/')
  return path.startsWith('ppt/') && /^[A-Za-z0-9_./-]+$/.test(path) ? path : undefined
}
async function related(zip: Zip, source: string, kind: string): Promise<string | undefined> {
  const file = zip.file(relationshipPath(source))
  if (!file) return
  const relationships = tags(xml(await file.async('string')), 'Relationship').filter(
    (node) =>
      attr(node, 'Type') ===
        `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}` &&
      attr(node, 'TargetMode') !== 'External',
  )
  if (relationships.length !== 1) return
  const path = relatedPath(source, attr(relationships[0], 'Target') ?? '')
  return path && zip.file(path) ? path : undefined
}
async function linkedThemeStyle(
  zip: Zip,
  slidePath: string,
): Promise<{
  colors: Record<string, string>
  typefaces: Record<string, string>
}> {
  const empty = { colors: {}, typefaces: {} }
  const layoutPath = await related(zip, slidePath, 'slideLayout')
  const masterPath = layoutPath && (await related(zip, layoutPath, 'slideMaster'))
  const themePath = masterPath && (await related(zip, masterPath, 'theme'))
  if (!layoutPath || !masterPath || !themePath) return empty
  const layout = xml(await zip.file(layoutPath)!.async('string'))
  const master = xml(await zip.file(masterPath)!.async('string'))
  const theme = xml(await zip.file(themePath)!.async('string'))
  const override = tags(layout, 'a:overrideClrMapping')[0]
  const mapping = override ?? tags(master, 'p:clrMap')[0]
  const scheme = tags(theme, 'a:clrScheme')[0]
  const colors: Record<string, string> = {}
  const typefaces: Record<string, string> = {}
  const fontScheme = tags(theme, 'a:fontScheme')[0]
  const fontChildren = (fontScheme?.['a:fontScheme'] as Node[] | undefined) ?? []
  for (const [symbol, tag] of [
    ['+mj-lt', 'a:majorFont'],
    ['+mn-lt', 'a:minorFont'],
  ] as const) {
    const font = tags(fontChildren, tag)[0]
    const face = attr(tags((font?.[tag] as Node[] | undefined) ?? [], 'a:latin')[0], 'typeface')
    if (face && face.length <= 256 && !Array.from(face).some((char) => char.charCodeAt(0) < 32))
      typefaces[symbol] = face
  }
  if (!mapping || !scheme) return { colors, typefaces }
  for (const key of [
    'accent1',
    'accent2',
    'accent3',
    'accent4',
    'accent5',
    'accent6',
    'hlink',
    'folHlink',
  ]) {
    const mapped = attr(mapping, key)
    if (!mapped || !/^(accent[1-6]|hlink|folHlink)$/.test(mapped)) continue
    const entry = tags(scheme['a:clrScheme'] as Node[], `a:${mapped}`)[0]
    const children = (entry?.[`a:${mapped}`] as Node[] | undefined) ?? []
    const rgbNode = tags(children, 'a:srgbClr')[0]
    const rgb = attr(rgbNode, 'val')
    if (
      rgb &&
      /^[0-9a-fA-F]{6}$/.test(rgb) &&
      children.length === 1 &&
      !((rgbNode?.['a:srgbClr'] as Node[] | undefined) ?? []).length
    )
      colors[key] = `#${rgb.toUpperCase()}`
  }
  return { colors, typefaces }
}

/** Local formatting plus exact linked theme RGB where provable; layout/master font inheritance remains unresolved. */
export async function inspectPowerPointRichText(
  base64: string,
  signal?: AbortSignal,
): Promise<{ shapes: RichTextShape[]; inheritanceResolved: false }> {
  const zip = await loadBoundedZip(base64, signal, true, 8 * 1024 * 1024)
  const slides = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  if (slides.length !== 1) throw new Error('office_api_unsupported')
  const slide = xml(await zip.file(slides[0]!)!.async('string'))
  const themeStyle = await linkedThemeStyle(zip, slides[0]!)
  const shapes: RichTextShape[] = []
  let runCount = 0
  let textLength = 0
  for (const shape of tags(slide, 'p:sp')) {
    const body = tags(shape['p:sp'] as Node[], 'p:txBody')[0]
    if (!body) continue
    const identity = tags(shape['p:sp'] as Node[], 'p:cNvPr')[0]
    const packageShapeId = attr(identity, 'id')
    if (!packageShapeId || !/^\d+$/.test(packageShapeId)) throw new Error('office_api_unsupported')
    const listStyle = tags(body['p:txBody'] as Node[], 'a:lstStyle')[0]
    const listChildren = (listStyle?.['a:lstStyle'] as Node[] | undefined) ?? []
    const paragraphs = tags(body['p:txBody'] as Node[], 'a:p').map((paragraph) => {
      const children = paragraph['a:p'] as Node[]
      const paragraphProperties = tags(children, 'a:pPr')[0]
      const alignment = attr(paragraphProperties, 'algn')
      const levelValue = attr(paragraphProperties, 'lvl') ?? '0'
      const level = /^[0-8]$/.test(levelValue) ? Number(levelValue) : undefined
      const defaultProperties = tags(listChildren, 'a:defPPr')[0]
      const levelProperties =
        level === undefined ? undefined : tags(listChildren, `a:lvl${level + 1}pPr`)[0]
      const defaultFont = defaultProperties
        ? directFont(defaultProperties['a:defPPr'] as Node[])
        : {}
      const levelFont = levelProperties
        ? directFont(levelProperties[`a:lvl${level! + 1}pPr`] as Node[])
        : {}
      const listStyleFont = knownFont(defaultFont, levelFont)
      const knownAlignment =
        alignment ?? attr(levelProperties, 'algn') ?? attr(defaultProperties, 'algn')
      const paragraphDefaultFont = paragraphProperties
        ? directFont(paragraphProperties['a:pPr'] as Node[])
        : {}
      const endParagraphFont = directFont(tags(children, 'a:endParaRPr'))
      const runs: RichTextRun[] = []
      const addRun = (text: string, font: RichTextRun['directFont']) => {
        const resolved = knownFont(listStyleFont, paragraphDefaultFont, font)
        const themeColor = resolved.themeColor
        runs.push({
          text,
          directFont: font,
          knownFont: resolved,
          ...(typeof themeColor === 'string' &&
          !resolved.themeColorModified &&
          themeStyle.colors[themeColor]
            ? { resolvedThemeColor: themeStyle.colors[themeColor] }
            : {}),
          ...(typeof resolved.typeface === 'string' && themeStyle.typefaces[resolved.typeface]
            ? { resolvedThemeTypeface: themeStyle.typefaces[resolved.typeface] }
            : {}),
        })
      }
      for (const child of children) {
        const entry = child['a:r'] ?? child['a:fld']
        if (Array.isArray(entry)) {
          const text = tags(entry as Node[], 'a:t')
            .map((node) => contents(node['a:t'] as Node[]))
            .join('')
          const font = directFont(entry as Node[])
          addRun(text, font)
        } else if (Array.isArray(child['a:br'])) {
          const font = directFont(child['a:br'] as Node[])
          addRun('\n', font)
        } else if (Array.isArray(child['a:tab'])) {
          addRun('\t', {})
        }
      }
      runCount += runs.length
      textLength += runs.reduce((total, run) => total + run.text.length, 0)
      return {
        ...(alignment ? { alignment } : {}),
        ...(knownAlignment ? { knownAlignment } : {}),
        ...(Object.keys(listStyleFont).length ? { listStyleFont } : {}),
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
