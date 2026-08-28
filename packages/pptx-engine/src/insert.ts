/**
 * Element insertion — synthesizes a raw <p:sp> fragment and hangs it on
 * slide.elements.
 *
 * Naturally compatible with patch-based saving: a new element's
 * anchor.originalXml is the generated XML, which patchSlideXml includes when
 * splicing elements together; deletion is removal from the array.
 * Both change the spTree structure, driving a full-slide rebuild via
 * slide.structureDirty.
 */
import type { EmuRect, Paragraph, PictureElement, Slide, SlideElement, TextElement } from './types'
import { generateParagraphXml, generateXfrmXml } from './generate'
import { escapeXmlAttr } from './xml-utils'
import { relsPathFor } from './zip'
import type { OpenedPptx } from './index'
import {
  CREATION_ID_EXTENSION_URI,
  CREATION_ID_NAMESPACE,
  collectSlideCreationIds,
  mintUniqueCreationIds,
  readCreationId,
  type CreationIdFactory,
} from './durable-targets'

function mintCreationIdFromXml(xml: string): string {
  const id = readCreationId(xml)
  if (!id) throw new Error('Generated element is missing its creationId')
  return id
}

export interface NewIdentityOptions {
  creationIdFactory?: CreationIdFactory
}

function creationIdXml(id: string): string {
  return `<a:extLst><a:ext uri="${CREATION_ID_EXTENSION_URI}"><a16:creationId xmlns:a16="${CREATION_ID_NAMESPACE}" id="${id}"/></a:ext></a:extLst>`
}

function cNvPrXml(id: number, name: string, creationId: string, attributes = ''): string {
  return `<p:cNvPr id="${id}" name="${escapeXmlAttr(name)}"${attributes}>${creationIdXml(creationId)}</p:cNvPr>`
}

function newCreationId(slide: Slide, factory?: CreationIdFactory): string {
  return mintUniqueCreationIds(1, collectSlideCreationIds(slide), factory)[0]!
}

/**
 * 'textbox' is a special value (plain text box without prstGeom); anything else is
 * an OOXML preset geometry name (rect/roundRect/ellipse/triangle/star5/rightArrow/
 * chevron…). Presets whose polygon approximation the render layer hasn't
 * implemented fall back to a rectangle; always correct in PowerPoint.
 */
export type NewShapeKind = 'textbox' | (string & {})

export interface NewElementOptions {
  kind: NewShapeKind
  offset: EmuRect
  paragraphs?: Paragraph[]
  /** Solid shape fill (#RRGGBB); textbox has no fill by default */
  fillColor?: string
  /** Shape stroke (solid color, width in EMU) */
  stroke?: { color: string; widthEmu: number }
}

let insertCounter = 1

// ── Line / connector insertion ─────────────────────────────

/** Insertable line/connector kinds: p:cxnSp fragments with optional arrow ends */
const LINE_KINDS: Record<string, { prst: string; head?: boolean; tail?: boolean }> = {
  line: { prst: 'line' },
  lineArrow: { prst: 'straightConnector1', tail: true },
  lineArrowDouble: { prst: 'straightConnector1', head: true, tail: true },
  lineBent: { prst: 'bentConnector3' },
  lineCurved: { prst: 'curvedConnector3' },
}

export function isLineKind(kind: string): boolean {
  return Object.prototype.hasOwnProperty.call(LINE_KINDS, kind)
}

const DEFAULT_LINE_STROKE = { color: '#000000', widthEmu: 12700 }

function buildCxnSpXml(
  slide: Slide,
  opts: NewElementOptions,
  def: { prst: string; head?: boolean; tail?: boolean },
  identity?: NewIdentityOptions,
): string {
  const id = nextCNvPrId(slide)
  const creationId = newCreationId(slide, identity?.creationIdFactory)
  const name = `${
    def.prst.startsWith('bentConnector')
      ? 'Elbow Connector'
      : def.prst.startsWith('curvedConnector')
        ? 'Curved Connector'
        : 'Straight Connector'
  } ${id}`
  const o = opts.offset
  const stroke = opts.stroke ?? DEFAULT_LINE_STROKE
  const color = stroke.color.replace(/^#/, '').slice(0, 6).toUpperCase()
  const head = def.head ? '<a:headEnd type="triangle" w="med" len="med"/>' : ''
  const tail = def.tail ? '<a:tailEnd type="triangle" w="med" len="med"/>' : ''
  return (
    `<p:cxnSp><p:nvCxnSpPr>${cNvPrXml(id, name, creationId)}` +
    '<p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>' +
    `<p:spPr><a:xfrm><a:off x="${o.x}" y="${o.y}"/><a:ext cx="${o.cx}" cy="${o.cy}"/></a:xfrm>` +
    `<a:prstGeom prst="${def.prst}"><a:avLst/></a:prstGeom>` +
    `<a:ln w="${Math.round(stroke.widthEmu)}" cap="flat">` +
    `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>${head}${tail}</a:ln>` +
    '</p:spPr></p:cxnSp>'
  )
}

/** Max cNvPr id used in the slide (including new elements); new elements take max+1 */
export function nextCNvPrId(slide: Slide): number {
  let max = 1
  const scan = (xml: string) => {
    for (const m of xml.matchAll(/<p:cNvPr\s[^>]*\bid="(\d+)"/g)) {
      max = Math.max(max, Number(m[1]))
    }
  }
  scan(slide.originalXml)
  for (const el of slide.elements) scan(el.anchor.originalXml)
  return max + 1
}

export function buildSpXml(
  slide: Slide,
  opts: NewElementOptions,
  identity?: NewIdentityOptions,
): string {
  const id = nextCNvPrId(slide)
  const creationId = newCreationId(slide, identity?.creationIdFactory)
  const isTextbox = opts.kind === 'textbox'
  const name = isTextbox ? `TextBox ${id}` : `Shape ${id}`
  const o = opts.offset
  const xfrm = `<a:xfrm><a:off x="${o.x}" y="${o.y}"/><a:ext cx="${o.cx}" cy="${o.cy}"/></a:xfrm>`
  // Parser convention: has txBody and no prstGeom → 'text'; textbox omits prstGeom
  const geom = isTextbox
    ? ''
    : `<a:prstGeom prst="${escapeXmlAttr(opts.kind)}"><a:avLst/></a:prstGeom>`
  const fill = opts.fillColor
    ? `<a:solidFill><a:srgbClr val="${opts.fillColor.replace(/^#/, '').slice(0, 6).toUpperCase()}"/></a:solidFill>`
    : ''
  const ln = opts.stroke
    ? `<a:ln w="${Math.round(opts.stroke.widthEmu)}"><a:solidFill><a:srgbClr val="${opts.stroke.color.replace(/^#/, '').slice(0, 6).toUpperCase()}"/></a:solidFill></a:ln>`
    : ''
  const paras = (opts.paragraphs?.length ? opts.paragraphs : [{ runs: [{ text: '' }] }])
    .map((p) => generateParagraphXml(p))
    .join('')
  return (
    `<p:sp><p:nvSpPr>${cNvPrXml(id, name, creationId)}` +
    `<p:cNvSpPr${isTextbox ? ' txBox="1"' : ''}/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>${xfrm}${geom}${fill}${ln}</p:spPr>` +
    `<p:txBody><a:bodyPr wrap="square" rtlCol="0"/><a:lstStyle/>${paras}</p:txBody></p:sp>`
  )
}

/** Synthesize a new element and hang it on the slide; returns the model element (immediately usable by the render layer). */
export function addElement(
  slide: Slide,
  opts: NewElementOptions,
  identity?: NewIdentityOptions,
): TextElement {
  const lineDef = LINE_KINDS[opts.kind]
  if (lineDef) {
    const stroke = opts.stroke ?? DEFAULT_LINE_STROKE
    const el: TextElement = {
      id: `spnew_${(insertCounter++).toString(36)}_${Date.now().toString(36)}`,
      type: 'shape',
      anchor: {
        spIndex: slide.elements.length,
        originalXml: buildCxnSpXml(slide, opts, lineDef, identity),
        range: [0, 0],
      },
      transform: { offset: { ...opts.offset }, rot: 0, flipH: false, flipV: false },
      presetGeometry: lineDef.prst,
      fill: { type: 'none' },
      stroke: {
        fill: { type: 'solid', color: stroke.color },
        width: Math.round(stroke.widthEmu),
        ...(lineDef.head ? { headEnd: { type: 'triangle' as const } } : {}),
        ...(lineDef.tail ? { tailEnd: { type: 'triangle' as const } } : {}),
      },
    }
    el.creationId = mintCreationIdFromXml(el.anchor.originalXml)
    slide.elements.push(el)
    slide.structureDirty = true
    return el
  }
  const xml = buildSpXml(slide, opts, identity)
  const el: TextElement = {
    id: `spnew_${(insertCounter++).toString(36)}_${Date.now().toString(36)}`,
    type: opts.kind === 'textbox' ? 'text' : 'shape',
    anchor: { spIndex: slide.elements.length, originalXml: xml, range: [0, 0] },
    transform: { offset: { ...opts.offset }, rot: 0, flipH: false, flipV: false },
    ...(opts.kind !== 'textbox' ? { presetGeometry: opts.kind } : {}),
    ...(opts.fillColor ? { fill: { type: 'solid' as const, color: opts.fillColor } } : {}),
    ...(opts.stroke
      ? {
          stroke: {
            fill: { type: 'solid' as const, color: opts.stroke.color },
            width: Math.round(opts.stroke.widthEmu),
          },
        }
      : {}),
    text: { paragraphs: opts.paragraphs?.length ? opts.paragraphs : [{ runs: [{ text: '' }] }] },
  }
  el.creationId = mintCreationIdFromXml(xml)
  slide.elements.push(el)
  slide.structureDirty = true
  return el
}

// ── Table insertion (graphicFrame + a:tbl) ─────────────────────────────

export interface NewTableOptions {
  rows: number
  cols: number
  offset: EmuRect
}

/** PowerPoint's default style for new tables (Medium Style 2 - Accent 1, built-in fallback in the render layer) */
const DEFAULT_TABLE_STYLE_ID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'

/**
 * Build the table graphicFrame fragment (equal-width columns / equal-height rows,
 * default built-in style, empty cells). Insertion goes through appendRawElements
 * (materialize+reparse), reusing the existing table parsing/rendering pipeline.
 */
export function buildTableXml(
  slide: Slide,
  opts: NewTableOptions,
  identity?: NewIdentityOptions,
): string {
  const id = nextCNvPrId(slide)
  const creationId = newCreationId(slide, identity?.creationIdFactory)
  const rows = Math.max(1, Math.floor(opts.rows))
  const cols = Math.max(1, Math.floor(opts.cols))
  const colW = Math.max(1, Math.floor(opts.offset.cx / cols))
  const rowH = Math.max(1, Math.floor(opts.offset.cy / rows))
  const grid = Array.from({ length: cols }, () => `<a:gridCol w="${colW}"/>`).join('')
  const cell = '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc>'
  const trs = Array.from(
    { length: rows },
    () => `<a:tr h="${rowH}">${cell.repeat(cols)}</a:tr>`,
  ).join('')
  return (
    `<p:graphicFrame><p:nvGraphicFramePr>${cNvPrXml(id, `Table ${id}`, creationId)}` +
    '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>' +
    `<p:xfrm><a:off x="${opts.offset.x}" y="${opts.offset.y}"/><a:ext cx="${opts.offset.cx}" cy="${opts.offset.cy}"/></p:xfrm>` +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">' +
    `<a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${DEFAULT_TABLE_STYLE_ID}</a:tableStyleId></a:tblPr>` +
    `<a:tblGrid>${grid}</a:tblGrid>${trs}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
  )
}

// ── Picture insertion (media part surgery) ─────────────────────────────

const IMAGE_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
}

const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

export interface NewPictureOptions {
  /** Image bytes */
  bytes: Uint8Array
  /** Lowercase extension (png/jpg/…) */
  ext: string
  offset: EmuRect
  /** cNvPr name (default Picture N; hand-drawn ink is marked with the aislides-ink prefix) */
  name?: string
  /** cNvPr descr: editor-specific payload (e.g. ink vector points), recoverable on reopen */
  descr?: string
}

/**
 * Insert a picture: media bytes into the package + [Content_Types] Default +
 * slide rels registration + synthesized <p:pic> fragment hung on slide.elements.
 * The dataUrl is generated on demand by the caller (media resolver).
 */
/**
 * Land an image into the package: media part + Content_Types Default + slide rels.
 * Returns the new relationship id and media path (shared by picture insertion /
 * shape picture fill).
 */
export function addImageMediaAndRel(
  opened: OpenedPptx,
  slide: Slide,
  bytes: Uint8Array,
  extRaw: string,
): { rid: string; mediaPath: string } | null {
  const { archive } = opened
  const ext = extRaw.toLowerCase()
  const mime = IMAGE_MIME[ext]
  if (!mime) return null

  // 1) media part: number = current max + 1
  let maxNum = 0
  for (const path of archive.entries.keys()) {
    const m = /^ppt\/media\/image(\d+)\./.exec(path)
    if (m) maxNum = Math.max(maxNum, Number(m[1]))
  }
  const mediaPath = `ppt/media/image${maxNum + 1}.${ext}`
  archive.entries.set(mediaPath, bytes)

  // 2) [Content_Types] Default (added the first time this extension appears)
  const ctPath = '[Content_Types].xml'
  const ct = archive.readText(ctPath)
  if (ct && !new RegExp(`<Default Extension="${ext}"`).test(ct)) {
    const dflt = `<Default Extension="${ext}" ContentType="${mime}"/>`
    archive.entries.set(ctPath, Buffer.from(ct.replace('</Types>', `${dflt}</Types>`), 'utf8'))
  }

  // 3) slide rels: new rId (the rels file may not exist)
  const relsPath = relsPathFor(slide.path)
  const rels =
    archive.readText(relsPath) ??
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
  let maxRid = 0
  for (const m of rels.matchAll(/Id="rId(\d+)"/g)) maxRid = Math.max(maxRid, Number(m[1]))
  const rid = `rId${maxRid + 1}`
  const relXml = `<Relationship Id="${rid}" Type="${IMAGE_REL_TYPE}" Target="../media/image${maxNum + 1}.${ext}"/>`
  archive.entries.set(
    relsPath,
    Buffer.from(rels.replace('</Relationships>', `${relXml}</Relationships>`), 'utf8'),
  )
  return { rid, mediaPath }
}

export function addPicture(
  opened: OpenedPptx,
  slide: Slide,
  opts: NewPictureOptions,
  identity?: NewIdentityOptions,
): PictureElement | null {
  // Identity is allocated before media/relationship writes so exhaustion is mutation-free.
  const creationId = newCreationId(slide, identity?.creationIdFactory)
  const added = addImageMediaAndRel(opened, slide, opts.bytes, opts.ext)
  if (!added) return null
  const { rid, mediaPath } = added

  // 4) <p:pic> fragment
  const id = nextCNvPrId(slide)
  const name = opts.name ?? `Picture ${id}`
  const descrAttr = opts.descr ? ` descr="${escapeXmlAttr(opts.descr)}"` : ''
  const xml =
    `<p:pic><p:nvPicPr>${cNvPrXml(id, name, creationId, descrAttr)}` +
    '<p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
    `<p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>${generateXfrmXml({ offset: opts.offset, rot: 0, flipH: false, flipV: false })}` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'

  const el: PictureElement = {
    id: `picnew_${(insertCounter++).toString(36)}_${Date.now().toString(36)}`,
    type: 'picture',
    anchor: { spIndex: slide.elements.length, originalXml: xml, range: [0, 0] },
    transform: { offset: { ...opts.offset }, rot: 0, flipH: false, flipV: false },
    name,
    ...(opts.descr ? { descr: opts.descr } : {}),
    mediaRef: mediaPath,
    creationId: mintCreationIdFromXml(xml),
  }
  slide.elements.push(el)
  slide.structureDirty = true
  return el
}

/** Delete by element id; returns whether anything was removed. */
export function deleteElement(slide: Slide, elementId: string): boolean {
  const idx = slide.elements.findIndex((e) => e.id === elementId)
  if (idx < 0) return false
  slide.elements.splice(idx, 1)
  slide.structureDirty = true
  return true
}

const OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const XML_NS = 'http://www.w3.org/XML/1998/namespace'
const XMLNS_NS = 'http://www.w3.org/2000/xmlns/'
const MAX_DELETE_XML_LENGTH = 16 * 1024 * 1024
const MAX_DELETE_XML_TAGS = 200_000

interface XmlAttributeToken {
  prefix: string
  localName: string
  namespaceUri?: string
  value: string
}

interface XmlElementToken {
  start: number
  end: number
  localName: string
  namespaceUri?: string
  attributes: XmlAttributeToken[]
}

const splitQName = (name: string): { prefix: string; localName: string } => {
  const separator = name.indexOf(':')
  return separator < 0
    ? { prefix: '', localName: name }
    : { prefix: name.slice(0, separator), localName: name.slice(separator + 1) }
}

const inspectXml = (
  xml: string,
  inheritedNamespaces: ReadonlyMap<string, string>,
  visit: (element: XmlElementToken) => void,
): void => {
  if (xml.length > MAX_DELETE_XML_LENGTH) throw new Error('XML inspection bound exceeded')
  const stack: Array<{ namespaces: Map<string, string>; name?: string }> = [
    { namespaces: new Map([...inheritedNamespaces, ['xml', XML_NS]]) },
  ]
  const tagPattern =
    /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<![^>]*>|<\/?([A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?)([^<>]*?)\/?\s*>/g
  let tags = 0
  let cursor = 0
  for (const match of xml.matchAll(tagPattern)) {
    if (xml.slice(cursor, match.index).includes('<')) throw new Error('Unrecognized XML token')
    if (++tags > MAX_DELETE_XML_TAGS) throw new Error('XML tag bound exceeded')
    const raw = match[0]
    cursor = match.index! + raw.length
    if (raw.startsWith('<!') || raw.startsWith('<?')) continue
    if (raw.startsWith('</')) {
      const frame = stack.pop()
      if (!frame?.name || frame.name !== match[1]) throw new Error('Malformed XML nesting')
      continue
    }
    const namespaces = new Map(stack.at(-1)!.namespaces)
    const attributes: Array<{ name: string; value: string }> = []
    const attributePattern = /([A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?)\s*=\s*(["'])([^"']*)\2/g
    const rawAttributes = match[2] ?? ''
    const parsedAttributeSpans: [number, number][] = []
    for (const attribute of rawAttributes.matchAll(attributePattern)) {
      parsedAttributeSpans.push([attribute.index!, attribute.index! + attribute[0].length])
      const name = attribute[1]!
      const value = attribute[3]!
      if (name === 'xmlns') {
        if (value === XML_NS || value === XMLNS_NS)
          throw new Error('Invalid default namespace declaration')
        namespaces.set('', value)
      } else if (name.startsWith('xmlns:')) {
        const prefix = name.slice(6)
        if (
          !prefix ||
          prefix === 'xmlns' ||
          !value ||
          value === XMLNS_NS ||
          (prefix === 'xml' && value !== XML_NS) ||
          (prefix !== 'xml' && value === XML_NS)
        )
          throw new Error('Invalid namespace declaration')
        namespaces.set(prefix, value)
      } else attributes.push({ name, value })
    }
    if (applyRemovals(rawAttributes, parsedAttributeSpans).trim())
      throw new Error('Malformed XML attributes')
    const qname = splitQName(match[1]!)
    if (qname.prefix && !namespaces.has(qname.prefix)) throw new Error('Unbound element prefix')
    visit({
      start: match.index!,
      end: match.index! + raw.length,
      localName: qname.localName,
      namespaceUri: namespaces.get(qname.prefix),
      attributes: attributes.map(({ name, value }) => {
        const attributeName = splitQName(name)
        if (attributeName.prefix && !namespaces.has(attributeName.prefix))
          throw new Error('Unbound attribute prefix')
        return {
          ...attributeName,
          ...(attributeName.prefix ? { namespaceUri: namespaces.get(attributeName.prefix) } : {}),
          value,
        }
      }),
    })
    if (!/\/\s*>$/.test(raw)) stack.push({ namespaces, name: match[1]! })
  }
  if (xml.slice(cursor).includes('<')) throw new Error('Unrecognized XML token')
  if (stack.length !== 1) throw new Error('Unclosed XML element')
}

const applyRemovals = (xml: string, removals: readonly [number, number][]): string => {
  let patched = xml
  for (const [start, end] of [...removals].sort((a, b) => b[0] - a[0]))
    patched = patched.slice(0, start) + patched.slice(end)
  return patched
}

interface SlideXmlSegment {
  start: number
  end: number
  xml: string
  element?: SlideElement
}

export interface DeleteElementCleanupMetrics {
  segmentComparisons: number
  segmentCount: number
  tokenCount: number
}

const reconstructSlideSegments = (
  slide: Slide,
  excluded?: SlideElement,
): {
  xml: string
  segments: SlideXmlSegment[]
  elementSegments: Map<SlideElement, SlideXmlSegment>
} => {
  const segments: SlideXmlSegment[] = []
  const elementSegments = new Map<SlideElement, SlideXmlSegment>()
  let xml = ''
  const append = (value: string, element?: SlideElement) => {
    if (!value) return
    const start = xml.length
    xml += value
    const segment = { start, end: xml.length, xml: value, ...(element ? { element } : {}) }
    segments.push(segment)
    if (element) elementSegments.set(element, segment)
  }
  append(slide.bodyPrefix)
  for (const element of slide.elements) {
    if (element === excluded) continue
    append(element.anchor.originalXml, element)
    if (element.anchor.gapAfter) append(element.anchor.gapAfter)
  }
  append(slide.bodySuffix)
  return { xml, segments, elementSegments }
}

const validateSlideSegments = (segments: readonly SlideXmlSegment[], xmlLength: number): void => {
  let previousEnd = 0
  for (const segment of segments) {
    if (segment.start !== previousEnd || segment.end <= segment.start)
      throw new Error('Invalid slide XML segment ranges')
    previousEnd = segment.end
  }
  if (previousEnd !== xmlLength) throw new Error('Incomplete slide XML segment coverage')
}

const createMonotonicSegmentLocator = (
  segments: readonly SlideXmlSegment[],
  xmlLength: number,
  metrics?: DeleteElementCleanupMetrics,
): ((start: number, end: number) => SlideXmlSegment | undefined) => {
  validateSlideSegments(segments, xmlLength)
  let index = 0
  let previousStart = -1
  return (start, end) => {
    if (start < previousStart || end < start) throw new Error('Non-monotonic XML token range')
    previousStart = start
    while (index < segments.length && start >= segments[index]!.end) {
      if (metrics) metrics.segmentComparisons++
      index++
    }
    if (index >= segments.length) return undefined
    if (metrics) metrics.segmentComparisons++
    const segment = segments[index]!
    return start >= segment.start && end <= segment.end ? segment : undefined
  }
}

/**
 * Authoritative package-aware deletion. It removes relationship records referenced only by the
 * deleted fragment and detaches surviving connectors from the deleted cNvPr id. Orphan target
 * parts are intentionally retained: deleting them is unsafe when another package part shares the
 * same media/chart target, while an unreferenced OPC part is valid and recoverable.
 */
export function deleteElementWithCleanup(
  opened: OpenedPptx,
  slide: Slide,
  elementId: string,
  metrics?: DeleteElementCleanupMetrics,
): boolean {
  const index = slide.elements.findIndex((element) => element.id === elementId)
  if (index < 0) return false
  const victim = slide.elements[index]!
  try {
    let nvId: string | undefined
    const candidateRIds = new Set<string>()
    const current = reconstructSlideSegments(slide)
    validateSlideSegments(current.segments, current.xml.length)
    const victimSegment = current.elementSegments.get(victim)
    if (!victimSegment) return false
    if (metrics) metrics.segmentCount += current.segments.length
    inspectXml(current.xml, new Map(), (element) => {
      if (metrics) metrics.tokenCount++
      if (element.start < victimSegment.start || element.end > victimSegment.end) return
      if (element.localName === 'cNvPr') {
        const id = element.attributes.find(
          (attribute) => !attribute.prefix && attribute.localName === 'id',
        )?.value
        if (id && /^\d+$/.test(id)) nvId = id
      }
      for (const attribute of element.attributes)
        if (
          attribute.namespaceUri === OFFICE_REL_NS &&
          (attribute.localName === 'id' ||
            attribute.localName === 'embed' ||
            attribute.localName === 'link')
        )
          candidateRIds.add(attribute.value)
    })
    if (!nvId) return false

    const elementRemovals = new Map<SlideElement, [number, number][]>()
    const survivingRIds = new Set<string>()
    const surviving = reconstructSlideSegments(slide, victim)
    if (metrics) metrics.segmentCount += surviving.segments.length
    const locateSurvivingSegment = createMonotonicSegmentLocator(
      surviving.segments,
      surviving.xml.length,
      metrics,
    )
    inspectXml(surviving.xml, new Map(), (element) => {
      if (metrics) metrics.tokenCount++
      const segment = locateSurvivingSegment(element.start, element.end)
      if (!segment) throw new Error('XML token crosses slide segment boundary')
      for (const attribute of element.attributes) {
        if (
          attribute.namespaceUri === OFFICE_REL_NS &&
          (attribute.localName === 'id' ||
            attribute.localName === 'embed' ||
            attribute.localName === 'link')
        )
          survivingRIds.add(attribute.value)
        if (!attribute.prefix && attribute.localName === 'spid' && attribute.value === nvId)
          throw new Error('Unsupported shape reference')
      }
      if (
        element.namespaceUri === DRAWING_NS &&
        (element.localName === 'stCxn' || element.localName === 'endCxn') &&
        element.attributes.some(
          (attribute) =>
            !attribute.prefix && attribute.localName === 'id' && attribute.value === nvId,
        )
      ) {
        if (!segment?.element) throw new Error('Connector reference is outside an element')
        const removals = elementRemovals.get(segment.element) ?? []
        removals.push([element.start - segment.start, element.end - segment.start])
        elementRemovals.set(segment.element, removals)
      }
    })
    const patchedElements = new Map<SlideElement, string>()
    for (const [element, removals] of elementRemovals)
      patchedElements.set(element, applyRemovals(element.anchor.originalXml, removals))

    const removableRIds = new Set([...candidateRIds].filter((id) => !survivingRIds.has(id)))
    const relsPath = relsPathFor(slide.path)
    const rels = opened.archive.readText(relsPath)
    let patchedRels: string | undefined
    if (removableRIds.size) {
      if (!rels) return false
      const relRemovals: [number, number][] = []
      inspectXml(rels, new Map([['', PACKAGE_REL_NS]]), (element) => {
        if (element.namespaceUri !== PACKAGE_REL_NS || element.localName !== 'Relationship') return
        const id = element.attributes.find(
          (attribute) => !attribute.prefix && attribute.localName === 'Id',
        )?.value
        if (id && removableRIds.has(id)) relRemovals.push([element.start, element.end])
      })
      if (relRemovals.length !== removableRIds.size) return false
      patchedRels = applyRemovals(rels, relRemovals)
    }

    const patchedRelsBytes =
      patchedRels === undefined ? undefined : Buffer.from(patchedRels, 'utf8')

    // Commit only after every namespace/reference check and patch plan has succeeded.
    for (const [element, xml] of patchedElements) element.anchor.originalXml = xml
    const syncConnections = (element: SlideElement): void => {
      if (element.connection?.start?.id === Number(nvId)) delete element.connection.start
      if (element.connection?.end?.id === Number(nvId)) delete element.connection.end
      if (element.connection && !element.connection.start && !element.connection.end)
        delete element.connection
      if (element.type === 'group') for (const child of element.children) syncConnections(child)
    }
    for (const element of patchedElements.keys()) syncConnections(element)
    if (patchedRelsBytes !== undefined) opened.archive.entries.set(relsPath, patchedRelsBytes)
    slide.elements.splice(index, 1)
    slide.structureDirty = true
    return true
  } catch {
    return false
  }
}

// ── Grouping (p:grpSp) ──────────────────────────────────────────────────────

/**
 * Compute the bounding box of a set of elements (slide coordinates, EMU).
 * Ignores rotation: uses the axis-aligned bounding box of each element's offset
 * rect.
 */
export function calcBoundingBox(elements: SlideElement[]): EmuRect {
  let minX = Infinity,
    minY = Infinity,
    maxX = -Infinity,
    maxY = -Infinity
  for (const el of elements) {
    const o = el.transform.offset
    minX = Math.min(minX, o.x)
    minY = Math.min(minY, o.y)
    maxX = Math.max(maxX, o.x + o.cx)
    maxY = Math.max(maxY, o.y + o.cy)
  }
  return { x: minX, y: minY, cx: maxX - minX, cy: maxY - minY }
}

/**
 * Build the <p:grpSp> XML fragment.
 *
 * OOXML conventions (ECMA 376 §19.3.1.22):
 *  - grpSpPr/xfrm describes the group's position and size on the slide (<a:off>/<a:ext>)
 *  - grpSpPr/xfrm/chOff + chExt define the child coordinate system's origin and size
 *  - This implementation sets chOff == bbox.xy and chExt == bbox.cxcy, i.e. the child
 *    coordinate system is 1:1 with the slide's → child elements can reuse their
 *    original slide coordinates inside the group with no transform
 *  - childrenXml: concatenation of each child's raw XML fragment (passthrough
 *    children keep their original bytes)
 */
export function buildGrpSpXml(
  slide: Slide,
  bbox: EmuRect,
  childrenXml: string,
  identity?: NewIdentityOptions,
): string {
  const id = nextCNvPrId(slide)
  const creationId = newCreationId(slide, identity?.creationIdFactory)
  const name = `Group ${id}`
  const { x, y, cx, cy } = bbox
  const grpXfrm =
    `<a:xfrm>` +
    `<a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/>` +
    `<a:chOff x="${x}" y="${y}"/><a:chExt cx="${cx}" cy="${cy}"/>` +
    `</a:xfrm>`
  return (
    `<p:grpSp>` +
    `<p:nvGrpSpPr>${cNvPrXml(id, name, creationId)}` +
    `<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr>${grpXfrm}</p:grpSpPr>` +
    childrenXml +
    `</p:grpSp>`
  )
}
