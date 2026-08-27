import {
  fingerprintSemanticValue,
  type PresentationElementType,
  type PresentationTarget,
} from '@wiswork/presentation-ops'
import type { Slide, SlideDeck, SlideElement, TextElement } from './types'
import { resolveTarget, type PackageArchive } from './zip'
import { getSlideNotes } from './notes'
import { canonicalizePresentationXml, MAX_PRESENTATION_XML_BYTES } from './presentation-xml'
import { readSlideAdvanceTimeXml, readSlideHiddenXml, readSlideTransitionXml } from './generate'
import { readSlideTimingXml } from './animation'

export const CREATION_ID_NAMESPACE = 'http://schemas.microsoft.com/office/drawing/2014/main'
export const CREATION_ID_EXTENSION_URI = '{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}'

export type CreationIdFactory = () => string

// Office treats this as a GUID-shaped opaque ID, not necessarily an RFC 4122 UUID.
const CREATION_ID_PATTERN = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/i

export const defaultCreationIdFactory: CreationIdFactory = () =>
  `{${globalThis.crypto.randomUUID().toUpperCase()}}`

export function normalizeCreationId(value: string): string {
  const normalized = value.startsWith('{') ? value.toUpperCase() : `{${value.toUpperCase()}}`
  if (!CREATION_ID_PATTERN.test(normalized)) throw new TypeError('Invalid DrawingML creationId')
  return normalized
}

export function readCreationId(xml: string): string | undefined {
  const identityContainer = /<p:cNvPr\b[^>]*(?:\/>|>[\s\S]*?<\/p:cNvPr>)/.exec(xml)?.[0]
  const tag = identityContainer
    ? /<a16:creationId\b[^>]*\/?\s*>/.exec(identityContainer)?.[0]
    : undefined
  const match = tag ? /\bid=(["'])(.*?)\1/.exec(tag) : null
  if (!match) return undefined
  try {
    return normalizeCreationId(match[2]!)
  } catch {
    return undefined
  }
}

function creationIdExtensionEntry(id: string): string {
  return `<a:ext uri="${CREATION_ID_EXTENSION_URI}"><a16:creationId xmlns:a16="${CREATION_ID_NAMESPACE}" id="${id}"/></a:ext>`
}

function creationIdExtension(id: string): string {
  return `<a:extLst>${creationIdExtensionEntry(id)}</a:extLst>`
}

export function setCreationIdInElementXml(xml: string, rawId: string): string {
  const id = normalizeCreationId(rawId)
  const match = /<p:cNvPr\b[^>]*(?:\/>|>[\s\S]*?<\/p:cNvPr>)/.exec(xml)
  if (!match) throw new TypeError('Element XML has no p:cNvPr identity container')
  let container = match[0]
  if (/<a16:creationId\b/.test(container)) {
    container = container.replace(/(<a16:creationId\b[^>]*\bid=)(["']).*?\2/, `$1"${id}"`)
  } else if (/<\/a:extLst>/.test(container)) {
    container = container.replace('</a:extLst>', `${creationIdExtensionEntry(id)}</a:extLst>`)
  } else if (/\/>$/.test(container)) {
    container = container.replace(/\/>$/, `>${creationIdExtension(id)}</p:cNvPr>`)
  } else {
    container = container.replace('</p:cNvPr>', `${creationIdExtension(id)}</p:cNvPr>`)
  }
  return xml.slice(0, match.index) + container + xml.slice(match.index + match[0].length)
}

export function mintCreationId(factory: CreationIdFactory = defaultCreationIdFactory): string {
  return normalizeCreationId(factory())
}

const MAX_CREATION_ID_ATTEMPTS = 32

function collectElementCreationIds(elements: readonly SlideElement[], ids: Set<string>): void {
  for (const element of elements) {
    if (element.creationId) ids.add(element.creationId)
    if (element.type === 'group') collectElementCreationIds(element.children, ids)
  }
}

export function collectSlideCreationIds(slide: Slide): Set<string> {
  const ids = new Set<string>()
  collectElementCreationIds(slide.elements, ids)
  return ids
}

export function collectDeckCreationIds(deck: SlideDeck): Set<string> {
  const ids = new Set<string>()
  for (const slide of deck.slides) collectElementCreationIds(slide.elements, ids)
  return ids
}

export function collectCreationIdsInXml(xml: string): Set<string> {
  const ids = new Set<string>()
  const pattern = /<p:cNvPr\b[^>]*(?:\/>|>[\s\S]*?<\/p:cNvPr>)/g
  for (const match of xml.matchAll(pattern)) {
    const id = readCreationId(match[0])
    if (id) ids.add(id)
  }
  return ids
}

export function mintUniqueCreationIds(
  count: number,
  reserved: ReadonlySet<string>,
  factory: CreationIdFactory = defaultCreationIdFactory,
): string[] {
  const allocated = new Set(reserved)
  const result: string[] = []
  for (let index = 0; index < count; index += 1) {
    let created: string | undefined
    for (let attempt = 0; attempt < MAX_CREATION_ID_ATTEMPTS; attempt += 1) {
      const candidate = mintCreationId(factory)
      if (!allocated.has(candidate)) {
        created = candidate
        break
      }
    }
    if (!created) throw new Error('Unable to mint a unique creationId')
    allocated.add(created)
    result.push(created)
  }
  return result
}

export function ensureElementCreationId(
  slide: Slide,
  element: SlideElement,
  factory: CreationIdFactory = defaultCreationIdFactory,
): string {
  if (element.creationId) return element.creationId
  const id = mintUniqueCreationIds(1, collectSlideCreationIds(slide), factory)[0]!
  element.anchor.originalXml = setCreationIdInElementXml(element.anchor.originalXml, id)
  element.creationId = id
  slide.structureDirty = true
  return id
}

/** Remint every top-level and nested shape identity in copied XML. */
export function remintCreationIdsInXml(
  xml: string,
  factory: CreationIdFactory = defaultCreationIdFactory,
  reserved: ReadonlySet<string> = new Set(),
): string {
  const pattern = /<p:cNvPr\b[^>]*(?:\/>|>[\s\S]*?<\/p:cNvPr>)/g
  const count = [...xml.matchAll(pattern)].length
  const ids = mintUniqueCreationIds(count, reserved, factory)
  let index = 0
  return xml.replace(pattern, (cNvPr) => setCreationIdInElementXml(cNvPr, ids[index++]!))
}

export function remintCreationIdsInXmlBatch(
  xmls: readonly string[],
  factory: CreationIdFactory = defaultCreationIdFactory,
  reserved: ReadonlySet<string> = new Set(),
): string[] {
  const pattern = /<p:cNvPr\b[^>]*(?:\/>|>[\s\S]*?<\/p:cNvPr>)/g
  const counts = xmls.map((xml) => [...xml.matchAll(pattern)].length)
  const ids = mintUniqueCreationIds(
    counts.reduce((total, count) => total + count, 0),
    reserved,
    factory,
  )
  let offset = 0
  return xmls.map((xml, xmlIndex) => {
    let local = 0
    const result = xml.replace(pattern, (cNvPr) =>
      setCreationIdInElementXml(cNvPr, ids[offset + local++]!),
    )
    offset += counts[xmlIndex]!
    return result
  })
}

export function ensureCreationIdsInXmlBatch(
  xmls: readonly string[],
  factory: CreationIdFactory = defaultCreationIdFactory,
  reserved: ReadonlySet<string> = new Set(),
): string[] {
  const pattern = /<p:cNvPr\b[^>]*(?:\/>|>[\s\S]*?<\/p:cNvPr>)/g
  const seen = new Set(reserved)
  let missing = 0
  for (const xml of xmls) {
    for (const match of xml.matchAll(pattern)) {
      const id = readCreationId(match[0])
      if (!id) {
        missing += 1
      } else {
        if (seen.has(id)) throw new Error('Duplicate creationId in inserted presentation content')
        seen.add(id)
      }
    }
  }
  const ids = mintUniqueCreationIds(missing, seen, factory)
  let index = 0
  return xmls.map((xml) =>
    xml.replace(pattern, (cNvPr) => {
      if (readCreationId(cNvPr)) return cNvPr
      return setCreationIdInElementXml(cNvPr, ids[index++]!)
    }),
  )
}

function presentationType(element: SlideElement): PresentationElementType | undefined {
  if (element.type === 'picture') return 'image'
  if (
    element.type === 'text' ||
    element.type === 'shape' ||
    element.type === 'table' ||
    element.type === 'chart' ||
    element.type === 'group'
  )
    return element.type
  return undefined
}

function semanticElement(element: SlideElement): unknown {
  const base: Record<string, unknown> = {
    type: presentationType(element) ?? element.type,
    transform: element.transform,
    placeholder: element.placeholder ?? null,
    name: element.name ?? null,
    descr: element.descr ?? null,
    nvId: element.nvId ?? null,
    connection: element.connection ?? null,
  }
  if (element.type === 'text' || element.type === 'shape') {
    const text = element as TextElement
    base.presetGeometry = text.presetGeometry ?? null
    base.adjust = text.adjust ?? null
    base.customGeometry = text.customGeometry ?? null
    base.fill = text.fill ?? null
    base.stroke = text.stroke ?? null
    base.shadow = text.shadow ?? null
    base.glow = text.glow ?? null
    base.text = text.text ?? null
  } else if (element.type === 'picture') {
    base.mediaRef = element.mediaRef
    base.srcRect = element.srcRect ?? null
    base.opacity = element.opacity ?? null
    base.softEdge = element.softEdge ?? null
    base.media = element.media ?? null
    base.presetGeometry = element.presetGeometry ?? null
    base.adjust = element.adjust ?? null
    base.fill = element.fill ?? null
    base.stroke = element.stroke ?? null
    base.shadow = element.shadow ?? null
    base.glow = element.glow ?? null
  } else if (element.type === 'table') {
    base.colWidths = element.colWidths
    base.rowHeights = element.rowHeights
    base.rows = element.rows
    base.styleFlags = element.styleFlags ?? null
  } else if (element.type === 'chart') {
    base.chart = element.chart
  } else if (element.type === 'group') {
    base.childOffset = element.childOffset ?? null
    base.children = element.children.map(semanticElement)
  } else {
    throw new TypeError('Passthrough content cannot be fingerprinted safely')
  }
  return base
}

const withoutUndefined = (value: unknown): unknown => {
  if (Array.isArray(value))
    return value.map((item) => (item === undefined ? null : withoutUndefined(item)))
  if (value && typeof value === 'object') {
    const result: Record<string, unknown> = {}
    for (const [key, item] of Object.entries(value)) {
      if (item !== undefined) result[key] = withoutUndefined(item)
    }
    return result
  }
  return value
}

export interface OpenedPresentationForFingerprint {
  deck: SlideDeck
  archive: PackageArchive
}

const digestBytes = async (bytes: Uint8Array | undefined): Promise<string> => {
  if (!bytes) throw new TypeError('Referenced presentation part cannot be fingerprinted safely')
  const digest = await globalThis.crypto.subtle.digest('SHA-256', bytes)
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('')
}

async function digestPart(archive: PackageArchive, path: string): Promise<string> {
  const bytes = archive.entries.get(path)
  if (!bytes) throw new TypeError('Referenced presentation part cannot be fingerprinted safely')
  if (/\.(?:xml|rels)$/i.test(path)) {
    if (bytes.byteLength > MAX_PRESENTATION_XML_BYTES)
      throw new TypeError('Presentation XML exceeds bounds before decoding')
    const xml = new TextDecoder().decode(bytes)
    return digestBytes(canonicalizePresentationXml(xml))
  }
  return digestBytes(bytes)
}

function collectMediaRefs(element: SlideElement, refs: Set<string>): void {
  const collectFill = (fill: { type: string; mediaRef?: string } | undefined) => {
    if (fill?.type === 'image' && fill.mediaRef) refs.add(fill.mediaRef)
  }
  if (element.type === 'passthrough')
    throw new TypeError('Passthrough content cannot be fingerprinted safely')
  if (element.type === 'text' || element.type === 'shape') {
    collectFill(element.fill)
    collectFill(element.stroke?.fill)
  } else if (element.type === 'picture') {
    refs.add(element.mediaRef)
    collectFill(element.fill)
    collectFill(element.stroke?.fill)
    if (element.media?.external)
      throw new TypeError('External media cannot be fingerprinted safely')
    if (element.media?.target) refs.add(element.media.target)
  } else if (element.type === 'table') {
    for (const row of element.rows) {
      for (const cell of row) {
        collectFill(cell.fill)
        for (const stroke of Object.values(cell.borders ?? {})) collectFill(stroke?.fill)
      }
    }
  } else if (element.type === 'group') {
    for (const child of element.children) collectMediaRefs(child, refs)
  }
}

async function mediaDigests(
  archive: PackageArchive,
  elements: readonly SlideElement[],
  extraRefs: readonly string[] = [],
): Promise<Record<string, string>> {
  const refs = new Set(extraRefs)
  for (const element of elements) collectMediaRefs(element, refs)
  const result: Record<string, string> = {}
  for (const ref of [...refs].sort()) result[ref] = await digestBytes(archive.entries.get(ref))
  return result
}

async function partGraphDigests(
  archive: PackageArchive,
  roots: readonly string[],
): Promise<Record<string, string>> {
  const pending = [...new Set(roots)].sort()
  const visited = new Set<string>()
  const result: Record<string, string> = {}
  const allowedRelationships = new Set([
    '/slideMaster',
    '/theme',
    '/notesMaster',
    '/image',
    '/chart',
    '/chartStyle',
    '/chartColorStyle',
    '/chartUserShapes',
    '/package',
    '/oleObject',
    '/media',
    '/video',
    '/audio',
  ])
  while (pending.length) {
    const path = pending.shift()!
    if (visited.has(path)) continue
    if (visited.size >= 256) throw new TypeError('Referenced presentation graph exceeds bounds')
    visited.add(path)
    result[path] = await digestPart(archive, path)
    for (const rel of archive.readRels(path).values()) {
      const suffix = rel.type.slice(rel.type.lastIndexOf('/'))
      if (!allowedRelationships.has(suffix)) {
        if (/^ppt\/charts\/chart[^/]*\.xml$/i.test(path))
          throw new TypeError('Unknown chart relationship cannot be fingerprinted safely')
        continue
      }
      if (rel.targetMode === 'External')
        throw new TypeError('External presentation parts cannot be fingerprinted safely')
      const target = resolveTarget(path, rel.target)
      if (!visited.has(target)) pending.push(target)
    }
    pending.sort()
  }
  return result
}

function relatedElementParts(
  archive: PackageArchive,
  slide: Slide,
  element: SlideElement,
): string[] {
  const roots = new Set<string>()
  const rels = archive.readRels(slide.path)
  for (const match of element.anchor.originalXml.matchAll(/\br:(?:id|embed|link)="([^"]+)"/g)) {
    const rel = rels.get(match[1]!)
    if (!rel) throw new TypeError('Referenced presentation part cannot be fingerprinted safely')
    if (rel.targetMode === 'External')
      throw new TypeError('External presentation parts cannot be fingerprinted safely')
    roots.add(resolveTarget(slide.path, rel.target))
  }
  return [...roots]
}

export async function fingerprintSlideElement(
  opened: OpenedPresentationForFingerprint,
  slide: Slide,
  element: SlideElement,
): Promise<string> {
  return fingerprintSemanticValue(
    withoutUndefined({
      element: semanticElement(element),
      deckSize: opened.deck.size,
      media: await mediaDigests(opened.archive, [element]),
      relatedParts: await partGraphDigests(
        opened.archive,
        relatedElementParts(opened.archive, slide, element),
      ),
    }),
  )
}

export async function fingerprintSlide(
  opened: OpenedPresentationForFingerprint,
  slide: Slide,
): Promise<string> {
  const chain = opened.archive.resolveSlideChain(slide.path)
  const roots = [chain.layoutPath, chain.masterPath, chain.themePath].filter(
    (path): path is string => !!path,
  )
  for (const element of [...slide.elements, ...(slide.decorations ?? [])])
    roots.push(...relatedElementParts(opened.archive, slide, element))
  const notesRel = [...opened.archive.readRels(slide.path).values()].find((rel) =>
    rel.type.endsWith('/notesSlide'),
  )
  if (notesRel && notesRel.targetMode !== 'External')
    roots.push(resolveTarget(slide.path, notesRel.target))
  return fingerprintSemanticValue(
    withoutUndefined({
      background: slide.background ?? null,
      elements: slide.elements.map(semanticElement),
      decorations: slide.decorations?.map(semanticElement) ?? null,
      hidden: readSlideHiddenXml(slide.bodyPrefix),
      transition: readSlideTransitionXml(slide.bodySuffix),
      advanceTime: readSlideAdvanceTimeXml(slide.bodySuffix),
      animations: readSlideTimingXml(slide.bodySuffix),
      notesDigest: await digestBytes(
        new TextEncoder().encode(getSlideNotes(opened.archive, slide.path)),
      ),
      partDigests: await partGraphDigests(opened.archive, roots),
      media: await mediaDigests(
        opened.archive,
        [...slide.elements, ...(slide.decorations ?? [])],
        slide.background?.type === 'image' ? [slide.background.mediaRef] : [],
      ),
      deckSize: opened.deck.size,
    }),
  )
}

type ResolvedTarget = {
  status: 'resolved'
  slide: Slide
  element?: SlideElement
  fingerprint?: string
}
type TargetConflict = {
  status: 'conflict'
  code: 'target_missing' | 'target_ambiguous' | 'target_stale'
}
export type PresentationTargetResolution = ResolvedTarget | TargetConflict

function collectCreationId(
  elements: readonly SlideElement[],
  id: string,
  matches: SlideElement[],
): void {
  for (const element of elements) {
    if (element.creationId === id) matches.push(element)
    if (element.type === 'group') collectCreationId(element.children, id, matches)
  }
}

export async function resolvePresentationTarget(
  opened: OpenedPresentationForFingerprint,
  target: PresentationTarget,
): Promise<PresentationTargetResolution> {
  const hasElement = target.elementId !== undefined
  if (hasElement) {
    if (!target.expectedType || !target.expectedFingerprint)
      return { status: 'conflict', code: 'target_stale' }
  } else if (target.expectedType || !target.expectedFingerprint) {
    return { status: 'conflict', code: 'target_stale' }
  }
  const { deck } = opened
  const slides = deck.slides.filter((slide) => slide.durableId === target.slideId)
  if (slides.length === 0) return { status: 'conflict', code: 'target_missing' }
  if (slides.length !== 1) return { status: 'conflict', code: 'target_ambiguous' }
  const slide = slides[0]!
  if (!hasElement) {
    try {
      const fingerprint = await fingerprintSlide(opened, slide)
      return fingerprint === target.expectedFingerprint
        ? { status: 'resolved', slide, fingerprint }
        : { status: 'conflict', code: 'target_stale' }
    } catch {
      return { status: 'conflict', code: 'target_stale' }
    }
  }

  let elementId: string
  try {
    elementId = normalizeCreationId(target.elementId!)
  } catch {
    return { status: 'conflict', code: 'target_missing' }
  }
  const matches: SlideElement[] = []
  collectCreationId(slide.elements, elementId, matches)
  if (matches.length === 0) return { status: 'conflict', code: 'target_missing' }
  if (matches.length !== 1) return { status: 'conflict', code: 'target_ambiguous' }
  const element = matches[0]!
  if (presentationType(element) !== target.expectedType) {
    return { status: 'conflict', code: 'target_stale' }
  }
  let fingerprint: string
  try {
    fingerprint = await fingerprintSlideElement(opened, slide, element)
  } catch {
    return { status: 'conflict', code: 'target_stale' }
  }
  if (fingerprint !== target.expectedFingerprint) {
    return { status: 'conflict', code: 'target_stale' }
  }
  return { status: 'resolved', slide, element, fingerprint }
}

/** Resolve an insertion container only. It intentionally rejects mutation preconditions and element ids. */
export async function resolvePresentationContainer(
  deck: SlideDeck,
  target: PresentationTarget,
): Promise<PresentationTargetResolution> {
  if (
    target.elementId !== undefined ||
    target.expectedType !== undefined ||
    target.expectedFingerprint !== undefined
  )
    return { status: 'conflict', code: 'target_stale' }
  const slides = deck.slides.filter((slide) => slide.durableId === target.slideId)
  if (slides.length === 0) return { status: 'conflict', code: 'target_missing' }
  if (slides.length !== 1) return { status: 'conflict', code: 'target_ambiguous' }
  return { status: 'resolved', slide: slides[0]! }
}
