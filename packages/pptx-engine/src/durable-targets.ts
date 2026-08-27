import {
  fingerprintSemanticValue,
  type PresentationElementType,
  type PresentationTarget,
} from '@wiswork/presentation-ops'
import type { Slide, SlideDeck, SlideElement, TextElement } from './types'

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

export function ensureElementCreationId(
  slide: Slide,
  element: SlideElement,
  factory: CreationIdFactory = defaultCreationIdFactory,
): string {
  if (element.creationId) return element.creationId
  const id = mintCreationId(factory)
  element.anchor.originalXml = setCreationIdInElementXml(element.anchor.originalXml, id)
  element.creationId = id
  slide.structureDirty = true
  return id
}

/** Remint every top-level and nested shape identity in copied XML. */
export function remintCreationIdsInXml(
  xml: string,
  factory: CreationIdFactory = defaultCreationIdFactory,
): string {
  return xml.replace(/<p:cNvPr\b[^>]*(?:\/>|>[\s\S]*?<\/p:cNvPr>)/g, (cNvPr) => {
    const id = mintCreationId(factory)
    return setCreationIdInElementXml(cNvPr, id)
  })
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
    name: element.name ?? null,
  }
  if (element.type === 'text' || element.type === 'shape') {
    const text = element as TextElement
    base.presetGeometry = text.presetGeometry ?? null
    base.fill = text.fill ?? null
    base.stroke = text.stroke ?? null
    base.text = text.text ?? null
  } else if (element.type === 'picture') {
    base.mediaRef = element.mediaRef
    base.srcRect = element.srcRect ?? null
  } else if (element.type === 'table') {
    base.colWidths = element.colWidths
    base.rowHeights = element.rowHeights
    base.rows = element.rows
  } else if (element.type === 'chart') {
    base.chart = element.chart
  } else if (element.type === 'group') {
    base.children = element.children.map(semanticElement)
  } else {
    base.kind = (element as Extract<SlideElement, { type: 'passthrough' }>).kind
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

export const fingerprintSlideElement = (element: SlideElement): Promise<string> =>
  fingerprintSemanticValue(withoutUndefined(semanticElement(element)))

export const fingerprintSlide = (slide: Slide): Promise<string> =>
  fingerprintSemanticValue(
    withoutUndefined({
      background: slide.background ?? null,
      elements: slide.elements.map(semanticElement),
    }),
  )

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
  deck: SlideDeck,
  target: PresentationTarget,
): Promise<PresentationTargetResolution> {
  const slides = deck.slides.filter((slide) => slide.durableId === target.slideId)
  if (slides.length === 0) return { status: 'conflict', code: 'target_missing' }
  if (slides.length !== 1) return { status: 'conflict', code: 'target_ambiguous' }
  const slide = slides[0]!
  if (!target.elementId) {
    if (target.expectedType) return { status: 'conflict', code: 'target_stale' }
    if (!target.expectedFingerprint) return { status: 'resolved', slide }
    try {
      const fingerprint = await fingerprintSlide(slide)
      return fingerprint === target.expectedFingerprint
        ? { status: 'resolved', slide, fingerprint }
        : { status: 'conflict', code: 'target_stale' }
    } catch {
      return { status: 'conflict', code: 'target_stale' }
    }
  }

  let elementId: string
  try {
    elementId = normalizeCreationId(target.elementId)
  } catch {
    return { status: 'conflict', code: 'target_missing' }
  }
  const matches: SlideElement[] = []
  collectCreationId(slide.elements, elementId, matches)
  if (matches.length === 0) return { status: 'conflict', code: 'target_missing' }
  if (matches.length !== 1) return { status: 'conflict', code: 'target_ambiguous' }
  const element = matches[0]!
  if (target.expectedType && presentationType(element) !== target.expectedType) {
    return { status: 'conflict', code: 'target_stale' }
  }
  let fingerprint: string
  try {
    fingerprint = await fingerprintSlideElement(element)
  } catch {
    return { status: 'conflict', code: 'target_stale' }
  }
  if (target.expectedFingerprint && fingerprint !== target.expectedFingerprint) {
    return { status: 'conflict', code: 'target_stale' }
  }
  return { status: 'resolved', slide, element, fingerprint }
}
