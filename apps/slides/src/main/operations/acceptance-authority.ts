import { fingerprintPresentation, type SlideElement } from '@wiswork/pptx-engine'
import { fingerprintSemanticValue } from '@wiswork/presentation-ops'
import { EMU_PER_PX_96 } from '@wiswork/pptx-render'
import type {
  SlidesAcceptanceAuthorityRequest,
  SlidesAcceptanceAuthoritySnapshot,
} from '../../shared/ipc'
import type { Session } from '../session-state'

const roleOf = (element: SlideElement): 'title' | 'body' | 'emphasis' | undefined => {
  if (element.placeholder === 'title' || element.placeholder === 'ctrTitle') return 'title'
  if (element.placeholder === 'subTitle') return 'emphasis'
  if (element.placeholder === 'body') return 'body'
  return undefined
}

const uniform = <T>(values: T[]): T | undefined =>
  values.length > 0 && values.every((value) => value === values[0]) ? values[0] : undefined

async function elementFact(slideId: string, element: SlideElement, locked: boolean) {
  if (!element.creationId) return undefined
  const role = roleOf(element)
  const targetDigest = await fingerprintSemanticValue({ slideId, elementId: element.creationId })
  const properties: Record<string, string | number | boolean | null> = {
    x: element.transform.offset.x / EMU_PER_PX_96,
    y: element.transform.offset.y / EMU_PER_PX_96,
    width: element.transform.offset.cx / EMU_PER_PX_96,
    height: element.transform.offset.cy / EMU_PER_PX_96,
  }
  if ('fill' in element && element.fill?.type === 'solid')
    properties.fill_color = element.fill.color
  if ('stroke' in element && element.stroke?.fill.type === 'solid')
    properties.stroke_color = element.stroke.fill.color
  if ((element.type === 'text' || element.type === 'shape') && element.text) {
    const runs = element.text.paragraphs.flatMap((paragraph) => paragraph.runs)
    const color = uniform(runs.map((run) => run.color).filter((value): value is string => !!value))
    const fontSize = uniform(
      runs.map((run) => run.fontSize).filter((value): value is number => value !== undefined),
    )
    const fontFamily = uniform(
      runs.map((run) => run.fontFamily).filter((value): value is string => !!value),
    )
    const bold = uniform(runs.map((run) => run.bold ?? false))
    const italic = uniform(runs.map((run) => run.italic ?? false))
    if (color !== undefined && runs.every((run) => run.color !== undefined))
      properties.color = color
    if (fontSize !== undefined && runs.every((run) => run.fontSize !== undefined))
      properties.font_size = fontSize
    if (fontFamily !== undefined && runs.every((run) => run.fontFamily !== undefined))
      properties.font_family = fontFamily
    if (bold !== undefined) properties.bold = bold
    if (italic !== undefined) properties.italic = italic
  }
  return {
    targetToken: `target:${targetDigest.slice('sha256:'.length)}`,
    ...(role ? { role } : {}),
    locked,
    properties,
  }
}

/** Pure authoritative inspection. It does not mint identities or enroll transaction targets. */
export async function inspectSlidesAcceptanceAuthority(
  session: Session,
  request: SlidesAcceptanceAuthorityRequest,
): Promise<SlidesAcceptanceAuthoritySnapshot | null> {
  if (!session.documentInstanceId || !session.sessionInstanceId) return null
  if (request.affectedSlides.length > 50 || request.referenceSlides.length > 50)
    throw new TypeError('Acceptance inspection request is not bounded')
  const requested = [...new Set([...request.affectedSlides, ...request.referenceSlides])].sort(
    (a, b) => a - b,
  )
  if (
    requested.some(
      (page) => !Number.isSafeInteger(page) || page < 1 || page > session.opened.deck.slides.length,
    )
  )
    throw new TypeError('Acceptance inspection page is invalid')
  const revision = await fingerprintPresentation(session.opened)
  if (
    request.expectedDocumentToken !== session.documentInstanceId ||
    request.expectedSessionToken !== session.sessionInstanceId ||
    request.expectedRevision !== revision
  )
    return null
  const slides = []
  let inspectedElements = 0
  for (const page of requested) {
    const index = page - 1
    const slide = session.opened.deck.slides[index]!
    const elements = []
    for (const element of slide.elements) {
      if (++inspectedElements > 2_000)
        throw new TypeError('Acceptance inspection element bound exceeded')
      const fact = await elementFact(slide.durableId, element, false)
      if (fact) elements.push(fact)
    }
    for (const decoration of slide.decorations ?? []) {
      if (++inspectedElements > 2_000)
        throw new TypeError('Acceptance inspection element bound exceeded')
      const fact = await elementFact(slide.durableId, decoration, true)
      if (fact) elements.push(fact)
    }
    const slideDigest = await fingerprintSemanticValue({ slideId: slide.durableId })
    slides.push({
      number: page,
      slideToken: `slide:${slideDigest.slice('sha256:'.length)}`,
      ...(slide.background?.type === 'solid' ? { backgroundColor: slide.background.color } : {}),
      elements,
    })
  }
  return {
    documentToken: session.documentInstanceId,
    sessionToken: session.sessionInstanceId,
    revision,
    slides,
  }
}
