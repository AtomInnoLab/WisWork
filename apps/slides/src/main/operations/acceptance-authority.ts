import { fingerprintPresentation, type SlideElement } from '@wiswork/pptx-engine'
import { fingerprintSemanticValue } from '@wiswork/presentation-ops'
import { EMU_PER_PX_96 } from '@wiswork/pptx-render'
import { randomUUID } from 'node:crypto'
import type {
  SlidesAcceptanceAuthorityLease,
  SlidesAcceptanceAuthorityRequest,
  SlidesAcceptanceAuthoritySnapshot,
  SlidesAcceptanceTextProofRequest,
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

const leases = new WeakMap<Session, SlidesAcceptanceAuthorityLease>()
const textProofSecret = randomUUID()
const normalizeText = (text: string): string =>
  text.replaceAll('\r\n', '\n').replaceAll('\r', '\n').normalize('NFC')

export async function inspectSlidesAcceptanceLease(
  session: Session,
): Promise<SlidesAcceptanceAuthorityLease | null> {
  if (!session.documentInstanceId || !session.sessionInstanceId) return null
  const lease = {
    documentToken: session.documentInstanceId,
    sessionToken: session.sessionInstanceId,
    revision: await fingerprintPresentation(session.opened),
    slideCount: session.opened.deck.slides.length,
    leaseToken: `lease:${randomUUID().replaceAll('-', '')}`,
  }
  leases.set(session, lease)
  return lease
}

async function elementFact(
  slideId: string,
  element: SlideElement,
  locked: boolean,
  leaseToken: string,
  textChecks: ReadonlyMap<string, { checkId: string; expectedText: string }>,
) {
  if (!element.creationId) return undefined
  const role = roleOf(element)
  const targetDigest = await fingerprintSemanticValue({ slideId, elementId: element.creationId })
  const targetToken = `target:${targetDigest.slice('sha256:'.length)}`
  const properties: Record<string, string | number | boolean | null> = {
    x: element.transform.offset.x / EMU_PER_PX_96,
    y: element.transform.offset.y / EMU_PER_PX_96,
    width: element.transform.offset.cx / EMU_PER_PX_96,
    height: element.transform.offset.cy / EMU_PER_PX_96,
  }
  let textMatch:
    { checkId: string; targetToken: string; matches: boolean; proof: string } | undefined
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
    const textCheck = textChecks.get(targetToken)
    if (textCheck) {
      const text = element.text.paragraphs
        .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
        .join('\n')
      const matches = normalizeText(text) === normalizeText(textCheck.expectedText)
      const proof = await fingerprintSemanticValue({
        secret: textProofSecret,
        leaseToken,
        checkId: textCheck.checkId,
        targetToken,
        expectedText: normalizeText(textCheck.expectedText),
        matches,
      })
      textMatch = { checkId: textCheck.checkId, targetToken, matches, proof }
    }
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
    fact: { targetToken, ...(role ? { role } : {}), locked, properties },
    ...(textMatch ? { textMatch } : {}),
  }
}

/** Pure authoritative inspection. It does not mint identities or enroll transaction targets. */
export async function inspectSlidesAcceptanceAuthority(
  session: Session,
  request: SlidesAcceptanceAuthorityRequest,
  provedLineage?: {
    baseRevision: string
    resultingRevision: string
    mutatedTargets: readonly string[]
  },
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
    (request.baseRevision !== undefined || request.mutationReceiptIds !== undefined) &&
    (!provedLineage ||
      provedLineage.baseRevision !== request.baseRevision ||
      provedLineage.resultingRevision !== revision)
  )
    return null
  const lease = leases.get(session)
  if (
    !lease ||
    request.leaseToken !== lease.leaseToken ||
    request.expectedDocumentToken !== session.documentInstanceId ||
    request.expectedSessionToken !== session.sessionInstanceId ||
    request.expectedRevision !== revision ||
    lease.documentToken !== session.documentInstanceId ||
    lease.sessionToken !== session.sessionInstanceId ||
    lease.revision !== revision
  )
    return null
  const textChecks = new Map(
    (request.textChecks ?? []).map((check) => [
      check.targetToken,
      { checkId: check.checkId, expectedText: check.expectedText },
    ]),
  )
  const textMatches: NonNullable<SlidesAcceptanceAuthoritySnapshot['textMatches']> = {}
  const requestedSources = new Map(
    (request.sourceTargets ?? []).map(({ slide, sourceId }) => [`${slide}:${sourceId}`, false]),
  )
  if (requestedSources.size !== (request.sourceTargets ?? []).length || requestedSources.size > 50)
    throw new TypeError('Acceptance source target request is invalid')
  const sourceTargetTokens: Record<string, string> = {}
  const slides = []
  let inspectedElements = 0
  for (const page of requested) {
    const index = page - 1
    const slide = session.opened.deck.slides[index]!
    const elements = []
    for (const element of slide.elements) {
      if (++inspectedElements > 2_000)
        throw new TypeError('Acceptance inspection element bound exceeded')
      const fact = await elementFact(slide.durableId, element, false, lease.leaseToken, textChecks)
      if (fact) {
        elements.push(fact.fact)
        const sourceKey = `${page}:${element.id}`
        if (requestedSources.has(sourceKey)) {
          requestedSources.set(sourceKey, true)
          sourceTargetTokens[sourceKey] = fact.fact.targetToken
        }
        if (fact.textMatch) textMatches[fact.textMatch.checkId] = fact.textMatch
      }
    }
    for (const decoration of slide.decorations ?? []) {
      if (++inspectedElements > 2_000)
        throw new TypeError('Acceptance inspection element bound exceeded')
      const fact = await elementFact(
        slide.durableId,
        decoration,
        true,
        lease.leaseToken,
        textChecks,
      )
      if (fact) {
        elements.push(fact.fact)
        if (fact.textMatch) textMatches[fact.textMatch.checkId] = fact.textMatch
      }
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
    ...(provedLineage
      ? {
          baseRevision: provedLineage.baseRevision,
          mutatedTargetTokens: [...provedLineage.mutatedTargets],
        }
      : {}),
    leaseToken: lease.leaseToken,
    ...(Object.keys(textMatches).length ? { textMatches } : {}),
    ...(Object.keys(sourceTargetTokens).length ? { sourceTargetTokens } : {}),
    slides,
  }
}

export async function verifySlidesAcceptanceTextProof(
  session: Session,
  request: SlidesAcceptanceTextProofRequest,
): Promise<boolean> {
  const snapshot = await inspectSlidesAcceptanceAuthority(session, {
    affectedSlides: [request.slide],
    referenceSlides: [],
    expectedDocumentToken: request.expectedDocumentToken,
    expectedSessionToken: request.expectedSessionToken,
    expectedRevision: request.expectedRevision,
    leaseToken: request.leaseToken,
    textChecks: [
      {
        checkId: request.checkId,
        targetToken: request.targetToken,
        expectedText: request.expectedText,
      },
    ],
  })
  const result = snapshot?.textMatches?.[request.checkId]
  return (
    !!result &&
    result.targetToken === request.targetToken &&
    result.matches === request.matches &&
    result.proof === request.proof
  )
}
