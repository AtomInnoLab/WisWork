import { MAX_PRESENTATION_SOURCE_TEXT_CHARS } from './presentation-source-limits'
import {
  parsePresentationClaimReview,
  type PresentationClaimReview,
} from './presentation-claim-review'
import {
  presentationSourceAttachmentId,
  assertDeckMatchesPresentationPlan,
  type PresentationPlan,
} from './presentation-plan'
import type { PresentationDeck } from './presentation'
import { array, choice, id, number, object, text, valid } from './presentation-schema'

type Outcome = PresentationClaimReview['outcome']
type SourceStatus = Outcome | 'unreviewed' | 'mixed'
type ClaimStatus = SourceStatus | 'partial' | 'no_sources'
type ReviewRef = Pick<
  PresentationClaimReview,
  'reviewId' | 'outcome' | 'evidenceDigest' | 'createdAt' | 'offset' | 'maxChars'
>
export interface PresentationPageReviews {
  version: 1
  projectId: string
  requestId: string
  pageId: string
  planRevision: number
  inputDigest: string
  planDigest: string
  claims: {
    claimId: string
    status: ClaimStatus
    sources: { sourceId: string; status: SourceStatus; reviews: ReviewRef[] }[]
  }[]
  checks: {
    support: 'historical_agent_reviews'
    sourceAuthority: 'not_verified'
    timeliness: 'not_verified'
    host: 'not_checked'
  }
}
const outcomes = ['supported', 'contradicted', 'insufficient_evidence'] as const
const digest = { ...text(64, 64), pattern: '^[a-f0-9]{64}$' }
const schema = object({
  version: { type: 'number', enum: [1] },
  projectId: { ...id, maxLength: 128 },
  requestId: { ...id, maxLength: 128 },
  pageId: id,
  planRevision: number(1, Number.MAX_SAFE_INTEGER),
  inputDigest: digest,
  planDigest: digest,
  claims: array(
    object({
      claimId: id,
      status: choice(...outcomes, 'unreviewed', 'mixed', 'partial', 'no_sources'),
      sources: array(
        object({
          sourceId: id,
          status: choice(...outcomes, 'unreviewed', 'mixed'),
          reviews: array(
            object({
              reviewId: { ...id, maxLength: 128 },
              outcome: choice(...outcomes),
              evidenceDigest: digest,
              createdAt: text(24, 24),
              offset: number(0, MAX_PRESENTATION_SOURCE_TEXT_CHARS),
              maxChars: number(1, 8000),
            }),
            32,
          ),
        }),
        3,
      ),
    }),
    32,
  ),
  checks: object({
    support: choice('historical_agent_reviews'),
    sourceAuthority: choice('not_verified'),
    timeliness: choice('not_verified'),
    host: choice('not_checked'),
  }),
})
function sourceStatus(reviews: ReviewRef[]): SourceStatus {
  const values = new Set(reviews.map((review) => review.outcome))
  return values.size === 0 ? 'unreviewed' : values.size === 1 ? reviews[0]!.outcome : 'mixed'
}
function claimStatus(sources: PresentationPageReviews['claims'][number]['sources']): ClaimStatus {
  if (!sources.length) return 'no_sources'
  const values = new Set(sources.map((source) => source.status))
  if (values.size === 1 && values.has('unreviewed')) return 'unreviewed'
  const reviewed = [...values].filter((value) => value !== 'unreviewed')
  if (values.has('mixed') || reviewed.length > 1) return 'mixed'
  if (values.has('unreviewed')) return 'partial'
  return reviewed[0]!
}
export function parsePresentationPageReviews(value: unknown): PresentationPageReviews {
  const reject = (): never => {
    throw new Error('presentation_page_reviews_invalid:schema')
  }
  if (!valid(value, schema)) reject()
  const report = value as PresentationPageReviews
  if (
    !Number.isSafeInteger(report.planRevision) ||
    new TextEncoder().encode(JSON.stringify(report)).byteLength > 64 * 1024
  )
    reject()
  const claims = new Set<string>(),
    reviews = new Set<string>()
  for (const claim of report.claims) {
    if (claims.has(claim.claimId) || claim.status !== claimStatus(claim.sources)) reject()
    claims.add(claim.claimId)
    const sources = new Set<string>()
    for (const source of claim.sources) {
      if (sources.has(source.sourceId) || source.status !== sourceStatus(source.reviews)) reject()
      sources.add(source.sourceId)
      for (const review of source.reviews) {
        if (
          reviews.has(review.reviewId) ||
          !Number.isSafeInteger(review.offset) ||
          !Number.isSafeInteger(review.maxChars) ||
          !Number.isFinite(Date.parse(review.createdAt)) ||
          new Date(review.createdAt).toISOString() !== review.createdAt
        )
          reject()
        reviews.add(review.reviewId)
      }
    }
  }
  if (reviews.size > 32) reject()
  return structuredClone(report)
}
export function summarizePresentationPageReviews(
  plan: PresentationPlan,
  deck: PresentationDeck,
  metadata: Pick<
    PresentationPageReviews,
    'projectId' | 'requestId' | 'pageId' | 'planRevision' | 'inputDigest' | 'planDigest'
  >,
  reviews: PresentationClaimReview[],
): PresentationPageReviews {
  assertDeckMatchesPresentationPlan(deck, plan)
  if (metadata.projectId !== plan.projectId || metadata.projectId !== deck.id)
    throw new Error('invalid_state')
  const page = plan.slides.find((page) => page.id === metadata.pageId)
  if (!page) throw new Error('not_found')
  const selected: PresentationClaimReview[] = []
  for (const value of reviews) {
    const review = parsePresentationClaimReview(value)
    if (
      review.projectId !== metadata.projectId ||
      review.requestId !== metadata.requestId ||
      review.planRevision !== metadata.planRevision ||
      review.inputDigest !== metadata.inputDigest ||
      review.planDigest !== metadata.planDigest
    )
      throw new Error('invalid_state')
    if (review.pageId !== page.id) continue
    const claim = plan.claims.find((claim) => claim.id === review.claimId)
    const source = plan.sources.find((source) => source.id === review.sourceId)
    if (
      !claim ||
      !page.claimIds.includes(claim.id) ||
      !claim.sourceIds.includes(review.sourceId) ||
      !source ||
      presentationSourceAttachmentId(source) !== review.attachmentId
    )
      throw new Error('invalid_state')
    selected.push(review)
  }
  const claims = page.claimIds.map((claimId) => {
    const claim = plan.claims.find((claim) => claim.id === claimId)!
    const sources = claim.sourceIds.map((sourceId) => {
      const refs = selected
        .filter((review) => review.claimId === claimId && review.sourceId === sourceId)
        .map(({ reviewId, outcome, evidenceDigest, createdAt, offset, maxChars }) => ({
          reviewId,
          outcome,
          evidenceDigest,
          createdAt,
          offset,
          maxChars,
        }))
      return { sourceId, status: sourceStatus(refs), reviews: refs }
    })
    return { claimId, status: claimStatus(sources), sources }
  })
  return parsePresentationPageReviews({
    version: 1,
    ...metadata,
    claims,
    checks: {
      support: 'historical_agent_reviews',
      sourceAuthority: 'not_verified',
      timeliness: 'not_verified',
      host: 'not_checked',
    },
  })
}
