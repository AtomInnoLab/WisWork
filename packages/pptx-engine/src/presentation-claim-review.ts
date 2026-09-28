import { MAX_PRESENTATION_SOURCE_TEXT_CHARS } from './presentation-source-limits'
import { parsePresentationClaimEvidence } from './presentation-claim-evidence'
import { choice, id, number, object, text, valid } from './presentation-schema'

export interface PresentationClaimReview {
  version: 1
  projectId: string
  requestId: string
  reviewId: string
  planRevision: number
  inputDigest: string
  planDigest: string
  pageId: string
  claimId: string
  sourceId: string
  attachmentId: string
  offset: number
  maxChars: number
  evidenceDigest: string
  outcome: 'supported' | 'contradicted' | 'insufficient_evidence'
  notes: string
  reviewer: 'agent'
  createdAt: string
  checks: {
    support: 'agent_reviewed'
    sourceAuthority: 'not_verified'
    timeliness: 'not_verified'
    host: 'not_checked'
  }
}
const digest = { ...text(64, 64), pattern: '^[a-f0-9]{64}$' }
const schema = object({
  version: { type: 'number', enum: [1] },
  projectId: { ...id, maxLength: 128 },
  requestId: { ...id, maxLength: 128 },
  reviewId: { ...id, maxLength: 128 },
  planRevision: number(1, Number.MAX_SAFE_INTEGER),
  inputDigest: digest,
  planDigest: digest,
  pageId: id,
  claimId: id,
  sourceId: id,
  attachmentId: digest,
  offset: number(0, MAX_PRESENTATION_SOURCE_TEXT_CHARS),
  maxChars: number(1, 8000),
  evidenceDigest: digest,
  outcome: choice('supported', 'contradicted', 'insufficient_evidence'),
  notes: text(2000, 1),
  reviewer: choice('agent'),
  createdAt: text(24, 24),
  checks: object({
    support: choice('agent_reviewed'),
    sourceAuthority: choice('not_verified'),
    timeliness: choice('not_verified'),
    host: choice('not_checked'),
  }),
})
/** A saved agent judgment does not verify source authority, timeliness, or host output. */
export function parsePresentationClaimReview(value: unknown): PresentationClaimReview {
  if (!valid(value, schema)) throw new Error('presentation_claim_review_invalid:schema')
  const review = value as PresentationClaimReview
  if (
    ![review.planRevision, review.offset, review.maxChars].every(Number.isSafeInteger) ||
    !review.notes.trim() ||
    !Number.isFinite(Date.parse(review.createdAt)) ||
    new Date(review.createdAt).toISOString() !== review.createdAt
  )
    throw new Error('presentation_claim_review_invalid:schema')
  return structuredClone(review)
}
/** Stable JSON of every validated evidence field, including the literal original text window. */
export function presentationClaimEvidenceContent(value: unknown): string {
  const canonical = (value: unknown): unknown =>
    value && typeof value === 'object'
      ? Object.fromEntries(
          Object.entries(value)
            .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
            .map(([key, child]) => [key, canonical(child)]),
        )
      : value
  return JSON.stringify(canonical(parsePresentationClaimEvidence(value)))
}
