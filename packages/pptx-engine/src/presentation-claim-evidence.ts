import { presentationSourceAttachmentId } from './presentation-plan'
import { choice, id, number, object, text, valid } from './presentation-schema'

export interface PresentationClaimEvidence {
  version: 1
  projectId: string
  requestId: string
  planRevision: number
  inputDigest: string
  planDigest: string
  pageId: string
  claimId: string
  statement: string
  source: {
    id: string
    uri: string
    snapshotAttachmentId?: string
    excerpt: string
    locator?: string
  }
  attachment: {
    id: string
    name: string
    offset: number
    totalChars: number
    text: string
    offsetUnit: 'utf16_code_unit'
  }
  excerptMatch:
    { status: 'found'; offset: number } | { status: 'not_found_in_window' | 'empty_excerpt' }
  checks: {
    support: 'not_verified'
    sourceAuthority: 'not_verified'
    timeliness: 'not_verified'
    host: 'not_checked'
  }
}

export function matchPresentationClaimExcerpt(
  excerpt: string,
  window: string,
  offset: number,
): PresentationClaimEvidence['excerptMatch'] {
  if (!excerpt.trim()) return { status: 'empty_excerpt' }
  const index = window.indexOf(excerpt)
  return index < 0 ? { status: 'not_found_in_window' } : { status: 'found', offset: offset + index }
}
const digest = { ...text(64, 64), pattern: '^[a-f0-9]{64}$' }
const schema = object({
  version: { type: 'number', enum: [1] },
  projectId: { ...id, maxLength: 128 },
  requestId: { ...id, maxLength: 128 },
  planRevision: number(1, Number.MAX_SAFE_INTEGER),
  inputDigest: digest,
  planDigest: digest,
  pageId: id,
  claimId: id,
  statement: text(12000, 1),
  source: object(
    {
      id,
      uri: text(500, 1),
      snapshotAttachmentId: digest,
      excerpt: text(12000),
      locator: text(200),
    },
    ['id', 'uri', 'excerpt'],
  ),
  attachment: object({
    id: digest,
    name: text(180, 1),
    offset: number(0, 1000000),
    totalChars: number(0, 1000000),
    text: text(8000),
    offsetUnit: choice('utf16_code_unit'),
  }),
  excerptMatch: {
    anyOf: [
      object({ status: choice('found'), offset: number(0, 1000000) }),
      object({ status: choice('not_found_in_window', 'empty_excerpt') }),
    ],
  },
  checks: object({
    support: choice('not_verified'),
    sourceAuthority: choice('not_verified'),
    timeliness: choice('not_verified'),
    host: choice('not_checked'),
  }),
})
/** Validates literal evidence only; matching does not verify truth or source authority. */
export function parsePresentationClaimEvidence(value: unknown): PresentationClaimEvidence {
  const reject = (): never => {
    throw new Error('presentation_claim_evidence_invalid:schema')
  }
  // Parsed attachment text is not XML: preserve form feeds and other original code units.
  const rawText = (value as Partial<PresentationClaimEvidence> | null)?.attachment?.text
  if (typeof rawText !== 'string' || rawText.length > 8000) reject()
  if (
    !valid(
      {
        ...(value as object),
        attachment: { ...(value as PresentationClaimEvidence).attachment, text: '' },
      },
      schema,
    )
  )
    reject()
  const report = value as PresentationClaimEvidence
  const { attachment, source } = report
  let attachmentId: string | undefined
  try {
    attachmentId = presentationSourceAttachmentId(source)
  } catch {
    reject()
  }
  if (
    new TextEncoder().encode(JSON.stringify(report)).byteLength > 256 * 1024 ||
    ![report.planRevision, attachment.offset, attachment.totalChars].every(Number.isSafeInteger) ||
    attachmentId !== attachment.id ||
    attachment.offset + attachment.text.length > attachment.totalChars
  )
    reject()
  const expected = matchPresentationClaimExcerpt(source.excerpt, attachment.text, attachment.offset)
  if (
    expected.status !== report.excerptMatch.status ||
    ('offset' in expected &&
      (!('offset' in report.excerptMatch) || expected.offset !== report.excerptMatch.offset))
  )
    reject()
  return structuredClone(report)
}
