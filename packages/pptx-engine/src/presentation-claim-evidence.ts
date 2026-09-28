import { presentationSourceAttachmentId } from './presentation-plan'
import { array, choice, id, number, object, text, valid } from './presentation-schema'

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
    locatorSpans?: { locator: string; start: number; end: number }[]
  }
  excerptMatch:
    | { status: 'found'; offset: number; locator?: string }
    | { status: 'not_found_in_window' | 'empty_excerpt' }
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
  locatorSpans?: { locator: string; start: number; end: number }[],
  preferredLocator?: string,
): PresentationClaimEvidence['excerptMatch'] {
  if (!excerpt.trim()) return { status: 'empty_excerpt' }
  let first: { offset: number; locator?: string } | undefined
  let located: { offset: number; locator: string } | undefined
  for (
    let index = window.indexOf(excerpt);
    index >= 0;
    index = window.indexOf(excerpt, index + 1)
  ) {
    const absolute = offset + index
    const section = locatorSpans?.find(
      (span) => span.start <= absolute && absolute + excerpt.length <= span.end,
    )
    const match = { offset: absolute, ...(section ? { locator: section.locator } : {}) }
    if (!first) first = match
    if (section && !located) located = { offset: absolute, locator: section.locator }
    if (preferredLocator !== undefined && match.locator === preferredLocator)
      return { status: 'found', ...match }
  }
  return located || first
    ? { status: 'found', ...(located ?? first!) }
    : { status: 'not_found_in_window' }
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
  attachment: object(
    {
      id: digest,
      name: text(180, 1),
      offset: number(0, 1000000),
      totalChars: number(0, 1000000),
      text: text(8000),
      offsetUnit: choice('utf16_code_unit'),
      locatorSpans: array(
        object({ locator: text(32, 1), start: number(0, 1000000), end: number(0, 1000000) }),
        4096,
      ),
    },
    ['id', 'name', 'offset', 'totalChars', 'text', 'offsetUnit'],
  ),
  excerptMatch: {
    anyOf: [
      object({ status: choice('found'), offset: number(0, 1000000), locator: text(32, 1) }, [
        'status',
        'offset',
      ]),
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
  const locatorUnit = /\.pdf$/i.test(attachment.name)
    ? '页'
    : /\.(docx|html|htm)$/i.test(attachment.name)
      ? '段'
      : undefined
  if (
    (attachment.locatorSpans !== undefined && !locatorUnit) ||
    attachment.locatorSpans?.some(
      (span, index) =>
        !new RegExp(`^第 [1-9]\\d{0,5} ${locatorUnit}$`).test(span.locator) ||
        !Number.isSafeInteger(span.start) ||
        !Number.isSafeInteger(span.end) ||
        span.start >= span.end ||
        span.end > attachment.totalChars ||
        span.end <= attachment.offset ||
        span.start >= attachment.offset + attachment.text.length ||
        (index > 0 && span.start <= attachment.locatorSpans![index - 1]!.end),
    )
  )
    reject()
  const preferred = /^第\s*([1-9]\d{0,5})\s*(页|段)$/.exec(source.locator?.trim() ?? '')
  const expected = matchPresentationClaimExcerpt(
    source.excerpt,
    attachment.text,
    attachment.offset,
    attachment.locatorSpans,
    preferred ? `第 ${Number(preferred[1])} ${preferred[2]}` : undefined,
  )
  if (
    expected.status !== report.excerptMatch.status ||
    ('offset' in expected &&
      (!('offset' in report.excerptMatch) ||
        expected.offset !== report.excerptMatch.offset ||
        expected.locator !== report.excerptMatch.locator))
  )
    reject()
  return structuredClone(report)
}
