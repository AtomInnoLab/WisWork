import { MAX_PRESENTATION_SOURCE_TEXT_CHARS } from './presentation-source-limits'
import { parsePresentationProfessionalContext } from '@wiswork/project-store/presentation-professional-context'
import {
  presentationSourceAttachmentId,
  PRESENTATION_PLAN_SCHEMA,
  type PresentationPlan,
} from './presentation-plan'
import {
  parsePresentationResearchRecord,
  type PresentationResearchRecord,
} from '@wiswork/project-store/presentation-research'
import { canonicalPresentationValue as canonical } from '@wiswork/project-store/presentation-canonical'
import {
  presentationResearchClaimBindingFindings,
  type PresentationResearchBindingFinding,
} from './presentation-research-binding'
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
  documentId?: string
  claim?: PresentationPlan['claims'][number]
  research?: {
    binding: NonNullable<PresentationPlan['research']>
    record: PresentationResearchRecord
    findings: PresentationResearchBindingFinding[]
  }
  source: {
    id: string
    uri: string
    snapshotAttachmentId?: string
    excerpt: string
    locator?: string
    asOf?: string
  }
  attachment: {
    id: string
    name: string
    offset: number
    totalChars: number
    text: string
    offsetUnit: 'utf16_code_unit'
    locatorSpans?: {
      locator: string
      start: number
      end: number
      imageBacked?: true
      invisibleTextLayer?: true
    }[]
    provenance?:
      { binding: 'fetched_url_matched'; retrievedAt: number } | { binding: 'user_supplied' }
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
      offset: number(0, MAX_PRESENTATION_SOURCE_TEXT_CHARS),
      totalChars: number(0, MAX_PRESENTATION_SOURCE_TEXT_CHARS),
      text: text(8000),
      offsetUnit: choice('utf16_code_unit'),
      locatorSpans: array(
        object(
          {
            locator: text(32, 1),
            start: number(0, MAX_PRESENTATION_SOURCE_TEXT_CHARS),
            end: number(0, MAX_PRESENTATION_SOURCE_TEXT_CHARS),
            imageBacked: { type: 'boolean', enum: [true] },
            invisibleTextLayer: { type: 'boolean', enum: [true] },
          },
          ['locator', 'start', 'end'],
        ),
        4096,
      ),
      provenance: {
        anyOf: [
          object({
            binding: choice('fetched_url_matched'),
            retrievedAt: number(1, Number.MAX_SAFE_INTEGER),
          }),
          object({ binding: choice('user_supplied') }),
        ],
      },
    },
    ['id', 'name', 'offset', 'totalChars', 'text', 'offsetUnit'],
  ),
  excerptMatch: {
    anyOf: [
      object(
        {
          status: choice('found'),
          offset: number(0, MAX_PRESENTATION_SOURCE_TEXT_CHARS),
          locator: text(32, 1),
        },
        ['status', 'offset'],
      ),
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
export const PRESENTATION_CLAIM_EVIDENCE_LIMIT = 256 * 1024
export const PRESENTATION_BOUND_CLAIM_EVIDENCE_LIMIT = 512 * 1024
export function presentationClaimEvidenceLimit(value: unknown): number {
  return (value as Partial<PresentationClaimEvidence> | null)?.research !== undefined
    ? PRESENTATION_BOUND_CLAIM_EVIDENCE_LIMIT
    : PRESENTATION_CLAIM_EVIDENCE_LIMIT
}
function validateResearchContext(evidence: PresentationClaimEvidence) {
  function fail(): never {
    throw new Error('presentation_claim_evidence_invalid:research')
  }
  const context = evidence.research,
    claim = evidence.claim
  if (
    !context ||
    !claim ||
    typeof evidence.documentId !== 'string' ||
    !evidence.documentId.trim() ||
    evidence.documentId.length > 4096 ||
    !valid(claim, PRESENTATION_PLAN_SCHEMA.properties!.claims!.items!) ||
    !valid(context.binding, PRESENTATION_PLAN_SCHEMA.properties!.research!) ||
    Object.keys(context).sort().join(',') !== 'binding,findings,record'
  )
    fail()
  const record = parsePresentationResearchRecord(context.record),
    binding = context.binding
  if (Object.hasOwn(claim, 'professionalContext'))
    parsePresentationProfessionalContext(claim.professionalContext)
  if (
    record.state !== 'completed' ||
    record.documentId !== evidence.documentId ||
    record.projectId !== evidence.projectId ||
    record.id !== binding.ledgerId ||
    record.sequence !== binding.sequence ||
    record.draftDigest !== binding.draftDigest ||
    !Number.isSafeInteger(binding.sequence) ||
    claim.id !== evidence.claimId ||
    claim.statement !== evidence.statement ||
    !claim.sourceIds.includes(evidence.source.id) ||
    new Set(claim.sourceIds).size !== claim.sourceIds.length ||
    new Set(binding.sources.map((m) => m.sourceId)).size !== binding.sources.length ||
    new Set(binding.claims.map((m) => m.claimId)).size !== binding.claims.length
  )
    fail()
  if (['fact', 'quote', 'calculation'].includes(claim.type) && !claim.sourceIds.length) fail()
  if (claim.type === 'calculation' && !claim.calculation) fail()
  const reproduction = claim.calculation?.reproduction
  if (
    reproduction &&
    (claim.type !== 'calculation' ||
      reproduction.bindings.length !== claim.calculation!.inputs.length ||
      new Set(reproduction.bindings.map((b) => b.name)).size !== reproduction.bindings.length ||
      new Set(reproduction.bindings.map((b) => b.inputIndex)).size !==
        reproduction.bindings.length ||
      reproduction.bindings.some(
        (b) =>
          !Number.isInteger(b.inputIndex) ||
          b.inputIndex >= claim.calculation!.inputs.length ||
          ['prototype', 'constructor', '__proto__'].includes(b.name) ||
          !claim.sourceIds.includes(b.sourceId),
      ))
  )
    fail()
  const sources = new Map(binding.sources.map((m) => [m.sourceId, m.researchSourceId]))
  if (
    binding.sources.some((m) => !record.draft.sources.some((s) => s.id === m.researchSourceId)) ||
    binding.claims.some((m) => !record.draft.facts.some((f) => f.claimId === m.researchClaimId))
  )
    fail()
  const pick = (v: object, keys: string[]) =>
    Object.fromEntries(keys.map((k) => [k, (v as Record<string, unknown>)[k]]))
  const sourceMapping = sources.get(evidence.source.id)
  if (sourceMapping) {
    const original = record.draft.sources.find((s) => s.id === sourceMapping)!
    if (
      canonical(
        pick(evidence.source, ['uri', 'snapshotAttachmentId', 'excerpt', 'locator', 'asOf']),
      ) !== canonical(pick(original, ['uri', 'snapshotAttachmentId', 'excerpt', 'locator', 'asOf']))
    )
      fail()
  }
  if (evidence.source.asOf !== undefined && !valid(evidence.source.asOf, text(100, 1))) fail()
  const mapping = binding.claims.find((m) => m.claimId === claim.id)
  if (mapping) {
    const fact = record.draft.facts.find((f) => f.claimId === mapping.researchClaimId)!
    if (
      canonical(
        pick(claim, ['statement', 'type', 'asOf', 'jurisdiction', 'professionalContext']),
      ) !==
        canonical(
          pick(fact, ['statement', 'type', 'asOf', 'jurisdiction', 'professionalContext']),
        ) ||
      canonical(
        claim.calculation
          ? pick(claim.calculation, ['formula', 'inputs', 'unit', 'currency'])
          : undefined,
      ) !==
        canonical(
          fact.calculation
            ? pick(fact.calculation, ['formula', 'inputs', 'unit', 'currency'])
            : undefined,
        ) ||
      claim.sourceIds.some((id) => !sources.has(id) || !fact.sourceRefs.includes(sources.get(id)!))
    )
      fail()
  }
  const findings = presentationResearchClaimBindingFindings(claim, binding, record)
  if (canonical(context.findings) !== canonical(findings)) fail()
}
/** Validates literal evidence only; matching does not verify truth or source authority. */
export function parsePresentationClaimEvidence(value: unknown): PresentationClaimEvidence {
  const reject = (): never => {
    throw new Error('presentation_claim_evidence_invalid:schema')
  }
  // Parsed attachment text is not XML: preserve form feeds and other original code units.
  const rawText = (value as Partial<PresentationClaimEvidence> | null)?.attachment?.text
  if (typeof rawText !== 'string' || rawText.length > 8000) reject()
  const base = {
    ...(value as PresentationClaimEvidence),
    attachment: { ...(value as PresentationClaimEvidence).attachment, text: '' },
    source: { ...(value as PresentationClaimEvidence).source },
  }
  delete base.documentId
  delete base.claim
  delete base.research
  delete base.source.asOf
  if (!valid(base, schema)) reject()
  const report = value as PresentationClaimEvidence
  const { attachment, source } = report
  if (report.research !== undefined) validateResearchContext(report)
  else if (report.claim?.professionalContext !== undefined) {
    const claim = report.claim
    if (
      !valid(claim, PRESENTATION_PLAN_SCHEMA.properties!.claims!.items!) ||
      typeof report.documentId !== 'string' ||
      !report.documentId.trim() ||
      report.documentId.length > 4096 ||
      claim.id !== report.claimId ||
      claim.statement !== report.statement ||
      !claim.sourceIds.includes(source.id) ||
      new Set(claim.sourceIds).size !== claim.sourceIds.length ||
      (claim.type === 'calculation' && !claim.calculation) ||
      Object.hasOwn(report, 'research') ||
      (source.asOf !== undefined && !valid(source.asOf, text(100, 1)))
    )
      reject()
    parsePresentationProfessionalContext(claim.professionalContext)
    const reproduction = claim.calculation?.reproduction
    if (
      reproduction &&
      (claim.type !== 'calculation' ||
        reproduction.bindings.length !== claim.calculation!.inputs.length ||
        new Set(reproduction.bindings.map((binding) => binding.name)).size !==
          reproduction.bindings.length ||
        new Set(reproduction.bindings.map((binding) => binding.inputIndex)).size !==
          reproduction.bindings.length ||
        reproduction.bindings.some(
          (binding) =>
            !Number.isInteger(binding.inputIndex) ||
            binding.inputIndex >= claim.calculation!.inputs.length ||
            ['prototype', 'constructor', '__proto__'].includes(binding.name) ||
            !claim.sourceIds.includes(binding.sourceId),
        ))
    )
      reject()
  } else if (
    Object.hasOwn(report, 'research') ||
    Object.hasOwn(report, 'documentId') ||
    Object.hasOwn(report, 'claim') ||
    Object.hasOwn(source, 'asOf')
  )
    reject()
  let attachmentId: string | undefined
  try {
    attachmentId = presentationSourceAttachmentId(source)
  } catch {
    reject()
  }
  if (
    new TextEncoder().encode(JSON.stringify(report)).byteLength >
      presentationClaimEvidenceLimit(report) ||
    ![report.planRevision, attachment.offset, attachment.totalChars].every(Number.isSafeInteger) ||
    attachmentId !== attachment.id ||
    attachment.offset + attachment.text.length > attachment.totalChars ||
    (attachment.provenance?.binding === 'fetched_url_matched' &&
      (!Number.isSafeInteger(attachment.provenance.retrievedAt) ||
        !source.snapshotAttachmentId ||
        !/^https?:\/\//i.test(source.uri)))
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
        span.start > span.end ||
        (span.imageBacked === true && locatorUnit !== '页') ||
        (span.invisibleTextLayer === true && span.imageBacked !== true) ||
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
