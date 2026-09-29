import { MAX_PRESENTATION_SOURCE_TEXT_CHARS } from './presentation-source-limits'
import { canonicalPresentationValue } from '@wiswork/project-store/presentation-canonical'
import {
  parsePresentationResearchRecord,
  type PresentationResearchRecord,
} from '@wiswork/project-store/presentation-research'
import {
  presentationResearchBindingFindings,
  type PresentationResearchBindingFinding,
} from './presentation-research-binding'
import {
  parsePresentationIssueLedger,
  type PresentationIssueLedger,
} from '@wiswork/project-store/presentation-issue'
import {
  parsePresentationPlan,
  presentationSourceAttachmentId,
  assertDeckMatchesPresentationPlan,
  type PresentationPlan,
} from './presentation-plan'
import { parsePresentationDeck, type PresentationDeck } from './presentation'
import { checkPresentationPageContent } from './presentation-content-check'
import { sourceAsOfFinding } from './presentation-source-time'
import {
  parsePresentationClaimReview,
  type PresentationClaimReview,
} from './presentation-claim-review'
import { summarizePresentationPageReviews } from './presentation-page-reviews'
import {
  reproducePresentationCalculation,
  type CalculationResult,
} from './presentation-calculation'
export type { CalculationResult } from './presentation-calculation'

export interface DeliveryIssue {
  id: string
  code: string
  claimId: string
  sourceId?: string
  research?: {
    ledgerId: string
    sequence: number
    draftDigest: string
    researchClaimId?: string
    relatedClaimIds: string[]
    sourceIds: string[]
  }
  digest: string
  category: 'needs_human' | 'unverifiable'
  disposition: { state: 'open' | 'deferred' | 'explained'; stale: boolean; actionId?: string }
}
export interface PresentationSourceAudit {
  sourceId: string
  attachmentId: string
  status:
    | 'found'
    | 'not_found'
    | 'empty_excerpt'
    | 'not_ready'
    | 'unsupported'
    | 'missing'
    | 'source_mismatch'
  offset?: number
  locator?: string
}
export interface PresentationDeliveryReport {
  version: 1
  projectId: string
  documentId: string
  requestId: string
  planRevision: number
  inputDigest: string
  planDigest: string
  plan: PresentationPlan
  research?: { record: PresentationResearchRecord; findings: PresentationResearchBindingFinding[] }
  sourceAudit?: PresentationSourceAudit[]
  reviews: PresentationClaimReview[]
  issueLedger: PresentationIssueLedger
  pages: {
    pageId: string
    title: string
    productionState: 'pending' | 'building' | 'compiled' | 'failed'
    calculations: CalculationResult[]
    issues: DeliveryIssue[]
  }[]
  checks: {
    scope: 'frozen_production'
    content: 'needs_review'
    sourceAuthority: 'not_verified'
    timeliness: 'not_verified'
    host: 'not_checked'
    roundTrip: 'not_run'
  }
}
export interface PresentationDeliveryReportInput {
  plan: PresentationPlan
  researchRecord?: PresentationResearchRecord
  sourceAudit?: PresentationSourceAudit[]
  deck: PresentationDeck
  metadata: Pick<
    PresentationDeliveryReport,
    'projectId' | 'documentId' | 'requestId' | 'planRevision' | 'inputDigest' | 'planDigest'
  >
  pageStates: {
    pageId: string
    state: PresentationDeliveryReport['pages'][number]['productionState']
  }[]
  reviews: PresentationClaimReview[]
  issueLedger: PresentationIssueLedger
}
const checks: PresentationDeliveryReport['checks'] = {
  scope: 'frozen_production',
  content: 'needs_review',
  sourceAuthority: 'not_verified',
  timeliness: 'not_verified',
  host: 'not_checked',
  roundTrip: 'not_run',
}
const canonical = (value: unknown): string => JSON.stringify(sort(value))
function sort(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(sort)
  if (value && typeof value === 'object')
    return Object.fromEntries(
      Object.entries(value)
        .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
        .map(([key, child]) => [key, sort(child)]),
    )
  return value
}
function invalid(): never {
  throw new Error('presentation_delivery_report_invalid:schema')
}
function exact(value: unknown, keys: string[]): asserts value is Record<string, unknown> {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).some((key) => !keys.includes(key))
  )
    invalid()
}
const digestValid = (value: unknown): boolean =>
  typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
const normalize = (value: string): string => value.replace(/\s+/g, ' ').trim()
type Seed = Pick<DeliveryIssue, 'id' | 'code' | 'claimId' | 'sourceId' | 'category' | 'research'>
function seeds(
  report: PresentationDeliveryReport,
  pageIndex: number,
  missing: Set<string>,
): Seed[] {
  const page = report.plan.slides[pageIndex]!
  const result: Seed[] = []
  for (const [claimIndex, claimId] of page.claimIds.entries()) {
    const claim = report.plan.claims.find((item) => item.id === claimId)!
    const add = (code: string, sourceId?: string): void => {
      const unverifiable = [
        'claim_no_sources',
        'source_excerpt_missing',
        'source_excerpt_not_in_attachment',
        'source_attachment_missing',
        'source_attachment_not_ready',
        'source_attachment_unsupported',
        'source_url_mismatch',
        'source_locator_mismatch',
        'source_original_not_frozen',
        'source_locator_missing',
        'source_review_missing',
        'source_review_insufficient',
        'source_authority_review_missing',
        'source_authority_review_uncertain',
        'source_authority_review_insufficient',
        'source_timeliness_review_missing',
        'source_timeliness_review_uncertain',
        'source_jurisdiction_review_missing',
        'source_jurisdiction_review_uncertain',
        'calculation_not_reproduced',
        'calculation_unsupported',
        'calculation_invalid_arithmetic',
      ].includes(code)
      result.push({
        id: `p${pageIndex}_c${claimIndex}_s${sourceId === undefined ? 'x' : claim.sourceIds.indexOf(sourceId)}_${code}`,
        code,
        claimId,
        ...(sourceId === undefined ? {} : { sourceId }),
        category: unverifiable ? 'unverifiable' : 'needs_human',
      })
    }
    if (missing.has(claimId)) add('claim_text_not_found')
    if (!claim.sourceIds.length) add('claim_no_sources')
    for (const sourceId of claim.sourceIds) {
      const source = report.plan.sources.find((item) => item.id === sourceId)!
      const audit = report.sourceAudit?.find((item) => item.sourceId === sourceId)
      if (!presentationSourceAttachmentId(source)) add('source_original_not_frozen', sourceId)
      if (audit?.status === 'not_found') add('source_excerpt_not_in_attachment', sourceId)
      if (audit?.status === 'missing') add('source_attachment_missing', sourceId)
      if (audit?.status === 'not_ready') add('source_attachment_not_ready', sourceId)
      if (audit?.status === 'unsupported') add('source_attachment_unsupported', sourceId)
      if (audit?.status === 'source_mismatch') add('source_url_mismatch', sourceId)
      const plannedLocator = /^第\s*(\d+)\s*(页|段)$/.exec(source.locator?.trim() ?? '')
      const observedLocator = audit?.locator
        ? /^第\s*(\d+)\s*(页|段)$/.exec(audit.locator)
        : undefined
      if (
        plannedLocator &&
        observedLocator &&
        (plannedLocator[1] !== observedLocator[1] || plannedLocator[2] !== observedLocator[2])
      )
        add('source_locator_mismatch', sourceId)
      const excerpt = normalize(source.excerpt)
      if (!excerpt) add('source_excerpt_missing', sourceId)
      if (!source.locator?.trim()) add('source_locator_missing', sourceId)
      const asOfFinding = sourceAsOfFinding(claim.asOf, source.asOf)
      if (asOfFinding) add(asOfFinding, sourceId)
      if (claim.type === 'quote' && excerpt && !excerpt.includes(normalize(claim.statement)))
        add('quote_not_in_excerpt', sourceId)
      const relevant = report.reviews.filter(
        (review) =>
          review.pageId === page.id && review.claimId === claimId && review.sourceId === sourceId,
      )
      const outcomes = new Set(relevant.map((review) => review.outcome))
      if (!outcomes.size) add('source_review_missing', sourceId)
      else if (outcomes.size > 1) add('source_review_mixed', sourceId)
      else if (outcomes.has('contradicted')) add('source_review_contradicted', sourceId)
      else if (outcomes.has('insufficient_evidence')) add('source_review_insufficient', sourceId)
      if (report.plan.research || report.reviews.some((review) => review.sourceAssessment)) {
        const assessments = relevant.flatMap((review) =>
          review.sourceAssessment ? [review.sourceAssessment] : [],
        )
        const authority = new Set(
          assessments.map((item) => canonical([item.authority.outcome, item.authority.sourceTier])),
        )
        if (!assessments.length) add('source_authority_review_missing', sourceId)
        else if (authority.size > 1) add('source_authority_review_mixed', sourceId)
        else if (assessments[0]!.authority.outcome === 'uncertain')
          add('source_authority_review_uncertain', sourceId)
        else if (assessments[0]!.authority.outcome === 'insufficient_authority')
          add('source_authority_review_insufficient', sourceId)
        const timeliness = new Set(
          assessments.map((item) =>
            canonical([
              item.timeliness.outcome,
              item.timeliness.referenceDate,
              item.timeliness.claimAsOf,
              item.timeliness.sourceAsOf,
            ]),
          ),
        )
        if (!assessments.length) add('source_timeliness_review_missing', sourceId)
        else if (timeliness.size > 1) add('source_timeliness_review_mixed', sourceId)
        else if (assessments[0]!.timeliness.outcome !== 'current_for_claim')
          add(`source_timeliness_review_${assessments[0]!.timeliness.outcome}`, sourceId)
        if (claim.jurisdiction !== undefined) {
          const jurisdictions = assessments.flatMap((item) =>
            item.jurisdiction ? [item.jurisdiction] : [],
          )
          const frames = new Set(
            assessments.map((item) =>
              canonical(
                item.jurisdiction
                  ? [item.jurisdiction.outcome, item.jurisdiction.claimJurisdiction]
                  : null,
              ),
            ),
          )
          if (!jurisdictions.length) add('source_jurisdiction_review_missing', sourceId)
          else if (frames.size > 1) add('source_jurisdiction_review_mixed', sourceId)
          else if (jurisdictions[0]!.outcome !== 'applicable')
            add(`source_jurisdiction_review_${jurisdictions[0]!.outcome}`, sourceId)
        }
      }
    }
    if (claim.type === 'calculation') {
      const calculation = reproducePresentationCalculation(claim)
      const code = {
        not_configured: 'calculation_not_reproduced',
        mismatch: 'calculation_mismatch',
        unsupported_expression: 'calculation_unsupported',
        invalid_arithmetic: 'calculation_invalid_arithmetic',
        reproduced: '',
      }[calculation.status]
      if (code) add(code)
    }
    if (report.research) {
      const { record, findings } = report.research
      const mapping = report.plan.research!.claims.find((item) => item.claimId === claimId)
      const original = record.draft.facts.find((item) => item.claimId === mapping?.researchClaimId)
      const researchIssue = (
        code: string,
        relatedClaimIds: string[],
        sourceIds: string[],
        category: DeliveryIssue['category'],
      ) => {
        result.push({
          id: `p${pageIndex}_c${claimIndex}_sx_${code}`,
          code,
          claimId,
          category,
          research: {
            ledgerId: record.id,
            sequence: record.sequence,
            draftDigest: record.draftDigest,
            ...(original ? { researchClaimId: original.claimId } : {}),
            relatedClaimIds: [...new Set(relatedClaimIds)].sort(),
            sourceIds: [...new Set(sourceIds)].sort(),
          },
        })
      }
      const codes = {
        unmapped_claim: ['research_unmapped_claim', 'unverifiable'],
        omitted_conflict_partner: ['research_conflict_partner_omitted', 'needs_human'],
        unselected_source_ref: ['research_source_reference_unselected', 'needs_human'],
        source_unavailable: ['research_source_unavailable', 'unverifiable'],
      } as const
      for (const [findingCode, [code, category]] of Object.entries(codes)) {
        const grouped = findings.filter(
          (item) => item.claimId === claimId && item.code === findingCode,
        )
        if (grouped.length)
          researchIssue(
            code,
            grouped.flatMap((item) =>
              item.relatedResearchClaimId ? [item.relatedResearchClaimId] : [],
            ),
            grouped.flatMap((item) => (item.sourceId ? [item.sourceId] : [])),
            category,
          )
      }
      if (original) {
        const partners = record.draft.facts.filter(
          (item) =>
            original.conflictsWith.includes(item.claimId) ||
            item.conflictsWith.includes(original.claimId),
        )
        if (partners.length)
          researchIssue(
            'research_claim_conflict',
            partners.map((item) => item.claimId),
            [original, ...partners].flatMap((item) => item.sourceRefs),
            'needs_human',
          )
      }
    }
  }
  return result
}
function disposition(
  report: PresentationDeliveryReport,
  issue: Pick<DeliveryIssue, 'id' | 'digest'>,
): DeliveryIssue['disposition'] {
  const last = [...report.issueLedger.actions]
    .reverse()
    .find((action) => action.issueId === issue.id)
  return last
    ? {
        state: last.issueDigest === issue.digest ? last.state : 'open',
        stale: last.issueDigest !== issue.digest,
        actionId: last.actionId,
      }
    : { state: 'open', stale: false }
}

/** Strict structural/semantic parsing. Digest authenticity is established by the server rebuilding
 * from frozen production before accepting an action; synchronous parsing never claims authenticity. */
export function parsePresentationDeliveryReport(value: unknown): PresentationDeliveryReport {
  exact(value, [
    'version',
    'projectId',
    'documentId',
    'requestId',
    'planRevision',
    'inputDigest',
    'planDigest',
    'plan',
    'research',
    'sourceAudit',
    'reviews',
    'issueLedger',
    'pages',
    'checks',
  ])
  if (new TextEncoder().encode(JSON.stringify(value)).byteLength > 8 * 1024 * 1024) invalid()
  const report = value as unknown as PresentationDeliveryReport
  if (
    report.version !== 1 ||
    !['projectId', 'requestId'].every(
      (key) =>
        typeof value[key] === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value[key] as string),
    ) ||
    typeof report.documentId !== 'string' ||
    !report.documentId.trim() ||
    report.documentId.length > 2048 ||
    !Number.isSafeInteger(report.planRevision) ||
    report.planRevision < 1 ||
    !digestValid(report.inputDigest) ||
    !digestValid(report.planDigest) ||
    canonical(report.checks) !== canonical(checks)
  )
    invalid()
  const plan = parsePresentationPlan(report.plan)
  if (plan.research) {
    exact(report.research, ['record', 'findings'])
    const record = parsePresentationResearchRecord(report.research!.record)
    if (
      record.documentId !== report.documentId ||
      canonical(report.research!.findings) !==
        canonical(presentationResearchBindingFindings(plan, record))
    )
      invalid()
  } else if (report.research !== undefined) invalid()
  if (report.sourceAudit !== undefined) {
    const expected = plan.sources.flatMap((source) => {
      const attachmentId = presentationSourceAttachmentId(source)
      return attachmentId ? [{ sourceId: source.id, attachmentId }] : []
    })
    if (!Array.isArray(report.sourceAudit) || report.sourceAudit.length !== expected.length)
      invalid()
    for (const [index, audit] of report.sourceAudit.entries()) {
      exact(audit, ['sourceId', 'attachmentId', 'status', 'offset', 'locator'])
      if (
        audit.sourceId !== expected[index]!.sourceId ||
        audit.attachmentId !== expected[index]!.attachmentId ||
        ![
          'found',
          'not_found',
          'empty_excerpt',
          'not_ready',
          'unsupported',
          'missing',
          'source_mismatch',
        ].includes(audit.status) ||
        (audit.status === 'found'
          ? !Number.isSafeInteger(audit.offset) ||
            audit.offset! < 0 ||
            audit.offset! > MAX_PRESENTATION_SOURCE_TEXT_CHARS
          : audit.offset !== undefined || audit.locator !== undefined) ||
        (audit.locator !== undefined && !/^第 [1-9]\d{0,5} (页|段)$/.test(audit.locator))
      )
        invalid()
    }
  }
  const ledger = parsePresentationIssueLedger(report.issueLedger)
  if (
    plan.projectId !== report.projectId ||
    ['projectId', 'documentId', 'requestId', 'inputDigest', 'planDigest'].some(
      (key) => ledger[key as keyof typeof ledger] !== value[key],
    )
  )
    invalid()
  if (
    !Array.isArray(report.reviews) ||
    report.reviews.length > 32 ||
    !Array.isArray(report.pages) ||
    report.pages.length !== plan.slides.length
  )
    invalid()
  const reviewIds = new Set<string>()
  for (const raw of report.reviews) {
    const review = parsePresentationClaimReview(raw)
    const page = plan.slides.find((page) => page.id === review.pageId)
    const claim = plan.claims.find((claim) => claim.id === review.claimId)
    const source = plan.sources.find((source) => source.id === review.sourceId)
    if (
      reviewIds.has(review.reviewId) ||
      ['projectId', 'requestId', 'planRevision', 'inputDigest', 'planDigest'].some(
        (key) => review[key as keyof typeof review] !== value[key],
      ) ||
      !page?.claimIds.includes(review.claimId) ||
      !claim?.sourceIds.includes(review.sourceId) ||
      !source ||
      (review.sourceAssessment !== undefined &&
        (review.sourceAssessment.timeliness.claimAsOf !== claim.asOf ||
          review.sourceAssessment.timeliness.sourceAsOf !== source.asOf ||
          (review.sourceAssessment.jurisdiction !== undefined &&
            review.sourceAssessment.jurisdiction.claimJurisdiction !== claim.jurisdiction))) ||
      presentationSourceAttachmentId(source) !== review.attachmentId
    )
      invalid()
    reviewIds.add(review.reviewId)
  }
  for (const [index, page] of report.pages.entries()) {
    exact(page, ['pageId', 'title', 'productionState', 'calculations', 'issues'])
    const slide = plan.slides[index]!
    if (report.reviews.filter((review) => review.pageId === slide.id).length > 32) invalid()
    if (
      page.pageId !== slide.id ||
      page.title !== slide.title ||
      !['pending', 'building', 'compiled', 'failed'].includes(page.productionState) ||
      !Array.isArray(page.issues) ||
      page.issues.length >
        (plan.research || report.reviews.some((review) => review.sourceAssessment) ? 1056 : 608)
    )
      invalid()
    const calculations = slide.claimIds
      .map((id) => plan.claims.find((claim) => claim.id === id)!)
      .filter((claim) => claim.type === 'calculation')
      .map(reproducePresentationCalculation)
    if (canonical(calculations) !== canonical(page.calculations)) invalid()
    const missing = new Set(
      page.issues
        .filter((issue) => issue.code === 'claim_text_not_found')
        .map((issue) => issue.claimId),
    )
    const expected = seeds(report, index, missing)
    if (expected.length !== page.issues.length) invalid()
    for (const [issueIndex, issue] of page.issues.entries()) {
      exact(issue, [
        'id',
        'code',
        'claimId',
        'sourceId',
        'research',
        'digest',
        'category',
        'disposition',
      ])
      const { digest, disposition: saved, ...seed } = issue
      if (
        !digestValid(digest) ||
        canonical(seed) !== canonical(expected[issueIndex]) ||
        canonical(saved) !== canonical(disposition(report, issue))
      )
        invalid()
    }
  }
  return structuredClone(report)
}

export async function buildPresentationDeliveryReport(
  input: PresentationDeliveryReportInput,
): Promise<PresentationDeliveryReport> {
  const plan = parsePresentationPlan(input.plan),
    deck = parsePresentationDeck(input.deck)
  const researchRecord = plan.research
    ? parsePresentationResearchRecord(input.researchRecord)
    : undefined
  if (plan.research) {
    const record = researchRecord!
    const digestBytes = await globalThis.crypto.subtle.digest(
      'SHA-256',
      new TextEncoder().encode(canonicalPresentationValue(record.draft)),
    )
    const digest = Array.from(new Uint8Array(digestBytes), (byte) =>
      byte.toString(16).padStart(2, '0'),
    ).join('')
    if (digest !== record.draftDigest) throw new Error('research_binding_invalid')
  }
  assertDeckMatchesPresentationPlan(deck, plan)
  if (
    input.pageStates.length !== plan.slides.length ||
    new Set(input.pageStates.map((page) => page.pageId)).size !== plan.slides.length ||
    input.pageStates.some((page) => !plan.slides.some((slide) => slide.id === page.pageId))
  )
    invalid()
  const report: PresentationDeliveryReport = {
    version: 1,
    ...input.metadata,
    plan,
    ...(plan.research
      ? {
          research: {
            record: researchRecord!,
            findings: presentationResearchBindingFindings(plan, researchRecord!),
          },
        }
      : {}),
    ...(input.sourceAudit ? { sourceAudit: input.sourceAudit } : {}),
    reviews: input.reviews,
    issueLedger: input.issueLedger,
    pages: [],
    checks,
  }
  for (const [index, page] of plan.slides.entries()) {
    summarizePresentationPageReviews(
      plan,
      deck,
      {
        projectId: report.projectId,
        requestId: report.requestId,
        planRevision: report.planRevision,
        inputDigest: report.inputDigest,
        planDigest: report.planDigest,
        pageId: page.id,
      },
      input.reviews,
    )
    const content = checkPresentationPageContent(plan, deck, page.id)
    const missing = new Set(
      content.findings
        .filter((finding) => finding.code === 'claim_text_not_found')
        .map((finding) => finding.claimId),
    )
    const calculations = page.claimIds
      .map((id) => plan.claims.find((claim) => claim.id === id)!)
      .filter((claim) => claim.type === 'calculation')
      .map(reproducePresentationCalculation)
    const issues: DeliveryIssue[] = []
    for (const seed of seeds(report, index, missing)) {
      const sourceAssessmentIssue = /^source_(authority|timeliness|jurisdiction)_review_/.test(
        seed.code,
      )
      const relevantReviews = input.reviews.filter(
        (review) =>
          (seed.code.startsWith('source_review_') ||
            (sourceAssessmentIssue && review.sourceAssessment !== undefined)) &&
          review.pageId === page.id &&
          review.claimId === seed.claimId &&
          (seed.sourceId === undefined || review.sourceId === seed.sourceId),
      )
      const calculation = seed.code.startsWith('calculation_')
        ? calculations.find((calculation) => calculation.claimId === seed.claimId)
        : undefined
      const content = canonical({
        inputDigest: report.inputDigest,
        planDigest: report.planDigest,
        pageId: page.id,
        code: seed.code,
        claimId: seed.claimId,
        ...(seed.sourceId ? { sourceId: seed.sourceId } : {}),
        ...(seed.sourceId && report.sourceAudit
          ? { sourceAudit: report.sourceAudit.find((item) => item.sourceId === seed.sourceId) }
          : {}),
        relevantReviews,
        ...(sourceAssessmentIssue
          ? {
              sourceAssessmentContext: {
                claim: plan.claims.find((claim) => claim.id === seed.claimId),
                source: plan.sources.find((source) => source.id === seed.sourceId),
                ...(report.research
                  ? {
                      research: {
                        ledgerId: report.research.record.id,
                        sequence: report.research.record.sequence,
                        draftDigest: report.research.record.draftDigest,
                        claimMapping: plan.research!.claims.find(
                          (mapping) => mapping.claimId === seed.claimId,
                        ),
                        sourceMapping: plan.research!.sources.find(
                          (mapping) => mapping.sourceId === seed.sourceId,
                        ),
                        claim: report.research.record.draft.facts.find(
                          (fact) =>
                            fact.claimId ===
                            plan.research!.claims.find(
                              (mapping) => mapping.claimId === seed.claimId,
                            )?.researchClaimId,
                        ),
                        source: report.research.record.draft.sources.find(
                          (source) =>
                            source.id ===
                            plan.research!.sources.find(
                              (mapping) => mapping.sourceId === seed.sourceId,
                            )?.researchSourceId,
                        ),
                        evidence: report.research.record.sources!.find(
                          (source) =>
                            source.sourceId ===
                            plan.research!.sources.find(
                              (mapping) => mapping.sourceId === seed.sourceId,
                            )?.researchSourceId,
                        ),
                      },
                    }
                  : {}),
              },
            }
          : {}),
        ...(seed.research ? { research: researchIssueContext(report, seed) } : {}),
        ...(calculation ? { calculation } : {}),
      })
      const bytes = await globalThis.crypto.subtle.digest(
        'SHA-256',
        new TextEncoder().encode(content),
      )
      const digest = Array.from(new Uint8Array(bytes), (byte) =>
        byte.toString(16).padStart(2, '0'),
      ).join('')
      issues.push({ ...seed, digest, disposition: disposition(report, { id: seed.id, digest }) })
    }
    report.pages.push({
      pageId: page.id,
      title: page.title,
      productionState: input.pageStates.find((state) => state.pageId === page.id)!.state,
      calculations,
      issues,
    })
  }
  return parsePresentationDeliveryReport(report)
}

function researchIssueContext(report: PresentationDeliveryReport, seed: Seed) {
  const record = report.research!.record
  const ids = new Set([seed.research!.researchClaimId, ...seed.research!.relatedClaimIds])
  const facts = record.draft.facts.filter((fact) => ids.has(fact.claimId))
  const sources = new Set([
    ...seed.research!.sourceIds,
    ...facts.flatMap((fact) => fact.sourceRefs),
  ])
  return {
    descriptor: seed.research,
    planClaim: report.plan.claims.find((claim) => claim.id === seed.claimId),
    facts,
    sources: record.draft.sources.filter((source) => sources.has(source.id)),
    evidence: record.sources!.filter((source) => sources.has(source.sourceId)),
  }
}

/** Encode all supplied text (including URI punctuation) as inert HTML character references. */
const safe = (value: unknown): string =>
  Array.from(String(value), (character) =>
    /[A-Za-z0-9 \u3400-\u9fff]/u.test(character) ? character : `&#${character.codePointAt(0)};`,
  ).join('')
export function presentationDeliveryMarkdown(value: PresentationDeliveryReport): string {
  const report = parsePresentationDeliveryReport(value)
  const lines = [
    '# Presentation evidence delivery',
    '',
    safe(report.plan.title),
    '',
    'Scope: frozen production. Content needs review. Source authority and timeliness NOT VERIFIED. Host NOT CHECKED; round trip NOT RUN.',
    'Attachment excerpt audit reflects current document files when this report was read; it is not part of the frozen production snapshot.',
    'Arithmetic reproduction checks IEEE double arithmetic only; inputs, source truth, units and conclusions require human judgment. Supported reviews are historical Agent judgments.',
    '分类：已核验（仅算术） / 待人工判断 / 无法核验。解释不关闭机器发现，过期处置恢复为待处理。',
    '',
    `Request: ${safe(report.requestId)}; plan revision: ${report.planRevision}`,
    `Input digest: ${report.inputDigest}; plan digest: ${report.planDigest}`,
  ]
  if (report.research) {
    const { record, findings } = report.research
    lines.push(
      '',
      '## Plan-bound frozen research (not fact verification)',
      `Ledger: ${safe(record.id)}; sequence: ${record.sequence}; draft digest: ${record.draftDigest}`,
      'Research completion means the record was archived. Source authority, timeliness and support remain NOT VERIFIED. Conflicts and reference gaps require review.',
      ...findings.map((finding) => `- Needs review: ${safe(JSON.stringify(finding))}`),
      'Complete original research record (including both conflict partners and unselected sources):',
      safe(JSON.stringify(record)),
    )
  }
  for (const page of report.pages) {
    lines.push(
      '',
      `## ${safe(page.pageId)} — ${safe(page.title)}`,
      `Production: ${page.productionState}`,
    )
    const claimIds = report.plan.slides.find((slide) => slide.id === page.pageId)!.claimIds
    lines.push(`Claims (see complete claim catalog): ${claimIds.map(safe).join(', ')}`)
    for (const calculation of page.calculations)
      lines.push(
        `Arithmetic ${calculation.status === 'reproduced' ? '已核验（仅算术） CHECKED (arithmetic only)' : '无法核验 / 待人工判断 UNVERIFIABLE / needs review'}: ${safe(JSON.stringify(calculation))}`,
      )
    for (const issue of page.issues)
      lines.push(
        `- ${issue.category === 'needs_human' ? '待人工判断' : '无法核验'} (${issue.category}): ${safe(issue.code)}; claim ${safe(issue.claimId)}${issue.sourceId ? `; source ${safe(issue.sourceId)}` : ''}; ${issue.disposition.state}; stale=${issue.disposition.stale}; issue ${safe(issue.id)}; digest ${issue.digest}`,
      )
    for (const issue of page.issues.filter((item) => item.research))
      lines.push(
        `- Frozen research issue context (not verified): ${safe(JSON.stringify(researchIssueContext(report, issue)))}`,
      )
  }
  lines.push('', '## Complete claim catalog')
  for (const claim of report.plan.claims) {
    lines.push(
      '',
      `### Claim ${safe(claim.id)}`,
      safe(claim.statement),
      `Sources (see source catalog): ${claim.sourceIds.map(safe).join(', ')}`,
    )
    if (claim.calculation)
      lines.push(
        `Formula: ${safe(claim.calculation.formula)}`,
        `Inputs and declared values: ${safe(JSON.stringify(claim.calculation))}`,
      )
  }
  lines.push('', '## Complete source catalog')
  for (const source of report.plan.sources) {
    lines.push(
      '',
      `### Source ${safe(source.id)}`,
      `Title: ${safe(source.title)}; URI: ${safe(source.uri)}; locator: ${safe(source.locator ?? '')}; as of: ${safe(source.asOf ?? '')}`,
      `Excerpt: ${safe(source.excerpt)}`,
      ...(report.sourceAudit?.find((item) => item.sourceId === source.id)
        ? [
            `Current attachment excerpt audit: ${safe(JSON.stringify(report.sourceAudit.find((item) => item.sourceId === source.id)))}`,
          ]
        : []),
    )
  }
  lines.push(
    '',
    '## Complete frozen plan (including unused claims and sources)',
    safe(JSON.stringify(report.plan)),
    '',
    '## All source review history',
  )
  for (const review of report.reviews) lines.push(`- ${safe(JSON.stringify(review))}`)
  if (report.reviews.some((review) => review.sourceAssessment))
    lines.push(
      '',
      'Source assessments above preserve every historical Agent opinion, literal basis and reference date. Mixed reference frames are differing judgments, not automatic factual contradictions. Positive opinions do not certify authority, timeliness or applicability; global checks remain NOT VERIFIED.',
    )
  lines.push('', '## All disposition history (explanations do not close machine findings)')
  for (const action of report.issueLedger.actions) lines.push(`- ${safe(JSON.stringify(action))}`)
  return lines.join('\n') + '\n'
}
