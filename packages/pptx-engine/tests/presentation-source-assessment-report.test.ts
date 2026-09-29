import { describe, expect, it } from 'vitest'
import { benchmarkPlan } from './fixtures/presentation-plan'
import { benchmarkDeck } from './fixtures/presentation-benchmark'
import { presentationPlanClaims } from '../src/presentation-plan'
import {
  buildPresentationDeliveryReport,
  parsePresentationDeliveryReport,
  presentationDeliveryMarkdown,
  type PresentationDeliveryReportInput,
} from '../src/presentation-delivery-report'
import { parsePresentationClaimReview } from '../src/presentation-claim-review'

function fixture(): PresentationDeliveryReportInput {
  const plan = benchmarkPlan(),
    deck = benchmarkDeck()
  plan.sources[0]!.uri = `attachment:${'a'.repeat(64)}`
  plan.claims[0]!.jurisdiction = 'China'
  deck.claims = presentationPlanClaims(plan)
  const metadata = {
    projectId: plan.projectId,
    documentId: 'doc',
    requestId: 'run',
    planRevision: 1,
    inputDigest: 'b'.repeat(64),
    planDigest: 'c'.repeat(64),
  }
  return {
    plan,
    deck,
    metadata,
    reviews: [],
    pageStates: plan.slides.map((slide) => ({ pageId: slide.id, state: 'pending' })),
    issueLedger: {
      version: 1,
      projectId: plan.projectId,
      documentId: 'doc',
      requestId: 'run',
      inputDigest: metadata.inputDigest,
      planDigest: metadata.planDigest,
      revision: 0,
      actions: [],
    },
  }
}
function assessed(input: PresentationDeliveryReportInput, reviewId = 'review') {
  return {
    version: 1 as const,
    projectId: input.plan.projectId,
    requestId: 'run',
    planRevision: 1,
    inputDigest: input.metadata.inputDigest,
    planDigest: input.metadata.planDigest,
    reviewId,
    pageId: input.plan.slides[0]!.id,
    claimId: 'source-1',
    sourceId: 'source',
    attachmentId: 'a'.repeat(64),
    offset: 0,
    maxChars: 50,
    evidenceDigest: 'd'.repeat(64),
    outcome: 'supported' as const,
    notes: 'Agent review',
    reviewer: 'agent' as const,
    createdAt: '2026-09-29T00:00:00.000Z',
    sourceAssessment: {
      scope: 'Only this claim',
      authority: {
        outcome: 'appropriate_for_claim' as const,
        sourceTier: 'primary' as const,
        reason: 'The original source is appropriate',
      },
      timeliness: {
        outcome: 'current_for_claim' as const,
        referenceDate: '2026-09-29',
        reason: 'Current within this frame',
      },
      jurisdiction: {
        claimJurisdiction: 'China',
        outcome: 'applicable' as const,
        reason: 'Applies to this scope',
      },
      basis: [{ offset: 0, text: 'Original evidence' }],
    },
    checks: {
      support: 'agent_reviewed' as const,
      sourceAuthority: 'not_verified' as const,
      timeliness: 'not_verified' as const,
      host: 'not_checked' as const,
    },
  }
}
describe('historical source assessment report', () => {
  it('parses complete assessment but rejects unknown fields and forged global checks', () => {
    const input = fixture(),
      review = assessed(input)
    expect(parsePresentationClaimReview(review).sourceAssessment).toEqual(review.sourceAssessment)
    expect(() =>
      parsePresentationClaimReview({
        ...review,
        sourceAssessment: { ...review.sourceAssessment, extra: true },
      }),
    ).toThrow()
  })
  it('does not certify positive assessments and reports missing dimensions on other pages', async () => {
    const input = fixture()
    input.reviews = [assessed(input)]
    const report = await buildPresentationDeliveryReport(input)
    expect(
      report.pages[0]!.issues.filter((issue) =>
        /source_(authority|timeliness|jurisdiction)_review_/.test(issue.code),
      ),
    ).toEqual([])
    expect(
      report.pages
        .slice(1)
        .flatMap((page) => page.issues)
        .map((issue) => issue.code),
    ).toContain('source_authority_review_missing')
    expect(report.checks.sourceAuthority).toBe('not_verified')
    expect(presentationDeliveryMarkdown(report)).toContain('Original evidence')
  })
  it('keeps all opinions and differing reference frames instead of choosing the latest', async () => {
    const input = fixture(),
      first = assessed(input),
      second = assessed(input, 'later')
    second.sourceAssessment.authority.sourceTier =
      'secondary' as typeof first.sourceAssessment.authority.sourceTier
    second.sourceAssessment.timeliness.referenceDate = '2025-09-29'
    second.sourceAssessment.jurisdiction.outcome =
      'uncertain' as typeof first.sourceAssessment.jurisdiction.outcome
    input.reviews = [first, second]
    const report = await buildPresentationDeliveryReport(input)
    expect(report.pages[0]!.issues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining([
        'source_authority_review_mixed',
        'source_timeliness_review_mixed',
        'source_jurisdiction_review_mixed',
      ]),
    )
    const forged = structuredClone(report)
    forged.pages[0]!.issues = forged.pages[0]!.issues.filter(
      (issue) => issue.code !== 'source_authority_review_mixed',
    )
    expect(() => parsePresentationDeliveryReport(forged)).toThrow()
  })
  it('preserves literal basis controls and rejects mismatched frozen labels', async () => {
    const input = fixture(),
      review = assessed(input)
    review.sourceAssessment.basis[0]!.text = 'Original\f evidence'
    expect(parsePresentationClaimReview(review).sourceAssessment?.basis[0]?.text).toBe(
      'Original\f evidence',
    )
    input.reviews = [review]
    await expect(buildPresentationDeliveryReport(input)).resolves.toBeDefined()
    review.sourceAssessment.timeliness = {
      ...review.sourceAssessment.timeliness,
      claimAsOf: 'invented',
    } as typeof review.sourceAssessment.timeliness
    await expect(buildPresentationDeliveryReport(input)).rejects.toThrow()
  })
  it('retains all negative dimensions and makes explained history stale after reasoning changes', async () => {
    const input = fixture(),
      review = assessed(input)
    review.sourceAssessment.authority.outcome =
      'insufficient_authority' as typeof review.sourceAssessment.authority.outcome
    review.sourceAssessment.timeliness.outcome =
      'historical_only' as typeof review.sourceAssessment.timeliness.outcome
    review.sourceAssessment.jurisdiction.outcome =
      'mismatch' as typeof review.sourceAssessment.jurisdiction.outcome
    input.reviews = [review]
    const before = await buildPresentationDeliveryReport(input)
    expect(before.pages[0]!.issues.map((item) => item.code)).toEqual(
      expect.arrayContaining([
        'source_authority_review_insufficient',
        'source_timeliness_review_historical_only',
        'source_jurisdiction_review_mismatch',
      ]),
    )
    const issue = before.pages[0]!.issues.find(
      (item) => item.code === 'source_authority_review_insufficient',
    )!
    input.issueLedger = {
      ...input.issueLedger,
      revision: 1,
      actions: [
        {
          actionId: 'explain',
          issueId: issue.id,
          issueDigest: issue.digest,
          state: 'explained',
          note: 'Human explanation',
          sequence: 1,
          createdAt: '2026-09-29T00:03:00.000Z',
        },
      ],
    }
    const explained = await buildPresentationDeliveryReport(input)
    expect(explained.pages[0]!.issues.find((item) => item.id === issue.id)?.disposition).toEqual({
      state: 'explained',
      stale: false,
      actionId: 'explain',
    })
    review.sourceAssessment.authority.reason = 'Changed judgment about scope'
    const after = await buildPresentationDeliveryReport(input)
    expect(after.pages[0]!.issues.find((item) => item.id === issue.id)?.disposition).toEqual({
      state: 'open',
      stale: true,
      actionId: 'explain',
    })
    expect(after.reviews).toHaveLength(1)
  })
  it('reports uniform uncertainty and superseded time frames without treating them as verified', async () => {
    const input = fixture(),
      review = assessed(input)
    review.sourceAssessment.authority.outcome =
      'uncertain' as typeof review.sourceAssessment.authority.outcome
    review.sourceAssessment.timeliness.outcome =
      'uncertain' as typeof review.sourceAssessment.timeliness.outcome
    review.sourceAssessment.jurisdiction.outcome =
      'uncertain' as typeof review.sourceAssessment.jurisdiction.outcome
    input.reviews = [review]
    const report = await buildPresentationDeliveryReport(input)
    expect(report.pages[0]!.issues.map((item) => item.code)).toEqual(
      expect.arrayContaining([
        'source_authority_review_uncertain',
        'source_timeliness_review_uncertain',
        'source_jurisdiction_review_uncertain',
      ]),
    )
    review.sourceAssessment.timeliness.outcome =
      'superseded' as typeof review.sourceAssessment.timeliness.outcome
    expect(
      (await buildPresentationDeliveryReport(input)).pages[0]!.issues.map((item) => item.code),
    ).toContain('source_timeliness_review_superseded')
  })
})
