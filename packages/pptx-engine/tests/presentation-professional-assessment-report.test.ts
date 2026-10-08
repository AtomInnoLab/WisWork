import { expect, it } from 'vitest'
import {
  professionalAssessmentInput,
  professionalAssessmentReview,
} from './fixtures/presentation-professional-assessment'
import {
  buildPresentationDeliveryReport,
  parsePresentationDeliveryReport,
  presentationDeliveryMarkdown,
  presentationProfessionalIssueContext,
} from '../src/presentation-delivery-report'
it('groups all historical professional opinions by page claim source and keeps positive opinions from certifying or closing gaps', async () => {
  const input = professionalAssessmentInput()
  input.reviews = [professionalAssessmentReview(input)]
  const first = await buildPresentationDeliveryReport(input)
  const issues = first.pages[0]!.issues.filter((issue) =>
    issue.code.startsWith('professional_review_'),
  )
  expect(issues.map((issue) => issue.code)).toEqual([
    'professional_review_conclusion_scope_conflict',
    'professional_review_qualifications_uncertain',
  ])
  expect(
    issues.every((issue) => issue.category === 'needs_human' && issue.sourceId === 'source'),
  ).toBe(true)
  expect(
    first.pages
      .slice(1)
      .flatMap((page) => page.issues)
      .some((issue) => issue.code.startsWith('professional_review_')),
  ).toBe(false)
  const context = presentationProfessionalIssueContext(first, issues[0]!)!
  expect(context.claim?.professionalContext).toEqual(input.plan.claims[0]!.professionalContext)
  expect(context.reviews.map((review) => review.reviewId)).toEqual(['professional-review'])
  input.issueLedger = {
    ...input.issueLedger,
    revision: 1,
    actions: [
      {
        actionId: 'explain',
        issueId: issues[0]!.id,
        issueDigest: issues[0]!.digest,
        state: 'explained',
        note: '保留人工说明',
        sequence: 1,
        createdAt: '2026-09-29T00:00:01.000Z',
      },
    ],
  }
  expect(
    (await buildPresentationDeliveryReport(input)).pages[0]!.issues.find(
      (issue) => issue.id === issues[0]!.id,
    )?.disposition.state,
  ).toBe('explained')
  input.reviews.push(professionalAssessmentReview(input, 'later', ['consistent', 'consistent']))
  const mixed = await buildPresentationDeliveryReport(input)
  expect(mixed.pages[0]!.issues.map((issue) => issue.code)).toContain(
    'professional_review_conclusion_scope_mixed',
  )
  expect(mixed.pages[0]!.issues.map((issue) => issue.code)).toContain(
    'professional_review_qualifications_mixed',
  )
  expect(mixed.checks.sourceAuthority).toBe('not_verified')
  expect(mixed.checks.timeliness).toBe('not_verified')
  expect(presentationDeliveryMarkdown(mixed)).toContain('Professional assessment history')
  expect(presentationDeliveryMarkdown(mixed)).toContain('原文逐字依据')
})
it('changes related professional digests when history changes without changing legacy unassessed digests', async () => {
  const input = professionalAssessmentInput()
  input.reviews = [professionalAssessmentReview(input)]
  const first = await buildPresentationDeliveryReport(input)
  const issue = first.pages[0]!.issues.find(
    (issue) => issue.code === 'professional_review_conclusion_scope_conflict',
  )!
  const legacy = first.pages[0]!.issues.find((issue) => issue.code === 'claim_text_not_found')!
  input.issueLedger = {
    ...input.issueLedger,
    revision: 1,
    actions: [
      {
        actionId: 'explain',
        issueId: issue.id,
        issueDigest: issue.digest,
        state: 'explained',
        note: '原说明',
        sequence: 1,
        createdAt: '2026-09-29T00:00:01.000Z',
      },
    ],
  }
  input.reviews.push(professionalAssessmentReview(input, 'second'))
  const changed = await buildPresentationDeliveryReport(input)
  expect(changed.pages[0]!.issues.find((item) => item.id === issue.id)?.disposition).toMatchObject({
    state: 'open',
    stale: true,
  })
  expect(changed.pages[0]!.issues.find((item) => item.id === legacy.id)?.digest).toBe(legacy.digest)
})
it('rejects mismatched full context and forged missing issues while retaining old review shape', async () => {
  const input = professionalAssessmentInput()
  input.reviews = [professionalAssessmentReview(input)]
  const report = await buildPresentationDeliveryReport(input)
  const altered = structuredClone(report)
  altered.reviews[0]!.sourceAssessment!.professional!.context = { domain: 'science' }
  expect(() => parsePresentationDeliveryReport(altered)).toThrow(
    'presentation_delivery_report_invalid',
  )
  const gaps = structuredClone(report)
  gaps.pages[0]!.issues = gaps.pages[0]!.issues.filter(
    (issue) => !issue.code.startsWith('professional_review_'),
  )
  expect(() => parsePresentationDeliveryReport(gaps)).toThrow()
  delete input.reviews[0]!.sourceAssessment!.professional
  const old = await buildPresentationDeliveryReport(input)
  expect(
    old.pages
      .flatMap((page) => page.issues)
      .some((issue) => issue.code.startsWith('professional_review_')),
  ).toBe(false)
  expect(parsePresentationDeliveryReport(old)).toEqual(old)
})
it('uses finance aspects, permits nonapplicable forecast without certification and preserves law qualifications', async () => {
  for (const context of [
    { domain: 'finance', unit: '万元', currency: 'CNY', accountingBasis: '审计口径' },
    { domain: 'law', jurisdiction: '中国', effectLevel: '法律', limitations: '只适用于明确案件' },
  ] as const) {
    const input = professionalAssessmentInput(context)
    input.reviews = [
      professionalAssessmentReview(
        input,
        'review',
        context.domain === 'finance' ? ['conflict', 'not_applicable'] : ['conflict', 'conflict'],
      ),
    ]
    const report = await buildPresentationDeliveryReport(input)
    expect(report.pages[0]!.issues.map((issue) => issue.code)).toContain(
      context.domain === 'finance'
        ? 'professional_review_comparability_conflict'
        : 'professional_review_qualifications_conflict',
    )
    expect(report.pages[0]!.issues.map((issue) => issue.code)).not.toContain(
      'professional_review_forecast_not_applicable',
    )
  }
})
it('keeps professional issue context isolated when the same claim and source appear on multiple pages', async () => {
  const input = professionalAssessmentInput(),
    second = professionalAssessmentReview(input, 'second-page')
  second.pageId = input.plan.slides[1]!.id
  input.reviews = [professionalAssessmentReview(input), second]
  const report = await buildPresentationDeliveryReport(input),
    issue = report.pages[1]!.issues.find(
      (issue) => issue.code === 'professional_review_conclusion_scope_conflict',
    )!
  expect(
    presentationProfessionalIssueContext(report, issue)?.reviews.map((review) => review.reviewId),
  ).toEqual(['second-page'])
})
