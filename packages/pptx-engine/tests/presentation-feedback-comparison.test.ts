import { expect, it } from 'vitest'
import {
  buildPresentationFeedbackComparison,
  parsePresentationFeedbackComparison,
  type PresentationFeedbackComparisonInput,
} from '../src/presentation-feedback-comparison'
import { PRESENTATION_DOMAIN_PROFILES } from '../src/presentation-plan'
import { benchmarkPlan } from './fixtures/presentation-plan'
function fixture(): PresentationFeedbackComparisonInput {
  const plan = benchmarkPlan()
  const candidate = structuredClone(plan)
  candidate.domain = 'pitch'
  candidate.slides.forEach(
    (s, i) => (s.domainSection = PRESENTATION_DOMAIN_PROFILES.pitch.sections[i % 5]),
  )
  const task = (requestId: string, plan: typeof candidate, corrected: boolean) => ({
    requestId,
    inputDigest: 'a'.repeat(64),
    planDigest: 'b'.repeat(64),
    planRevision: 1,
    plan,
    feedback: {
      version: 1 as const,
      source: 'user_reported' as const,
      projectId: plan.projectId,
      documentId: 'doc',
      requestId,
      inputDigest: 'a'.repeat(64),
      planDigest: 'b'.repeat(64),
      planRevision: 1,
      pageIds: plan.slides.map((s) => s.id),
      revision: 1,
      snapshots: [
        {
          revision: 1,
          recordedAt: '2026-09-29T00:00:00.000Z',
          pages: plan.slides.map((s) => ({
            pageId: s.id,
            status: corrected ? ('needs_correction' as const) : ('no_correction' as const),
            note: 'must not export notes',
          })),
        },
      ],
    },
  })
  return {
    documentId: 'doc',
    projectId: plan.projectId,
    baseline: task('base', plan, true),
    candidate: task('industry', candidate, false),
  }
}
it('derives descriptive comparable counts from two full frozen plans without notes/history or certification', () => {
  const report = buildPresentationFeedbackComparison(fixture())
  expect(report).toMatchObject({
    source: 'user_reported',
    effect: 'not_verified',
    comparable: true,
    gaps: [],
    delta: { needsCorrectionPages: -8, needsCorrectionRate: -1 },
  })
  expect(report.conditions).toHaveLength(7)
  expect(JSON.stringify(report)).not.toContain('must not export notes')
  expect(parsePresentationFeedbackComparison(report)).toEqual(report)
})
it('rejects forged derived fields, statuses, flags, scope and extra keys', () => {
  const report = buildPresentationFeedbackComparison(fixture())
  for (const altered of [
    { ...report, extra: true },
    { ...report, effect: 'verified' },
    { ...report, comparable: false },
    { ...report, delta: { needsCorrectionPages: 0, needsCorrectionRate: 0 } },
    { ...report, conditions: [] },
    { ...report, gaps: ['input_conditions_differ'] },
  ])
    expect(() => parsePresentationFeedbackComparison(altered)).toThrow(
      'presentation_feedback_comparison_invalid',
    )
  const status = structuredClone(report)
  status.candidate.pages[0]!.status = 'needs_correction'
  expect(() => parsePresentationFeedbackComparison(status)).toThrow(
    'presentation_feedback_comparison_invalid',
  )
})
it.each(['brief', 'sources', 'claims', 'research', 'style', 'brandKit', 'parallelism'] as const)(
  'checks the full frozen %s condition rather than titles or domain labels',
  (key) => {
    const input = fixture(),
      plan = input.candidate.plan
    if (key === 'brief') plan.brief.constraints.push('附加限制')
    if (key === 'sources') plan.sources[0]!.asOf = '2024-12-01'
    if (key === 'claims') plan.claims[0]!.statement += '不同主张'
    if (key === 'research')
      plan.research = {
        ledgerId: 'ledger',
        sequence: 1,
        draftDigest: 'c'.repeat(64),
        sources: plan.sources.map((s) => ({ sourceId: s.id, researchSourceId: s.id })),
        claims: plan.claims.map((c) => ({ claimId: c.id, researchClaimId: c.id })),
      }
    if (key === 'style') plan.style.fontFace = 'Courier New'
    if (key === 'brandKit')
      plan.brandKit = {
        id: 'brand',
        revision: 1,
        name: '实际规则',
        allowedColors: [plan.style.background, plan.style.textColor, plan.style.accentColor],
      }
    if (key === 'parallelism') {
      plan.parallelism = 2
      plan.slides.forEach((s) => (s.dependsOn = []))
    }
    const report = buildPresentationFeedbackComparison(input)
    expect(report.conditions.find((c) => c.key === key)?.match).toBe(false)
    expect(report.gaps).toContain('input_conditions_differ')
    expect(report.delta).toBeNull()
    expect(parsePresentationFeedbackComparison(report)).toEqual(report)
  },
)
it('treats omitted and explicit serial execution alike while keeping output title/pages outside input conditions', () => {
  const input = fixture()
  input.candidate.plan.parallelism = 1
  input.candidate.plan.title = '不同标题'
  input.candidate.plan.slides[0]!.title = '不同页面标题'
  const report = buildPresentationFeedbackComparison(input)
  expect(report.comparable).toBe(true)
  expect(report.candidate.plan.title).toBe('不同标题')
  report.candidate.plan.brief.objective = '污染'
  expect(buildPresentationFeedbackComparison(input).candidate.plan.brief.objective).not.toBe('污染')
})
it('keeps missing or incomplete user evaluations unknown and never computes an effect delta', () => {
  for (const side of ['baseline', 'candidate'] as const) {
    const missing = fixture()
    missing[side].feedback = null
    const report = buildPresentationFeedbackComparison(missing)
    expect(report[side].feedbackRevision).toBeNull()
    expect(report[side].counts.needsCorrectionRate).toBeNull()
    expect(report.gaps).toContain(`${side}_feedback_missing`)
    expect(report.gaps).toContain(`${side}_not_fully_evaluated`)
    expect(report.delta).toBeNull()
    const partial = fixture()
    partial[side].feedback!.snapshots[0]!.pages[0]!.status = 'not_evaluated'
    const value = buildPresentationFeedbackComparison(partial)
    expect(value[side].counts.evaluatedPages).toBe(7)
    expect(value.delta).toBeNull()
    expect(parsePresentationFeedbackComparison(value)).toEqual(value)
  }
})
it('requires genuinely absent baseline domain and one of the five candidate industry roles', () => {
  const sameDomain = fixture()
  sameDomain.baseline.plan = structuredClone(sameDomain.candidate.plan)
  const result = buildPresentationFeedbackComparison(sameDomain)
  expect(result.gaps).toContain('baseline_not_generic')
  expect(result.delta).toBeNull()
  const generic = fixture()
  delete generic.candidate.plan.domain
  generic.candidate.plan.slides.forEach((s) => delete s.domainSection)
  expect(buildPresentationFeedbackComparison(generic).gaps).toContain('candidate_not_industry')
  const professional = fixture()
  professional.candidate.plan.domain = 'science'
  professional.candidate.plan.slides.forEach(
    (s, i) => (s.domainSection = PRESENTATION_DOMAIN_PROFILES.science.sections[i % 5]),
  )
  expect(buildPresentationFeedbackComparison(professional).gaps).toContain('candidate_not_industry')
})
it('accepts different page counts with actual denominators and negative or positive descriptive deltas', () => {
  const input = fixture()
  input.baseline.plan.slides = input.baseline.plan.slides.slice(0, 2)
  input.baseline.feedback!.pageIds = input.baseline.feedback!.pageIds.slice(0, 2)
  input.baseline.feedback!.snapshots[0]!.pages = input.baseline.feedback!.snapshots[0]!.pages.slice(
    0,
    2,
  )
  input.candidate.feedback!.snapshots[0]!.pages[0]!.status = 'needs_correction'
  const report = buildPresentationFeedbackComparison(input)
  expect(report.baseline.counts.totalPages).toBe(2)
  expect(report.candidate.counts.totalPages).toBe(8)
  expect(report.delta).toEqual({ needsCorrectionPages: -1, needsCorrectionRate: -0.875 })
})
it('rejects forged original ledger scope/digests/revision/history and wrong report page order or feedback time', () => {
  for (const key of [
    'projectId',
    'documentId',
    'requestId',
    'inputDigest',
    'planDigest',
    'planRevision',
  ] as const) {
    const input = fixture()
    Object.assign(input.candidate.feedback!, {
      [key]: key === 'planRevision' ? 2 : key.endsWith('Digest') ? 'd'.repeat(64) : 'foreign',
    })
    expect(() => buildPresentationFeedbackComparison(input)).toThrow(
      'presentation_feedback_comparison_invalid',
    )
  }
  const malformed = fixture()
  malformed.candidate.feedback!.revision = 2
  expect(() => buildPresentationFeedbackComparison(malformed)).toThrow(
    'presentation_feedback_comparison_invalid',
  )
  const valid = buildPresentationFeedbackComparison(fixture())
  for (const field of ['feedbackRevision', 'feedbackRecordedAt', 'requestId'] as const) {
    const changed = structuredClone(valid)
    Object.assign(changed.candidate, {
      [field]:
        field === 'feedbackRevision'
          ? null
          : field === 'feedbackRecordedAt'
            ? '2026-02-30T00:00:00.000Z'
            : changed.baseline.requestId,
    })
    expect(() => parsePresentationFeedbackComparison(changed)).toThrow(
      'presentation_feedback_comparison_invalid',
    )
  }
  const reordered = structuredClone(valid)
  reordered.candidate.pages.reverse()
  expect(() => parsePresentationFeedbackComparison(reordered)).toThrow(
    'presentation_feedback_comparison_invalid',
  )
  const extra = structuredClone(valid)
  Object.assign(extra.candidate.pages[0]!, { note: 'do not copy original notes' })
  expect(() => parsePresentationFeedbackComparison(extra)).toThrow(
    'presentation_feedback_comparison_invalid',
  )
})
it('selects only the latest genuine revision and rejects fabricated count metadata', () => {
  const input = fixture()
  const feedback = input.candidate.feedback!
  feedback.revision = 2
  feedback.snapshots.push({
    revision: 2,
    recordedAt: '2026-09-29T00:00:01.000Z',
    pages: feedback.snapshots[0]!.pages.map((page, index) => ({
      ...page,
      status: index === 0 ? 'needs_correction' : 'no_correction',
      note: '最新说明也不导出',
    })),
  })
  const report = buildPresentationFeedbackComparison(input)
  expect(report.candidate.feedbackRevision).toBe(2)
  expect(report.candidate.feedbackRecordedAt).toBe('2026-09-29T00:00:01.000Z')
  expect(report.candidate.counts.needsCorrectionPages).toBe(1)
  expect(JSON.stringify(report)).not.toContain('最新说明')
  expect(report.candidate).not.toHaveProperty('snapshots')
  for (const key of [
    'totalPages',
    'evaluatedPages',
    'needsCorrectionPages',
    'noCorrectionPages',
    'notEvaluatedPages',
    'needsCorrectionRate',
  ] as const) {
    const altered = structuredClone(report)
    altered.candidate.counts[key] = 99
    expect(() => parsePresentationFeedbackComparison(altered)).toThrow(
      'presentation_feedback_comparison_invalid',
    )
  }
  const invalid = fixture()
  invalid.candidate.plan.parallelism = 0 as never
  expect(() => buildPresentationFeedbackComparison(invalid)).toThrow(
    'presentation_feedback_comparison_invalid',
  )
})
