import { expect, it } from 'vitest'
import { createHash } from 'node:crypto'
import { canonicalPresentationValue } from '@wiswork/project-store/presentation-canonical'
import { benchmarkPlan } from './fixtures/presentation-plan'
import { parsePresentationPlan } from '../src/presentation-plan'
import { presentationPlanClaims } from '../src/presentation-plan'
import { benchmarkDeck } from './fixtures/presentation-benchmark'
import {
  buildPresentationDeliveryReport,
  parsePresentationDeliveryReport,
  presentationDeliveryMarkdown,
} from '../src/presentation-delivery-report'
import { researchFixture } from './fixtures/presentation-research'
import { assertPresentationResearchBinding } from '../src/presentation-research-binding'

function input() {
  const plan = benchmarkPlan(),
    deck = benchmarkDeck()
  return {
    plan,
    deck,
    metadata: {
      projectId: plan.projectId,
      documentId: 'doc',
      requestId: 'run',
      planRevision: 1,
      inputDigest: 'b'.repeat(64),
      planDigest: 'c'.repeat(64),
    },
    reviews: [],
    pageStates: plan.slides.map((slide) => ({ pageId: slide.id, state: 'pending' as const })),
    issueLedger: {
      version: 1 as const,
      projectId: plan.projectId,
      documentId: 'doc',
      requestId: 'run',
      inputDigest: 'b'.repeat(64),
      planDigest: 'c'.repeat(64),
      revision: 0,
      actions: [],
    },
  }
}

it('preserves partial professional context and rejects cross-domain fields', () => {
  const plan = benchmarkPlan()
  Object.assign(plan.claims[0]!, { professionalContext: { domain: 'science' } })
  expect(parsePresentationPlan(plan).claims[0]).toHaveProperty(
    'professionalContext.domain',
    'science',
  )
  Object.assign(plan.claims[0]!, { professionalContext: { domain: 'science', currency: 'USD' } })
  expect(() => parsePresentationPlan(plan)).toThrow()
})

it('requires exact research professional context including omitted values', () => {
  const { plan, record } = researchFixture()
  record.draft.facts[0]!.professionalContext = { domain: 'science', sample: 'Declared sample' }
  expect(() => assertPresentationResearchBinding(plan, record)).toThrow('research_binding_invalid')
  plan.claims[0]!.professionalContext = { ...record.draft.facts[0]!.professionalContext }
  expect(() => assertPresentationResearchBinding(plan, record)).not.toThrow()
  plan.claims[0]!.professionalContext.sample = 'Changed sample'
  expect(() => assertPresentationResearchBinding(plan, record)).toThrow()
})

it('warns once per mapped professional claim about original secondary tier', async () => {
  const value = input(),
    { plan, record } = researchFixture()
  plan.claims[0]!.professionalContext = { domain: 'science' }
  record.draft.facts[0]!.professionalContext = { domain: 'science' }
  record.draft.facts[0]!.sourceTier = 'secondary'
  record.draftDigest = createHash('sha256')
    .update(canonicalPresentationValue(record.draft))
    .digest('hex')
  plan.research!.draftDigest = record.draftDigest
  value.plan = plan
  value.deck.claims = presentationPlanClaims(plan)
  const report = await buildPresentationDeliveryReport({ ...value, researchRecord: record })
  expect(
    report.pages[0]!.issues.filter((issue) => issue.code === 'professional_source_secondary'),
  ).toHaveLength(1)
  expect(
    report.pages[0]!.issues.find((issue) => issue.code === 'professional_context_incomplete')
      ?.category,
  ).toBe('unverifiable')
})

it('reports incomplete and legal warnings with inclusive dates and strict report findings', async () => {
  const value = input(),
    claim = value.plan.claims[0]!
  claim.jurisdiction = 'General'
  claim.professionalContext = {
    domain: 'law',
    jurisdiction: 'Professional',
    effectiveFrom: '2026-01-01',
    effectiveUntil: '2026-12-31',
    applicabilityDate: '2026-01-01',
  }
  value.deck.claims = presentationPlanClaims(value.plan)
  const boundary = await buildPresentationDeliveryReport(value)
  const codes = boundary.pages[0]!.issues.map((issue) => issue.code)
  expect(codes).toContain('professional_context_incomplete')
  expect(codes).toContain('professional_jurisdiction_mismatch')
  expect(codes).not.toContain('professional_legal_rule_inactive')
  claim.professionalContext.applicabilityDate = '2025-12-31'
  const inactive = await buildPresentationDeliveryReport(value)
  expect(inactive.pages[0]!.issues.map((issue) => issue.code)).toContain(
    'professional_legal_rule_inactive',
  )
  const forged = structuredClone(inactive)
  forged.pages[0]!.issues = forged.pages[0]!.issues.filter(
    (issue) => issue.code !== 'professional_context_incomplete',
  )
  expect(() => parsePresentationDeliveryReport(forged)).toThrow()
  expect(presentationDeliveryMarkdown(inactive)).toContain('missing fields')
  expect(inactive.checks.sourceAuthority).toBe('not_verified')
})

it('retains financial disagreements without certifying calculations and changes digest', async () => {
  const value = input(),
    claim = value.plan.claims[0]!
  claim.asOf = '2026-09-29'
  claim.calculation = { formula: '1+1', inputs: ['1', '1'], unit: 'thousands', currency: 'USD' }
  claim.professionalContext = {
    domain: 'finance',
    asOf: '2025-09-29',
    unit: 'millions',
    currency: 'CNY',
  }
  value.deck.claims = presentationPlanClaims(value.plan)
  const first = await buildPresentationDeliveryReport(value)
  expect(first.pages[0]!.issues.map((issue) => issue.code)).toEqual(
    expect.arrayContaining([
      'professional_financial_time_mixed',
      'professional_financial_unit_mismatch',
      'professional_financial_currency_mismatch',
    ]),
  )
  const issue = first.pages[0]!.issues.find(
    (issue) => issue.code === 'professional_financial_unit_mismatch',
  )!
  claim.professionalContext.limitations = 'Declared limitation'
  const second = await buildPresentationDeliveryReport(value)
  expect(second.pages[0]!.issues.find((item) => item.id === issue.id)?.digest).not.toBe(
    issue.digest,
  )
})
