import { expect, it } from 'vitest'
import * as plan from '../src/presentation-plan'
const domains = ['pitch', 'report', 'training', 'research', 'sales'] as const
it.each(domains)('provides concrete independent %s narrative and executable guidance', (domain) => {
  const fn = plan.presentationDomainWorkflow
  expect(typeof fn).toBe('function')
  const a = fn(domain)!,
    b = fn(domain)!
  expect(a).toMatchObject({ version: 1, domain })
  expect(a.sections.map((s) => s.id)).toEqual(plan.PRESENTATION_DOMAIN_PROFILES[domain].sections)
  expect(a.sections.every((s) => s.title && s.instruction.length > 30)).toBe(true)
  expect(a.reviewSteps.length).toBeGreaterThanOrEqual(4)
  expect(a.manualChecks.length).toBeGreaterThanOrEqual(3)
  expect(a.disclosure).toContain('不')
  ;(a.sections[0]! as { instruction: string }).instruction = '污染'
  ;(a.reviewSteps[0]!.tools as string[]).push('fake')
  expect(fn(domain)).toEqual(b)
})
it('does not alter professional or unknown domain contract', () => {
  const fn = plan.presentationDomainWorkflow
  expect(typeof fn).toBe('function')
  for (const d of [undefined, 'unknown', 'science', 'law', 'finance']) expect(fn(d)).toBeUndefined()
})
it.each(domains)(
  'requires visible confirmed import before %s host comparison and screenshots',
  (domain) => {
    const steps = plan.presentationDomainWorkflow(domain)!.reviewSteps
    const imported = steps.findIndex((s) => s.tools.includes('import_presentation_production'))
    const compared = steps.findIndex((s) => s.tools.includes('compare_presentation_page_structure'))
    expect(imported).toBeGreaterThanOrEqual(0)
    expect(compared).toBeGreaterThan(imported)
    expect(steps[imported]!.tools).toContain('prepare_presentation_production_import')
    expect(steps[imported]!.instruction).toContain('用户确认')
    expect(steps[compared]!.instruction).toContain('回执')
  },
)
it.each(domains)(
  'freezes %s production before claim review and reviews before import',
  (domain) => {
    const steps = plan.presentationDomainWorkflow(domain)!.reviewSteps
    const start = steps.findIndex((s) => s.tools.includes('start_presentation_production'))
    const review = steps.findIndex((s) => s.tools.includes('read_presentation_claim_evidence'))
    const imported = steps.findIndex((s) => s.tools.includes('import_presentation_production'))
    expect(start).toBeLessThan(review)
    expect(review).toBeLessThan(imported)
    expect(steps[start]!.instruction).toContain('request_id')
    expect(steps[review]!.instruction).toContain('request_id')
  },
)
