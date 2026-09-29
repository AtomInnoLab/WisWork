import { expect, it } from 'vitest'
import {
  presentationProfessionalWorkflow,
  PRESENTATION_DOMAIN_PROFILES,
  parsePresentationPlan,
} from '../src/presentation-plan'
import { benchmarkPlan } from './fixtures/presentation-plan'
it.each(['science', 'law', 'finance'] as const)(
  'provides independent %s workflow and strict matching sections',
  (domain) => {
    const workflow = presentationProfessionalWorkflow(domain)!
    expect(workflow.domain).toBe(domain)
    expect(workflow.version).toBe(1)
    expect(workflow.sourcePriority.length).toBeGreaterThan(2)
    expect(workflow.reviewSteps).toHaveLength(6)
    expect(workflow.manualChecks.length).toBeGreaterThan(2)
    const plan = benchmarkPlan()
    plan.domain = domain
    plan.slides = PRESENTATION_DOMAIN_PROFILES[domain].sections.map((domainSection, index) => ({
      ...plan.slides[0]!,
      id: 'p' + index,
      domainSection,
    }))
    expect(parsePresentationPlan(plan).domain).toBe(domain)
    const savedSection = plan.slides[0]!.domainSection
    plan.slides[0]!.domainSection = 'question'
    expect(() => parsePresentationPlan(plan)).toThrow('domain_section')
    plan.slides[0]!.domainSection = savedSection
    plan.slides.pop()
    expect(() => parsePresentationPlan(plan)).toThrow('domain_section')
  },
)
it('keeps legacy domains without workflows and returns clones', () => {
  for (const domain of [undefined, 'unknown', 'pitch', 'report', 'training', 'research', 'sales'])
    expect(presentationProfessionalWorkflow(domain)).toBeUndefined()
  const original = presentationProfessionalWorkflow('science')!,
    copy = presentationProfessionalWorkflow('science')!
  expect(copy).toEqual(original)
  expect(copy).not.toBe(original)
  expect(copy.reviewSteps).not.toBe(original.reviewSteps)
})

it('does not leak nested workflow objects or instructions across reads', () => {
  const copy = presentationProfessionalWorkflow('science')!
  ;(copy.sourcePriority as string[]).push('污染')
  ;(copy.reviewSteps[0]!.tools as string[]).push('fake_tool')
  expect(presentationProfessionalWorkflow('science')!.sourcePriority).not.toContain('污染')
  expect(presentationProfessionalWorkflow('science')!.reviewSteps[0]!.tools).not.toContain(
    'fake_tool',
  )
})
