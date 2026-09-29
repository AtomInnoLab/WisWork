import { expect, it } from 'vitest'
import * as planModule from '../src/presentation-plan'
import { benchmarkPlan } from './fixtures/presentation-plan'
import { benchmarkDeck } from './fixtures/presentation-benchmark'
import { presentationPlanClaims } from '../src/presentation-plan'
import {
  buildPresentationDeliveryReport,
  parsePresentationDeliveryReport,
  presentationDeliveryMarkdown,
} from '../src/presentation-delivery-report'

function fixture(domain: 'pitch' | 'report' | 'training' | 'research' | 'sales') {
  const plan = benchmarkPlan(),
    deck = benchmarkDeck()
  {
    plan.domain = domain
    const sections = planModule.PRESENTATION_DOMAIN_PROFILES[domain].sections
    plan.slides.forEach((slide, index) => (slide.domainSection = sections[index % sections.length]))
  }
  deck.claims = presentationPlanClaims(plan)
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
it.each(['pitch', 'report', 'training', 'research', 'sales'] as const)(
  'exports complete %s domain guidance and strictly rejects alteration',
  async (domain) => {
    const report = await buildPresentationDeliveryReport(fixture(domain))
    const guide = planModule.presentationDomainWorkflow(domain)!
    expect(guide).toBeDefined()
    expect(report.domainWorkflow).toEqual(guide)
    const md = presentationDeliveryMarkdown(report)
    expect(md).toContain('manualChecks')
    expect(md).toContain('NOT VERIFIED')
    const escaped = (value: string) =>
      Array.from(value, (c) =>
        /[A-Za-z0-9 \u3400-\u9fff]/u.test(c) ? c : `&#${c.codePointAt(0)};`,
      ).join('')
    for (const text of [
      ...guide.sections.map((s) => s.instruction),
      ...guide.reviewSteps.map((s) => s.instruction),
      ...guide.manualChecks,
      guide.disclosure,
    ])
      expect(md).toContain(escaped(text))
    expect(JSON.parse(JSON.stringify(report)).domainWorkflow).toEqual(guide)

    for (const changed of [
      { ...guide, domain: 'science' },
      { ...guide, version: 2 },
      { ...guide, disclosure: 'forged' },
      { ...guide, extra: true },
      {
        ...guide,
        sections: guide.sections.map((section, index) =>
          index === 0 ? { ...section, instruction: 'forged' } : section,
        ),
      },
      {
        ...guide,
        reviewSteps: guide.reviewSteps.map((step, index) =>
          index === 0 ? { ...step, tools: ['fake_tool'] } : step,
        ),
      },
    ])
      expect(() =>
        parsePresentationDeliveryReport({ ...report, domainWorkflow: changed }),
      ).toThrow()
    const legacy = { ...report }
    delete legacy.domainWorkflow
    expect(parsePresentationDeliveryReport(legacy)).toEqual(legacy)
  },
)
it('rejects a workflow from a different known domain and any workflow on an unclassified plan', async () => {
  const value = fixture('pitch')
  const report = await buildPresentationDeliveryReport(value)
  expect(() =>
    parsePresentationDeliveryReport({
      ...report,
      domainWorkflow: planModule.presentationDomainWorkflow('sales'),
    }),
  ).toThrow()
  expect(() => parsePresentationDeliveryReport({ ...report, domainWorkflow: undefined })).toThrow()
  const generic = fixture('pitch')
  delete generic.plan.domain
  generic.plan.slides.forEach((s) => delete s.domainSection)
  const old = await buildPresentationDeliveryReport(generic)
  expect(old).not.toHaveProperty('domainWorkflow')
  expect(() =>
    parsePresentationDeliveryReport({ ...old, domainWorkflow: report.domainWorkflow }),
  ).toThrow()
})
it('keeps professional workflow contract and rejects domain guidance on a professional report', async () => {
  const value = fixture('pitch')
  value.plan.domain = 'science'
  value.plan.slides.forEach(
    (slide, index) =>
      (slide.domainSection = planModule.PRESENTATION_DOMAIN_PROFILES.science.sections[index % 5]),
  )
  const report = await buildPresentationDeliveryReport(value)
  expect(report).not.toHaveProperty('domainWorkflow')
  expect(report.professionalWorkflow).toEqual(
    planModule.presentationProfessionalWorkflow('science'),
  )
  expect(() =>
    parsePresentationDeliveryReport({
      ...report,
      domainWorkflow: planModule.presentationDomainWorkflow('pitch'),
    }),
  ).toThrow()
})
