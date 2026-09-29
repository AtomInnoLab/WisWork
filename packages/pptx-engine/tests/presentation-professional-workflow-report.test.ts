import { expect, it, vi } from 'vitest'
import * as planModule from '../src/presentation-plan'
import { benchmarkPlan } from './fixtures/presentation-plan'
import { benchmarkDeck } from './fixtures/presentation-benchmark'
import { presentationPlanClaims, presentationProfessionalWorkflow } from '../src/presentation-plan'
import {
  buildPresentationDeliveryReport,
  parsePresentationDeliveryReport,
  presentationDeliveryMarkdown,
} from '../src/presentation-delivery-report'

function fixture(professional = true) {
  const plan = benchmarkPlan(),
    deck = benchmarkDeck()
  if (professional) {
    plan.domain = 'science'
    const sections = [
      'research_question',
      'methods_and_sample',
      'results_and_data',
      'scope_and_limitations',
      'research_references',
    ] as const
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
it('requires declared context for every professional workflow claim and rejects omitted issues', async () => {
  const value = fixture(),
    report = await buildPresentationDeliveryReport(value)
  expect(
    report.pages[0]!.issues.find((issue) => issue.code === 'professional_context_missing')
      ?.category,
  ).toBe('unverifiable')
  const forged = structuredClone(report)
  forged.pages[0]!.issues = forged.pages[0]!.issues.filter(
    (issue) => issue.code !== 'professional_context_missing',
  )
  expect(() => parsePresentationDeliveryReport(forged)).toThrow()
  const markdown = presentationDeliveryMarkdown(report)
  expect(report.professionalWorkflow).toEqual(presentationProfessionalWorkflow('science'))
  for (const professionalWorkflow of [
    undefined,
    { ...report.professionalWorkflow!, disclosure: 'forged' },
  ]) {
    expect(() => parsePresentationDeliveryReport({ ...report, professionalWorkflow })).toThrow()
  }
  expect(markdown).toContain('manualChecks')
  expect(markdown).toContain('NOT VERIFIED')
  expect(markdown).toContain(
    presentationProfessionalWorkflow('science')!.reviewSteps[0]!.id.replaceAll('_', '&#95;'),
  )
})
it('accepts supporting contexts from another domain and retains partial warnings', async () => {
  const value = fixture()
  value.plan.claims[0]!.professionalContext = { domain: 'finance' }
  const report = await buildPresentationDeliveryReport(value)
  expect(report.pages[0]!.issues.map((issue) => issue.code)).not.toContain(
    'professional_context_missing',
  )
  expect(report.pages[0]!.issues.map((issue) => issue.code)).toContain(
    'professional_context_incomplete',
  )
})
it('keeps legacy reports unchanged and explained missing issues open and stale after claim changes', async () => {
  const legacy = await buildPresentationDeliveryReport(fixture(false))
  expect(legacy.pages.flatMap((page) => page.issues).map((issue) => issue.code)).not.toContain(
    'professional_context_missing',
  )
  expect(presentationDeliveryMarkdown(legacy)).not.toContain('Professional workflow')
  expect(legacy).not.toHaveProperty('professionalWorkflow')
  expect(() =>
    parsePresentationDeliveryReport({ ...legacy, professionalWorkflow: undefined }),
  ).toThrow()
  const value = fixture(),
    first = await buildPresentationDeliveryReport(value)
  const issue = first.pages[0]!.issues.find(
    (issue) => issue.code === 'professional_context_missing',
  )!
  const ledger = {
    ...value.issueLedger,
    revision: 1,
    actions: [
      {
        actionId: 'explain',
        issueId: issue.id,
        issueDigest: issue.digest,
        state: 'explained' as const,
        note: 'Human explanation',
        sequence: 1,
        createdAt: '2026-09-29T00:00:00.000Z',
      },
    ],
  }
  const explained = await buildPresentationDeliveryReport({ ...value, issueLedger: ledger })
  expect(explained.pages[0]!.issues.find((item) => item.id === issue.id)?.disposition.state).toBe(
    'explained',
  )
  value.plan.claims[0]!.statement += ' Updated scope'
  value.deck.claims = presentationPlanClaims(value.plan)
  const changed = await buildPresentationDeliveryReport({ ...value, issueLedger: ledger })
  expect(changed.pages[0]!.issues.find((item) => item.id === issue.id)?.disposition).toMatchObject({
    state: 'open',
    stale: true,
  })
})
it('includes workflow rules in the issue digest', async () => {
  const value = fixture(),
    first = await buildPresentationDeliveryReport(value)
  const original = presentationProfessionalWorkflow
  const spy = vi
    .spyOn(planModule, 'presentationProfessionalWorkflow')
    .mockImplementation((domain) => {
      const workflow = original(domain)
      return workflow
        ? { ...workflow, disclosure: workflow.disclosure + ' Additional review rule' }
        : undefined
    })
  try {
    const changed = await buildPresentationDeliveryReport(value)
    const before = first.pages[0]!.issues.find(
      (issue) => issue.code === 'professional_context_missing',
    )!
    expect(changed.pages[0]!.issues.find((issue) => issue.id === before.id)?.digest).not.toBe(
      before.digest,
    )
  } finally {
    spy.mockRestore()
  }
})
