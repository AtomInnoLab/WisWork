import { expect, it } from 'vitest'
import { benchmarkPlan } from './fixtures/presentation-plan'
import { benchmarkDeck } from './fixtures/presentation-benchmark'
import { presentationPlanClaims } from '../src/presentation-plan'
import {
  checkPresentationPageContent,
  parsePresentationPageContentCheck,
} from '../src/presentation-content-check'
import {
  buildPresentationDeliveryReport,
  parsePresentationDeliveryReport,
  presentationDeliveryMarkdown,
} from '../src/presentation-delivery-report'
function fixture(declared = true) {
  const plan = benchmarkPlan(),
    deck = benchmarkDeck(),
    slide = plan.slides[0]!
  plan.sources[0]!.excerpt = '10'
  if (declared)
    slide.chartData = [
      {
        elementId: 'chart',
        categories: ['A'],
        series: [
          {
            name: 'Series',
            points: [
              {
                value: 10,
                claimId: 'source-1',
                basis: { kind: 'source', sourceId: 'source', excerptOffset: 0, excerptText: '10' },
              },
            ],
          },
        ],
      },
    ]
  deck.claims = presentationPlanClaims(plan)
  deck.slides[0]!.elements = [
    {
      id: 'chart',
      kind: 'chart',
      chartType: 'bar',
      x: 1,
      y: 1,
      w: 4,
      h: 3,
      categories: ['A'],
      series: [{ name: 'Series', values: [11] }],
    },
  ]
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
    pageStates: plan.slides.map((page) => ({ pageId: page.id, state: 'pending' as const })),
    issueLedger: {
      version: 1 as const,
      projectId: metadata.projectId,
      documentId: metadata.documentId,
      requestId: metadata.requestId,
      inputDigest: metadata.inputDigest,
      planDigest: metadata.planDigest,
      revision: 0,
      actions: [],
    },
  }
}
it('requires check for declared chart data and rebuilds checks to reject omitted findings', async () => {
  const value = fixture(),
    pageId = value.plan.slides[0]!.id
  const content = checkPresentationPageContent(value.plan, value.deck, pageId)
  expect(content.chartData?.charts[0]!.findings.map((finding) => finding.code)).toContain(
    'chart_data_value_mismatch',
  )
  expect(() =>
    parsePresentationPageContentCheck({ ...content, chartData: undefined }, value.plan),
  ).toThrow()
  const foreign = structuredClone(content)
  foreign.chartData!.charts[0]!.findings[0]!.claimIds = ['foreign-claim']
  expect(() => parsePresentationPageContentCheck(foreign)).toThrow(
    'presentation_content_check_invalid',
  )
  const report = await buildPresentationDeliveryReport(value)
  expect(
    report.pages[0]!.issues.find((issue) => issue.code === 'chart_data_value_mismatch')?.category,
  ).toBe('unverifiable')
  const forged = structuredClone(report)
  forged.pages[0]!.chartData!.charts[0]!.findings = []
  expect(() => parsePresentationDeliveryReport(forged)).toThrow(
    'presentation_delivery_report_invalid',
  )
  delete forged.pages[0]!.chartData
  expect(() => parsePresentationDeliveryReport(forged)).toThrow()
  expect(presentationDeliveryMarkdown(report)).toContain('sourceTruth NOT VERIFIED')
})
it('keeps unbound charts visible without invented claim issues and accepts old reports', async () => {
  const value = fixture(false),
    report = await buildPresentationDeliveryReport(value)
  expect(report.pages[0]!.chartData?.charts[0]!.findings).toContainEqual({
    code: 'chart_data_unbound',
    claimIds: [],
  })
  expect(report.pages[0]!.issues.map((issue) => issue.code)).not.toContain('chart_data_unbound')
  delete report.pages[0]!.chartData
  expect(() => parsePresentationDeliveryReport(report)).not.toThrow()
})
it('keeps matched data unverified and makes explained chart gaps stale when actual data changes', async () => {
  const value = fixture(),
    first = await buildPresentationDeliveryReport(value)
  const issue = first.pages[0]!.issues.find((issue) => issue.code === 'chart_data_value_mismatch')!
  const issueLedger = {
    ...value.issueLedger,
    revision: 1,
    actions: [
      {
        actionId: 'explain',
        issueId: issue.id,
        issueDigest: issue.digest,
        state: 'explained' as const,
        note: 'Requires human data review',
        sequence: 1,
        createdAt: '2026-09-29T00:00:00.000Z',
      },
    ],
  }
  const chart = value.deck.slides[0]!.elements[0]!
  if (chart.kind !== 'chart') throw new Error('fixture')
  chart.series[0]!.values[0] = 12
  const changed = await buildPresentationDeliveryReport({ ...value, issueLedger })
  const legacyIssue = first.pages[0]!.issues.find((item) => item.code === 'claim_text_not_found')!
  expect(changed.pages[0]!.issues.find((item) => item.id === legacyIssue.id)?.digest).toBe(
    legacyIssue.digest,
  )
  expect(changed.pages[0]!.issues.find((item) => item.id === issue.id)?.disposition).toMatchObject({
    state: 'open',
    stale: true,
  })
  chart.series[0]!.values[0] = 10
  const matched = await buildPresentationDeliveryReport(value)
  expect(matched.pages[0]!.chartData?.charts[0]!.findings).toEqual([])
  expect(matched.pages[0]!.chartData?.checks.sourceTruth).toBe('not_verified')
})
it('reports declared charts missing from actual elements and keeps plain content shape exact', async () => {
  const value = fixture(),
    pageId = value.plan.slides[0]!.id
  value.deck.slides[0]!.elements = [
    { id: 'plain', kind: 'text', x: 1, y: 1, w: 4, h: 1, text: 'Unrelated body' },
  ]
  const report = await buildPresentationDeliveryReport(value)
  expect(report.pages[0]!.issues.map((issue) => issue.code)).toContain('chart_data_missing')
  const content = checkPresentationPageContent(value.plan, value.deck, pageId)
  const absent = structuredClone(content)
  delete absent.chartData
  expect(() => parsePresentationPageContentCheck(absent, value.plan)).toThrow()
  delete value.plan.slides[0]!.chartData
  const plain = checkPresentationPageContent(value.plan, value.deck, pageId)
  expect(plain).not.toHaveProperty('chartData')
  expect(parsePresentationPageContentCheck(plain)).toEqual(plain)
})
