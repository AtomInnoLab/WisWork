import { describe, expect, it } from 'vitest'
import type { SlideIRElement } from '../src/presentation'
import { presentationPlanClaims } from '../src/presentation-plan'
import {
  checkPresentationPageContent,
  parsePresentationPageContentCheck,
} from '../src/presentation-content-check'
import { benchmarkPlan } from './fixtures/presentation-plan'
import { benchmarkDeck } from './fixtures/presentation-benchmark'

const box = { id: 'text', x: 1, y: 1, w: 5, h: 1 }
function fixture() {
  const plan = benchmarkPlan()
  plan.claims[0]!.statement = 'Revenue grew 20%'
  plan.claims[0]!.type = 'quote'
  plan.sources[0]!.excerpt = 'Report: Revenue grew 20% in 2025.'
  const deck = benchmarkDeck()
  deck.claims = presentationPlanClaims(plan)
  deck.slides[0]!.elements = [{ ...box, kind: 'text', text: 'Revenue\n grew   20%' }]
  const pageId = deck.slides[0]!.id
  return { plan, deck, pageId }
}
const conservative = {
  content: 'needs_review',
  sources: 'not_verified',
  calculations: 'not_verified',
  timeliness: 'not_verified',
  host: 'not_checked',
}

describe('presentation page content precheck', () => {
  it('normalizes whitespace, stays conservative and leaves inputs unchanged', () => {
    const { plan, deck, pageId } = fixture()
    const before = structuredClone({ plan, deck })
    const report = checkPresentationPageContent(plan, deck, pageId)
    expect(report).toEqual({
      version: 1,
      pageId,
      claimIds: ['source-1'],
      findings: [],
      checks: conservative,
    })
    expect({ plan, deck }).toEqual(before)
    expect(parsePresentationPageContentCheck(report)).toEqual(report)
    expect(parsePresentationPageContentCheck(report)).not.toBe(report)
  })
  it.each([
    [undefined, 'source_as_of_missing'],
    [' \n\t ', 'source_as_of_missing'],
    ['FY 2024', 'source_as_of_differs'],
    ['FY  2025', undefined],
  ])('compares source asOf %j as normalized metadata only', (asOf, code) => {
    const { plan, deck, pageId } = fixture()
    plan.claims[0]!.asOf = ' FY\n2025 '
    if (asOf !== undefined) plan.sources[0]!.asOf = asOf
    deck.claims = presentationPlanClaims(plan)
    const before = structuredClone({ plan, deck })
    const report = checkPresentationPageContent(plan, deck, pageId)
    expect(report.findings).toEqual(code ? [{ code, claimId: 'source-1', sourceId: 'source' }] : [])
    expect(report.checks).toEqual(conservative)
    expect({ plan, deck }).toEqual(before)
    expect(parsePresentationPageContentCheck(report)).toEqual(report)
  })
  it.each([
    ['2026-02-28', '2026-03-01', 'source_as_of_earlier'],
    ['2024-02-29', '2024-03-01', 'source_as_of_earlier'],
    ['2025-03-01', '2025-02-28', 'source_as_of_differs'],
    ['2025-02-29', '2025-03-01', 'source_as_of_differs'],
    ['FY 2024', 'FY 2025', 'source_as_of_differs'],
  ])('only orders valid exact calendar dates', (sourceDate, claimDate, code) => {
    const { plan, deck, pageId } = fixture()
    plan.claims[0]!.asOf = claimDate
    plan.sources[0]!.asOf = sourceDate
    deck.claims = presentationPlanClaims(plan)
    expect(checkPresentationPageContent(plan, deck, pageId).findings).toContainEqual({
      code,
      claimId: 'source-1',
      sourceId: 'source',
    })
  })
  it.each([undefined, ' \n '])('does not infer an unspecified claim asOf %j', (asOf) => {
    const { plan, deck, pageId } = fixture()
    if (asOf !== undefined) plan.claims[0]!.asOf = asOf
    plan.sources[0]!.asOf = 'FY 2024'
    deck.claims = presentationPlanClaims(plan)
    expect(checkPresentationPageContent(plan, deck, pageId).findings).toEqual([])
  })
  it('does not report temporal findings for claims on other pages', () => {
    const { plan, deck, pageId } = fixture()
    plan.claims[0]!.asOf = 'FY 2025'
    plan.slides[0]!.claimIds = []
    deck.slides[0]!.claimIds = []
    deck.claims = presentationPlanClaims(plan)
    expect(checkPresentationPageContent(plan, deck, pageId).findings).toEqual([])
  })
  it('matches individual table cells, chart categories and series names', () => {
    for (const element of [
      { ...box, kind: 'table', rows: [['Revenue grew 20%']] },
      {
        ...box,
        kind: 'chart',
        chartType: 'bar',
        categories: ['Revenue grew 20%'],
        series: [{ name: 'series', values: [1] }],
      },
      {
        ...box,
        kind: 'chart',
        chartType: 'bar',
        categories: ['category'],
        series: [
          { name: 'Revenue grew 20%', values: [1] },
          { name: 'Other series', values: [2] },
        ],
      },
    ] satisfies SlideIRElement[]) {
      const { plan, deck, pageId } = fixture()
      deck.slides[0]!.elements = [element]
      expect(checkPresentationPageContent(plan, deck, pageId).findings).toEqual([])
    }
  })
  it('ignores hidden single-series names and pie category labels', () => {
    for (const chartType of ['bar', 'line', 'pie'] as const) {
      const { plan, deck, pageId } = fixture()
      deck.slides[0]!.elements = [
        {
          ...box,
          kind: 'chart',
          chartType,
          categories: [chartType === 'pie' ? 'Revenue grew 20%' : 'Category'],
          series: [{ name: 'Revenue grew 20%', values: [1] }],
        },
      ]
      expect(checkPresentationPageContent(plan, deck, pageId).findings).toEqual([
        { code: 'claim_text_not_found', claimId: 'source-1' },
      ])
    }
  })
  it('does not join elements, cells or labels, or count metadata and notes', () => {
    for (const elements of [
      [
        { ...box, kind: 'text', text: 'Revenue grew' },
        { ...box, id: 'second', kind: 'text', text: '20%' },
      ],
      [{ ...box, kind: 'table', rows: [['Revenue grew', '20%']] }],
      [
        {
          ...box,
          kind: 'chart',
          chartType: 'bar',
          categories: ['Revenue grew'],
          series: [{ name: '20%', values: [1] }],
        },
      ],
    ] satisfies SlideIRElement[][]) {
      const { plan, deck, pageId } = fixture()
      deck.slides[0]!.elements = elements
      deck.slides[0]!.notes = 'Revenue grew 20%'
      deck.slides[0]!.title = plan.slides[0]!.title = 'Revenue grew 20%'
      expect(checkPresentationPageContent(plan, deck, pageId).findings).toEqual([
        { code: 'claim_text_not_found', claimId: 'source-1' },
      ])
    }
  })
  it('reports trimmed missing evidence and locators, and quote disagreement', () => {
    const { plan, deck, pageId } = fixture()
    plan.sources[0]!.excerpt = ' \n '
    plan.sources[0]!.locator = ' \t '
    deck.claims = presentationPlanClaims(plan)
    expect(checkPresentationPageContent(plan, deck, pageId).findings).toEqual([
      { code: 'source_excerpt_missing', claimId: 'source-1', sourceId: 'source' },
      { code: 'source_locator_missing', claimId: 'source-1', sourceId: 'source' },
    ])
    plan.sources[0]!.excerpt = 'Revenue fell 20%'
    expect(checkPresentationPageContent(plan, deck, pageId).findings).toContainEqual({
      code: 'quote_not_in_excerpt',
      claimId: 'source-1',
      sourceId: 'source',
    })
  })
  it('never evaluates formulas, and checks only claims on the requested page', () => {
    const { plan, deck, pageId } = fixture()
    plan.claims[0]!.type = 'calculation'
    plan.claims[0]!.calculation = { formula: 'throw new Error("executed")', inputs: ['1'] }
    expect(checkPresentationPageContent(plan, deck, pageId).findings).toEqual([
      { code: 'calculation_not_reproduced', claimId: 'source-1' },
    ])
    plan.slides[0]!.claimIds = []
    deck.slides[0]!.claimIds = []
    expect(checkPresentationPageContent(plan, deck, pageId).claimIds).toEqual([])
    expect(checkPresentationPageContent(plan, deck, pageId).findings).toEqual([])
  })
  it('rejects unknown pages, invalid plans/decks and mismatched bindings', () => {
    const { plan, deck, pageId } = fixture()
    expect(() => checkPresentationPageContent(plan, deck, 'missing')).toThrow('not_found')
    expect(() =>
      checkPresentationPageContent({ ...plan, version: 2 } as never, deck, pageId),
    ).toThrow('presentation_plan_invalid:')
    expect(() =>
      checkPresentationPageContent(plan, { ...deck, version: 2 } as never, pageId),
    ).toThrow('presentation_invalid:')
    expect(() => checkPresentationPageContent(plan, { ...deck, id: 'different' }, pageId)).toThrow(
      'presentation_plan_mismatch:',
    )
  })
  it('supports the maximum 32 claims with three missing sources and calculations', () => {
    const { plan, deck, pageId } = fixture()
    plan.sources = Array.from({ length: 3 }, (_, i) => ({
      id: `s${i}`,
      title: 'Source',
      uri: 'unfetched',
      excerpt: ' ',
    }))
    plan.claims = Array.from({ length: 32 }, (_, i) => ({
      ...plan.claims[0]!,
      id: `c${i}`,
      statement: `Missing ${i}`,
      asOf: 'FY 2025',
      type: 'calculation' as const,
      sourceIds: ['s0', 's1', 's2'],
      calculation: { formula: '1+1', inputs: ['1'] },
    }))
    plan.slides.forEach((slide) => {
      slide.claimIds = []
    })
    plan.slides[0]!.claimIds = plan.claims.map((claim) => claim.id)
    deck.slides.forEach((slide, i) => {
      slide.claimIds = [...plan.slides[i]!.claimIds]
    })
    deck.claims = presentationPlanClaims(plan)
    const report = checkPresentationPageContent(plan, deck, pageId)
    expect(report.findings).toHaveLength(352)
    expect(parsePresentationPageContentCheck(report)).toEqual(report)
  })
  it('strictly rejects forged reports and invalid finding attribution', () => {
    const { plan, deck, pageId } = fixture()
    const base = checkPresentationPageContent(plan, deck, pageId)
    for (const change of [
      { extra: true },
      { version: 2 },
      { pageId: '' },
      { claimIds: ['source-1', 'source-1'] },
      { claimIds: Array.from({ length: 33 }, (_, i) => `c${i}`) },
      { checks: { ...conservative, content: 'passed' } },
      { checks: { ...conservative, extra: true } },
      { findings: [{ code: 'unknown', claimId: 'source-1' }] },
      ...['source_as_of_missing', 'source_as_of_differs'].flatMap((code) => [
        { findings: [{ code, claimId: 'source-1' }] },
        { findings: [{ code, sourceId: 'source' }] },
        { findings: [{ code, claimId: 'other', sourceId: 'source' }] },
        { findings: [{ code, claimId: 'source-1', sourceId: '' }] },
        { findings: [{ code, claimId: 'source-1', sourceId: 'source', extra: true }] },
        { findings: Array(2).fill({ code, claimId: 'source-1', sourceId: 'source' }) },
      ]),
      { findings: [{ code: 'claim_text_not_found', claimId: 'other' }] },
      { findings: [{ code: 'source_excerpt_missing', claimId: 'source-1' }] },
      { findings: [{ code: 'claim_text_not_found', claimId: 'source-1', sourceId: 'source' }] },
      { findings: [{ code: 'claim_text_not_found', claimId: 'source-1', text: 'untrusted' }] },
      { findings: Array(2).fill({ code: 'claim_text_not_found', claimId: 'source-1' }) },
      {
        findings: Array.from({ length: 353 }, (_, i) => ({
          code: 'source_excerpt_missing',
          claimId: 'source-1',
          sourceId: `s${i}`,
        })),
      },
    ])
      expect(() => parsePresentationPageContentCheck({ ...base, ...change })).toThrow(
        'presentation_content_check_invalid:',
      )
  })
})
