import { expect, it } from 'vitest'
import {
  checkPresentationChartData,
  parsePresentationChartDataCheck,
} from '../src/presentation-chart-data'
import { benchmarkPlan } from './fixtures/presentation-plan'
import { parsePresentationPlan } from '../src/presentation-plan'
function fixture() {
  const plan = benchmarkPlan(),
    page = plan.slides.find((p) => p.claimIds.includes(plan.claims[0]!.id))!
  plan.sources[0]!.excerpt = '12'
  page.chartData = [
    {
      elementId: 'chart',
      categories: ['A'],
      series: [
        {
          name: 'S',
          points: [
            {
              value: 12,
              claimId: plan.claims[0]!.id,
              basis: {
                kind: 'source',
                sourceId: plan.sources[0]!.id,
                excerptOffset: 0,
                excerptText: '12',
              },
            },
          ],
        },
      ],
    },
  ]
  return {
    plan,
    page,
    actual: [{ id: 'chart', categories: ['A'], series: [{ name: 'S', values: [12] }] }],
  }
}
it('checks literal frozen declared data without certifying truth', () => {
  const f = fixture()
  parsePresentationPlan(f.plan)
  const result = checkPresentationChartData(f.plan, f.page.id, f.actual)
  expect(result.charts[0]!.findings).toEqual([])
  expect(result.checks.sourceTruth).toBe('not_verified')
  expect(parsePresentationChartDataCheck(result, f.plan)).toEqual(result)
})
it('reports changed values, absent charts and unbound actual charts', () => {
  const f = fixture()
  f.actual[0]!.series[0]!.values = [13]
  expect(
    checkPresentationChartData(f.plan, f.page.id, f.actual).charts[0]!.findings,
  ).toContainEqual({ code: 'chart_data_value_mismatch', claimIds: [f.plan.claims[0]!.id] })
  expect(checkPresentationChartData(f.plan, f.page.id, []).charts[0]!.findings[0]!.code).toBe(
    'chart_data_missing',
  )
  f.actual[0]!.id = 'other'
  expect(checkPresentationChartData(f.plan, f.page.id, f.actual).charts).toHaveLength(2)
})
it.each([' 12', '12 ', '1,200', '12%', '十二', '12\f'])(
  'does not normalize numeric source text %j',
  (literal) => {
    const f = fixture(),
      basis = f.page.chartData![0]!.series[0]!.points[0]!.basis
    if (basis.kind !== 'source') throw Error('fixture')
    basis.excerptText = literal
    const result = checkPresentationChartData(f.plan, f.page.id, f.actual)
    expect(result.charts[0]!.findings.map((v) => v.code)).toContain(
      'chart_data_source_basis_mismatch',
    )
    expect(
      parsePresentationPlan(f.plan).slides.find((p) => p.id === f.page.id)!.chartData![0]!
        .series[0]!.points[0]!.basis,
    ).toEqual(basis)
  },
)
it.each(['+12.0', '1.2e1', '.12E+2'])('accepts exact numeric source literal %s', (literal) => {
  const f = fixture()
  f.plan.sources[0]!.excerpt = literal
  const basis = f.page.chartData![0]!.series[0]!.points[0]!.basis
  if (basis.kind === 'source') basis.excerptText = literal
  expect(checkPresentationChartData(f.plan, f.page.id, f.actual).charts[0]!.findings).toEqual([])
})
it('preserves absolute UTF16 source offsets and reports shape/labels rather than assuming equivalence', () => {
  const f = fixture()
  f.plan.sources[0]!.excerpt = '😀12'
  const p = f.page.chartData![0]!.series[0]!.points[0]!
  if (p.basis.kind === 'source') p.basis.excerptOffset = 2
  expect(checkPresentationChartData(f.plan, f.page.id, f.actual).charts[0]!.findings).toEqual([])
  if (p.basis.kind === 'source') p.basis.excerptOffset = 1
  f.actual[0]!.categories = ['B']
  f.actual[0]!.series[0]!.name = 'other'
  expect(
    checkPresentationChartData(f.plan, f.page.id, f.actual).charts[0]!.findings.map((f) => f.code),
  ).toEqual(['chart_data_shape_mismatch', 'chart_data_source_basis_mismatch'])
})
it('reuses bounded arithmetic and keeps explicit currency/unit mismatches independent', () => {
  const f = fixture(),
    claim = f.plan.claims[0]!,
    p = f.page.chartData![0]!.series[0]!.points[0]!
  claim.type = 'calculation'
  claim.calculation = {
    formula: 'a',
    inputs: ['12'],
    unit: '万元',
    currency: 'CNY',
    reproduction: {
      bindings: [{ name: 'a', inputIndex: 0, value: 12, sourceId: claim.sourceIds[0]! }],
      expected: 12,
    },
  }
  claim.professionalContext = { domain: 'finance', unit: '元', currency: 'USD' }
  p.basis = { kind: 'calculation' }
  f.page.chartData![0]!.unit = '万元'
  f.page.chartData![0]!.currency = 'CNY'
  parsePresentationPlan(f.plan)
  const report = checkPresentationChartData(f.plan, f.page.id, f.actual)
  expect(report.charts[0]!.findings.map((f) => f.code)).toEqual([
    'chart_data_unit_mismatch',
    'chart_data_currency_mismatch',
  ])
  claim.calculation.reproduction!.expected = 13
  expect(
    checkPresentationChartData(f.plan, f.page.id, f.actual).charts[0]!.findings.map((f) => f.code),
  ).toContain('chart_data_calculation_not_reproduced')
})
it('rejects forged findings/checks/claim identities and reconstructs missing declared charts', () => {
  const f = fixture()
  f.actual[0]!.series[0]!.values = [13]
  const good = checkPresentationChartData(f.plan, f.page.id, f.actual)
  for (const changed of [
    { ...good, checks: { ...good.checks, data: 'passed' } },
    { ...good, charts: good.charts.map((c) => ({ ...c, findings: [] })) },
    {
      ...good,
      charts: good.charts.map((c) => ({
        ...c,
        findings: c.findings.map((v) => ({ ...v, claimIds: ['fake'] })),
      })),
    },
    { ...good, charts: [] },
    { ...good, pageId: 'fake' },
    { ...good, charts: [...good.charts, ...good.charts] },
    { ...good, unknown: true },
  ])
    expect(() => parsePresentationChartDataCheck(changed, f.plan)).toThrow()
  const duplicate = structuredClone(good)
  duplicate.charts[0]!.findings.push(duplicate.charts[0]!.findings[0]!)
  expect(() => parsePresentationChartDataCheck(duplicate)).toThrow()
})
it('enforces finite values, identity, dimension, bounded source references and old plan absence', () => {
  const old = benchmarkPlan()
  expect(parsePresentationPlan(old)).toEqual(old)
  expect(old.slides[0]).not.toHaveProperty('chartData')
  for (const change of [
    (f: ReturnType<typeof fixture>) => {
      f.page.chartData!.push(f.page.chartData![0]!)
    },
    (f: ReturnType<typeof fixture>) => {
      f.page.chartData![0]!.series[0]!.points[0]!.value = Infinity
    },
    (f: ReturnType<typeof fixture>) => {
      f.page.chartData![0]!.series[0]!.points[0]!.claimId = 'fake'
    },
    (f: ReturnType<typeof fixture>) => {
      f.page.chartData![0]!.series[0]!.points = []
    },
    (f: ReturnType<typeof fixture>) => {
      f.page.chartData![0]!.series[0]!.points[0]!.basis = { kind: 'calculation' }
    },
    (f: ReturnType<typeof fixture>) => {
      const b = f.page.chartData![0]!.series[0]!.points[0]!.basis
      if (b.kind === 'source') {
        b.sourceId = 'fake'
      }
    },
    (f: ReturnType<typeof fixture>) => {
      const b = f.page.chartData![0]!.series[0]!.points[0]!.basis
      if (b.kind === 'source') {
        b.excerptOffset = 0.5
      }
    },
  ]) {
    const f = fixture()
    change(f)
    expect(() => parsePresentationPlan(f.plan)).toThrow()
  }
  const f = fixture()
  expect(() =>
    checkPresentationChartData(f.plan, f.page.id, [{ ...f.actual[0]!, kind: 'text' } as never]),
  ).toThrow()
  expect(() => checkPresentationChartData(f.plan, f.page.id, [...f.actual, ...f.actual])).toThrow()
})
it('bounds chart union and input sizes without inventing claim IDs for unbound charts', () => {
  const f = fixture()
  f.page.chartData = Array.from({ length: 128 }, (_, i) => ({
    ...f.page.chartData![0]!,
    elementId: 'declared' + i,
  }))
  const actual = Array.from({ length: 128 }, (_, i) => ({ ...f.actual[0]!, id: 'actual' + i }))
  const result = checkPresentationChartData(f.plan, f.page.id, actual)
  expect(result.charts).toHaveLength(256)
  expect(
    result.charts.filter((c) => c.actual).every((c) => c.findings[0]!.claimIds.length === 0),
  ).toBe(true)
  expect(parsePresentationChartDataCheck(result, f.plan)).toEqual(result)
  expect(() =>
    checkPresentationChartData(f.plan, f.page.id, [...actual, { ...actual[0]!, id: 'overflow' }]),
  ).toThrow()
  expect(() =>
    parsePresentationChartDataCheck({ ...result, padding: 'x'.repeat(8 * 1024 * 1024) }),
  ).toThrow()
})
it.each([
  ['212', 1, '12', 12],
  ['-12', 1, '12', 12],
  ['12e3', 0, '12', 12],
  ['1,200', 2, '200', 200],
  ['12.0', 0, '12', 12],
])('flags partial numeric tokens in %s', (excerpt, offset, literal, value) => {
  const f = fixture()
  f.plan.sources[0]!.excerpt = excerpt as string
  const p = f.page.chartData![0]!.series[0]!.points[0]!
  p.value = value as number
  f.actual[0]!.series[0]!.values = [value as number]
  if (p.basis.kind === 'source') {
    p.basis.excerptOffset = offset as number
    p.basis.excerptText = literal as string
  }
  expect(
    checkPresentationChartData(f.plan, f.page.id, f.actual).charts[0]!.findings.map((f) => f.code),
  ).toContain('chart_data_source_basis_mismatch')
})
it('flags omitted known unit/currency without inferring labels when both are unknown', () => {
  const f = fixture()
  f.plan.claims[0]!.professionalContext = { domain: 'finance', unit: '元', currency: 'CNY' }
  expect(
    checkPresentationChartData(f.plan, f.page.id, f.actual).charts[0]!.findings.map((f) => f.code),
  ).toEqual(['chart_data_unit_mismatch', 'chart_data_currency_mismatch'])
  delete f.plan.claims[0]!.professionalContext
  expect(checkPresentationChartData(f.plan, f.page.id, f.actual).charts[0]!.findings).toEqual([])
})
it('rejects a final newline as part of a numeric literal rather than regex end-anchor normalization', () => {
  const f = fixture()
  f.plan.sources[0]!.excerpt = '12\n'
  const basis = f.page.chartData![0]!.series[0]!.points[0]!.basis
  if (basis.kind === 'source') basis.excerptText = '12\n'
  expect(
    checkPresentationChartData(f.plan, f.page.id, f.actual).charts[0]!.findings.map((f) => f.code),
  ).toContain('chart_data_source_basis_mismatch')
})
