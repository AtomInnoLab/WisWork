import type { PresentationPlan } from './presentation-plan'
import { reproducePresentationCalculation } from './presentation-calculation'
import { canonicalPresentationValue as canonical } from '@wiswork/project-store/presentation-canonical'
import { array, choice, id, number, object, text, valid } from './presentation-schema'
export interface PresentationChartDataBinding {
  elementId: string
  categories: string[]
  series: {
    name: string
    points: {
      value: number
      claimId: string
      basis:
        | { kind: 'source'; sourceId: string; excerptOffset: number; excerptText: string }
        | { kind: 'calculation' }
    }[]
  }[]
  unit?: string
  currency?: string
}
export const PRESENTATION_CHART_DATA_SCHEMA = array(
  object(
    {
      elementId: id,
      categories: array(text(200, 1), 50, 1),
      series: array(
        object({
          name: text(200, 1),
          points: array(
            object({
              value: number(-1e15, 1e15),
              claimId: id,
              basis: {
                anyOf: [
                  object({
                    kind: choice('source'),
                    sourceId: id,
                    excerptOffset: number(0, 12000),
                    excerptText: text(80, 1),
                  }),
                  object({ kind: choice('calculation') }),
                ],
              },
            }),
            50,
            1,
          ),
        }),
        10,
        1,
      ),
      unit: text(100, 1),
      currency: text(100, 1),
    },
    ['elementId', 'categories', 'series'],
  ),
  128,
)
const codes = [
  'chart_data_unbound',
  'chart_data_missing',
  'chart_data_shape_mismatch',
  'chart_data_value_mismatch',
  'chart_data_source_basis_mismatch',
  'chart_data_calculation_not_reproduced',
  'chart_data_unit_mismatch',
  'chart_data_currency_mismatch',
] as const
export type PresentationChartDataFindingCode = (typeof codes)[number]
type Actual = { categories: string[]; series: { name: string; values: number[] }[] }
export type PresentationActualChartData = Actual & { id: string; kind?: 'chart' }
export interface PresentationChartDataCheck {
  version: 1
  pageId: string
  scope: 'frozen_declared_data'
  charts: {
    elementId: string
    actual?: Actual
    findings: { code: PresentationChartDataFindingCode; claimIds: string[] }[]
  }[]
  checks: { data: 'needs_review'; sourceTruth: 'not_verified'; host: 'not_checked' }
}
const actualSchema = object({
  categories: array(text(200, 1), 50, 1),
  series: array(object({ name: text(200, 1), values: array(number(-1e15, 1e15), 50, 1) }), 10, 1),
})
const checks = { data: 'needs_review', sourceTruth: 'not_verified', host: 'not_checked' } as const
const checkSchema = object({
  version: { type: 'number', enum: [1] },
  pageId: id,
  scope: choice('frozen_declared_data'),
  charts: array(
    object(
      {
        elementId: id,
        actual: actualSchema,
        findings: array(object({ code: choice(...codes), claimIds: array(id, 32) }), 8),
      },
      ['elementId', 'findings'],
    ),
    256,
  ),
  checks: object({
    data: choice('needs_review'),
    sourceTruth: choice('not_verified'),
    host: choice('not_checked'),
  }),
})
function invalid(): never {
  throw new Error('presentation_chart_data_invalid')
}
/** Literal basis preserves code units; the number check later deliberately rejects non-numeric text. */
export function parsePresentationChartDataBindings(value: unknown): PresentationChartDataBinding[] {
  const literalSafe = (v: unknown): unknown =>
    Array.isArray(v)
      ? v.map(literalSafe)
      : v && typeof v === 'object'
        ? Object.fromEntries(
            Object.entries(v).map(([k, item]) => [
              k,
              k === 'excerptText' && typeof item === 'string'
                ? 'x'.repeat(item.length)
                : literalSafe(item),
            ]),
          )
        : v
  if (!valid(literalSafe(value), PRESENTATION_CHART_DATA_SCHEMA)) invalid()
  const charts = value as PresentationChartDataBinding[]
  if (new Set(charts.map((c) => c.elementId)).size !== charts.length) invalid()
  for (const chart of charts)
    for (const s of chart.series) {
      if (s.points.length !== chart.categories.length) invalid()
      for (const p of s.points)
        if (p.basis.kind === 'source' && !Number.isSafeInteger(p.basis.excerptOffset)) invalid()
    }
  return structuredClone(charts)
}
export function checkPresentationChartData(
  plan: PresentationPlan,
  pageId: string,
  charts: readonly PresentationActualChartData[],
): PresentationChartDataCheck {
  const page = plan.slides.find((p) => p.id === pageId)
  if (!page || !Array.isArray(charts) || charts.length > 128) invalid()
  const actuals = new Map<string, Actual>()
  for (const chart of charts) {
    if (
      !chart ||
      !valid(chart.id, id) ||
      (chart.kind !== undefined && chart.kind !== 'chart') ||
      actuals.has(chart.id)
    )
      invalid()
    const actual: Actual = { categories: chart.categories, series: chart.series }
    if (
      !valid(actual, actualSchema) ||
      actual.series.some((s) => s.values.length !== actual.categories.length)
    )
      invalid()
    actuals.set(chart.id, actual)
  }
  const declared = page.chartData ?? [],
    ids = [...declared.map((c) => c.elementId), ...actuals.keys()].filter(
      (v, i, a) => a.indexOf(v) === i,
    )
  const report: PresentationChartDataCheck = {
    version: 1,
    pageId,
    scope: 'frozen_declared_data',
    charts: [],
    checks: { ...checks },
  }
  for (const elementId of ids) {
    const binding = declared.find((c) => c.elementId === elementId),
      actual = actuals.get(elementId),
      findings: PresentationChartDataCheck['charts'][number]['findings'] = []
    const add = (code: PresentationChartDataFindingCode, claimIds: string[] = []) => {
      const found = findings.find((f) => f.code === code)
      if (found) found.claimIds = [...new Set([...found.claimIds, ...claimIds])]
      else findings.push({ code, claimIds: [...new Set(claimIds)] })
    }
    if (!binding) add('chart_data_unbound')
    else {
      const claimIds = binding.series.flatMap((s) => s.points.map((p) => p.claimId))
      if (!actual) add('chart_data_missing', claimIds)
      else {
        if (
          canonical(actual.categories) !== canonical(binding.categories) ||
          actual.series.length !== binding.series.length ||
          actual.series.some(
            (s, i) =>
              s.name !== binding.series[i]?.name ||
              s.values.length !== binding.series[i]?.points.length,
          )
        )
          add('chart_data_shape_mismatch', claimIds)
        for (const [i, s] of binding.series.entries())
          for (const [j, p] of s.points.entries())
            if (
              actual.series[i]?.values[j] !== undefined &&
              actual.series[i]!.values[j] !== p.value
            )
              add('chart_data_value_mismatch', [p.claimId])
      }
      for (const s of binding.series)
        for (const p of s.points) {
          const claim = plan.claims.find((c) => c.id === p.claimId)
          if (!claim || !page.claimIds.includes(p.claimId)) invalid()
          if (p.basis.kind === 'source') {
            const basis = p.basis,
              source = plan.sources.find((s) => s.id === basis.sourceId)
            if (!source || !claim.sourceIds.includes(basis.sourceId)) invalid()
            const start = basis.excerptOffset,
              end = start + basis.excerptText.length
            const before = source.excerpt[start - 1] ?? '',
              after = source.excerpt[end] ?? ''
            // Conservative token boundaries prevent cherry-picking digits from a signed, exponent or grouped number.
            const fragment =
              /[0-9.eE+-]/.test(before + after) ||
              (before === ',' && /[0-9]/.test(source.excerpt[start - 2] ?? '')) ||
              (after === ',' && /[0-9]/.test(source.excerpt[end + 1] ?? ''))
            if (
              fragment ||
              source.excerpt.slice(
                basis.excerptOffset,
                basis.excerptOffset + basis.excerptText.length,
              ) !== basis.excerptText ||
              !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(basis.excerptText) ||
              Number(basis.excerptText) !== p.value
            )
              add('chart_data_source_basis_mismatch', [p.claimId])
          } else {
            const calculation = reproducePresentationCalculation(claim)
            if (calculation.status !== 'reproduced' || calculation.actual !== p.value)
              add('chart_data_calculation_not_reproduced', [p.claimId])
          }
          for (const key of ['unit', 'currency'] as const) {
            const professional =
              claim.professionalContext?.domain === 'finance'
                ? claim.professionalContext[key]
                : undefined
            if (
              [claim.calculation?.[key], professional].some(
                (v) => v !== undefined && v !== binding[key],
              )
            )
              add(key === 'unit' ? 'chart_data_unit_mismatch' : 'chart_data_currency_mismatch', [
                p.claimId,
              ])
          }
        }
    }
    report.charts.push({
      elementId,
      ...(actual ? { actual: structuredClone(actual) } : {}),
      findings,
    })
  }
  return report
}
export function parsePresentationChartDataCheck(
  value: unknown,
  plan?: PresentationPlan,
): PresentationChartDataCheck {
  if (
    new TextEncoder().encode(JSON.stringify(value)).length > 8 * 1024 * 1024 ||
    !valid(value, checkSchema)
  )
    invalid()
  const result = value as PresentationChartDataCheck
  if (new Set(result.charts.map((c) => c.elementId)).size !== result.charts.length) invalid()
  for (const chart of result.charts) {
    if (
      new Set(chart.findings.map((f) => f.code)).size !== chart.findings.length ||
      chart.findings.some((f) => new Set(f.claimIds).size !== f.claimIds.length) ||
      chart.actual?.series.some((s) => s.values.length !== chart.actual!.categories.length)
    )
      invalid()
  }
  if (plan) {
    const rebuilt = checkPresentationChartData(
      plan,
      result.pageId,
      result.charts.flatMap((c) => (c.actual ? [{ id: c.elementId, ...c.actual }] : [])),
    )
    if (canonical(rebuilt) !== canonical(result)) invalid()
  }
  return structuredClone(result)
}
