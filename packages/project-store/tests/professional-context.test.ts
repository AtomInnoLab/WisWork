import { expect, it } from 'vitest'
import {
  parsePresentationProfessionalContext,
  presentationProfessionalContextMissingFields,
} from '../src/presentation-professional-context'
it('accepts partial context without invented fields', () => {
  expect(parsePresentationProfessionalContext({ domain: 'science' })).toEqual({ domain: 'science' })
  expect(
    presentationProfessionalContextMissingFields({ domain: 'finance' }, 'calculation'),
  ).toContain('formula')
})
it.each([
  { domain: 'science', jurisdiction: '跨领域' },
  { domain: 'law', materialKind: 'paper' },
  { domain: 'finance', asOf: '2026-02-30' },
  { domain: 'law', effectiveFrom: '2026-09-29', effectiveUntil: '2026-09-28' },
  { domain: 'science', version: undefined },
  { domain: 'science', method: '\ud800' },
  { domain: 'finance', unknown: true },
])('strictly rejects malformed context %#', (v) =>
  expect(() => parsePresentationProfessionalContext(v)).toThrow(),
)
it('uses exact conditional required fields without inferring expiry or applicability', () => {
  expect(
    presentationProfessionalContextMissingFields({ domain: 'law', materialKind: 'contract' }),
  ).not.toContain('effectiveFrom')
  expect(
    presentationProfessionalContextMissingFields({ domain: 'law', materialKind: 'case' }),
  ).toEqual([
    'jurisdiction',
    'effectLevel',
    'applicabilityDate',
    'originalLocation',
    'limitations',
    'effectiveFrom',
    'caseNumber',
  ])
  expect(presentationProfessionalContextMissingFields({ domain: 'finance' })).not.toContain(
    'formula',
  )
  expect(
    parsePresentationProfessionalContext({
      domain: 'law',
      effectiveFrom: '2024-02-29',
      effectiveUntil: '2024-02-29',
    }),
  ).toHaveProperty('effectiveUntil', '2024-02-29')
})
it('rejects text overflow and XML controls', () => {
  for (const context of [
    { domain: 'science', method: 'x'.repeat(801) },
    { domain: 'science', limitations: 'bad\f' },
  ])
    expect(() => parsePresentationProfessionalContext(context)).toThrow()
})
it('accepts complete science and finance contexts with no missing fields', () => {
  const science = {
    domain: 'science',
    materialKind: 'paper',
    publicationId: 'doi:example',
    version: 'v1',
    sample: '声明样本',
    method: '声明方法',
    statisticalBasis: '声明统计依据',
    limitations: '尚待专家检查',
  } as const
  expect(presentationProfessionalContextMissingFields(science)).toEqual([])
  const finance = {
    domain: 'finance',
    materialKind: 'disclosure',
    reportingPeriod: '2025全年',
    asOf: '2025-12-31',
    currency: 'CNY',
    unit: '万元',
    accountingBasis: '声明口径',
    formula: 'x+y',
    limitations: '未认证',
  } as const
  expect(presentationProfessionalContextMissingFields(finance, 'calculation')).toEqual([])
})
