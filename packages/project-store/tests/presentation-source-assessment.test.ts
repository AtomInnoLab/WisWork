import { expect, it } from 'vitest'
import {
  parsePresentationSourceAssessment,
  assertPresentationProfessionalAssessmentContext,
  assertPresentationSourceAssessmentBasis,
} from '../src/presentation-source-assessment'
const sample = {
  scope: '仅此主张',
  authority: { outcome: 'appropriate_for_claim', sourceTier: 'primary', reason: '原文判断' },
  timeliness: { outcome: 'current_for_claim', referenceDate: '2026-09-29', reason: '时点判断' },
  basis: [{ offset: 2, text: '原\f文' }],
}
it('preserves literal UTF16 basis and validates actual window', () => {
  const value = parsePresentationSourceAssessment(sample)
  expect(value).toEqual(sample)
  assertPresentationSourceAssessmentBasis(value, { offset: 0, text: '😀原\f文' })
  expect(() =>
    assertPresentationSourceAssessmentBasis(value, { offset: 1, text: '😀原\f文' }),
  ).toThrow()
})
it.each([
  { ...sample, unknown: true },
  { ...sample, timeliness: { ...sample.timeliness, referenceDate: '2026-02-30' } },
  { ...sample, basis: [] },
  { ...sample, basis: [...sample.basis, ...sample.basis] },
  { ...sample, scope: 'bad\f' },
  { ...sample, scope: '\ud800' },
])('rejects malformed assessment %#', (value) =>
  expect(() => parsePresentationSourceAssessment(value)).toThrow(),
)
it('allows uncertain judgment without basis', () =>
  expect(
    parsePresentationSourceAssessment({
      ...sample,
      authority: { ...sample.authority, outcome: 'uncertain' },
      timeliness: { ...sample.timeliness, outcome: 'uncertain' },
      basis: [],
    }),
  ).toHaveProperty('basis', []))
it('preserves whitespace literal data but rejects whitespace reasons and basis outside window', () => {
  const value = { ...sample, basis: [{ offset: 2, text: ' \f ' }] }
  expect(parsePresentationSourceAssessment(value).basis[0]!.text).toBe(' \f ')
  expect(() =>
    assertPresentationSourceAssessmentBasis(parsePresentationSourceAssessment(value), {
      offset: 3,
      text: ' \f ',
    }),
  ).toThrow()
  expect(() =>
    parsePresentationSourceAssessment({
      ...sample,
      authority: { ...sample.authority, reason: '  ' },
    }),
  ).toThrow()
})
it.each([
  { ...sample, basis: [{ offset: 0.5, text: 'x' }] },
  { ...sample, basis: [{ offset: 1000001, text: 'x' }] },
  { ...sample, basis: Array.from({ length: 5 }, (_, offset) => ({ offset, text: 'x' })) },
  { ...sample, authority: { ...sample.authority, sourceTier: 'certified' } },
  { ...sample, timeliness: { ...sample.timeliness, referenceDate: '2026-9-29' } },
])('rejects bounds and unauthorized enum %#', (v) =>
  expect(() => parsePresentationSourceAssessment(v)).toThrow(),
)
it('caps UTF8 JSON even when literal basis contains preserved escaped controls', () => {
  const value = {
    ...sample,
    scope: '界'.repeat(400),
    authority: { ...sample.authority, reason: '界'.repeat(600) },
    timeliness: { ...sample.timeliness, reason: '界'.repeat(600) },
    basis: Array.from({ length: 4 }, (_, offset) => ({ offset, text: '\u0000'.repeat(600) })),
  }
  expect(new TextEncoder().encode(JSON.stringify(value)).length).toBeGreaterThan(16 * 1024)
  expect(() => parsePresentationSourceAssessment(value)).toThrow('source_assessment_invalid')
})
it('preserves full professional context and requires the exact frozen context', () => {
  const context = { domain: 'science' as const, sample: ' sample ', limitations: 'scope only' }
  const value = parsePresentationSourceAssessment({
    ...sample,
    professional: {
      context,
      checks: [
        { aspect: 'conclusion_scope', outcome: 'consistent', reason: ' bounded opinion ' },
        { aspect: 'qualifications', outcome: 'uncertain', reason: 'unverified' },
      ],
    },
  })
  expect(value.professional?.context).toEqual(context)
  expect(value.professional?.checks[0]?.reason).toBe(' bounded opinion ')
  assertPresentationProfessionalAssessmentContext(value, {
    limitations: 'scope only',
    sample: ' sample ',
    domain: 'science',
  })
  expect(() => assertPresentationProfessionalAssessmentContext(value)).toThrow()
  expect(() =>
    assertPresentationProfessionalAssessmentContext(value, { ...context, sample: 'sample' }),
  ).toThrow()
  assertPresentationProfessionalAssessmentContext(parsePresentationSourceAssessment(sample))
})
it.each([
  ['science', ['conclusion_scope', 'qualifications']],
  ['law', ['conclusion_scope', 'qualifications']],
  ['finance', ['comparability', 'forecast']],
])('requires exact two domain aspects for %s', (domain, aspects) => {
  const professional = {
    context: { domain },
    checks: (aspects as string[]).map((aspect) => ({
      aspect,
      outcome: 'uncertain',
      reason: 'historical opinion',
    })),
  }
  expect(parsePresentationSourceAssessment({ ...sample, professional })).toHaveProperty(
    'professional',
    professional,
  )
  for (const checks of [
    professional.checks.slice(0, 1),
    [...professional.checks, professional.checks[0]],
    [professional.checks[0], professional.checks[0]],
    professional.checks.map((c) => ({ ...c, aspect: 'forecast' })),
  ]) {
    expect(() =>
      parsePresentationSourceAssessment({ ...sample, professional: { ...professional, checks } }),
    ).toThrow()
  }
})
it('requires basis for all definite professional outcomes and allows not_applicable only for forecast', () => {
  const base = {
    ...sample,
    authority: { ...sample.authority, outcome: 'uncertain' },
    timeliness: { ...sample.timeliness, outcome: 'uncertain' },
    basis: [],
  }
  const professional = {
    context: { domain: 'finance' },
    checks: [
      { aspect: 'comparability', outcome: 'uncertain', reason: 'unknown' },
      { aspect: 'forecast', outcome: 'uncertain', reason: 'unknown' },
    ],
  }
  expect(parsePresentationSourceAssessment({ ...base, professional })).toHaveProperty('basis', [])
  for (const outcome of ['consistent', 'conflict', 'not_applicable']) {
    expect(() =>
      parsePresentationSourceAssessment({
        ...base,
        professional: {
          ...professional,
          checks: [professional.checks[0], { ...professional.checks[1], outcome }],
        },
      }),
    ).toThrow()
  }
  expect(
    parsePresentationSourceAssessment({
      ...sample,
      professional: {
        ...professional,
        checks: [professional.checks[0], { ...professional.checks[1], outcome: 'not_applicable' }],
      },
    }),
  ).toHaveProperty('professional')
  expect(() =>
    parsePresentationSourceAssessment({
      ...sample,
      professional: {
        ...professional,
        checks: [{ ...professional.checks[0], outcome: 'not_applicable' }, professional.checks[1]],
      },
    }),
  ).toThrow()
  expect(() =>
    parsePresentationSourceAssessment({
      ...sample,
      professional: { ...professional, context: { domain: 'law', effectiveFrom: '2026-02-30' } },
    }),
  ).toThrow()
})
