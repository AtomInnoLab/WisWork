import { describe, expect, it } from 'vitest'
import { reproducePresentationCalculation } from '../src/presentation-calculation'
import type { PresentationPlan } from '../src/presentation-plan'
const claim = (formula: string, expected = 3): PresentationPlan['claims'][number] => ({
  id: 'c',
  statement: 'result',
  type: 'calculation',
  sourceIds: ['s'],
  confidence: 'low',
  reviewStatus: 'needs_review',
  calculation: {
    formula,
    inputs: ['one'],
    reproduction: { bindings: [{ name: 'a', inputIndex: 0, value: 1, sourceId: 's' }], expected },
  },
})
describe('bounded reproduction', () => {
  it('reproduces precedence and IEEE rounding', () => {
    expect(reproducePresentationCalculation(claim('a + 2 * (4 - 3)')).status).toBe('reproduced')
    expect(reproducePresentationCalculation(claim('a * .1 + 2e-1', 0.3)).status).toBe('reproduced')
    expect(reproducePresentationCalculation(claim('a', 2)).status).toBe('mismatch')
  })
  it('reproduces a labeled two-decimal FX conversion and decimal half-away rounding', () => {
    const conversion = claim('round(a / 144.4, 2)', 90171.52)
    conversion.calculation!.reproduction!.bindings[0]!.value = 13020768
    expect(reproducePresentationCalculation(conversion).status).toBe('reproduced')
    const half = claim('round(a + 0.005, 2)', 1.01)
    expect(reproducePresentationCalculation(half).status).toBe('reproduced')
    expect(reproducePresentationCalculation(claim('round(-a - 0.005, 2)', -1.01)).status).toBe(
      'reproduced',
    )
  })
  it.each(['round(a, -1)', 'round(a, 7)', 'round(a, 1.5)', 'round(a)', 'round(a, 2, 3)'])(
    'rejects unsupported rounding precision or arity %s',
    (formula) =>
      expect(reproducePresentationCalculation(claim(formula)).status).toBe(
        'unsupported_expression',
      ),
  )
  it.each([
    'a.constructor',
    'eval(a)',
    'a=3',
    'a**2',
    'a;fetch(1)',
    'b+1',
    '3',
    'a + ' + '('.repeat(33) + '1' + ')'.repeat(33),
    'a' + '+1'.repeat(129),
  ])('rejects syntax and limits %s', (formula) =>
    expect(reproducePresentationCalculation(claim(formula)).status).toBe('unsupported_expression'),
  )
  it.each(['a/0', 'a*1e13', 'a+1e999'])('rejects invalid arithmetic %s', (formula) =>
    expect(reproducePresentationCalculation(claim(formula)).status).toBe('invalid_arithmetic'),
  )
  it('rejects rounding whose scaled integer exceeds exact JS precision', () => {
    const value = claim('round(a, 6)', 1e12)
    value.calculation!.reproduction!.bindings[0]!.value = 1e12
    expect(reproducePresentationCalculation(value).status).toBe('invalid_arithmetic')
  })
  it('allows unary operators within the separate operation budget', () =>
    expect(reproducePresentationCalculation(claim('-'.repeat(40) + 'a', 1)).status).toBe(
      'reproduced',
    ))
  it('retains unconfigured legacy claims', () => {
    const value = claim('a')
    delete value.calculation!.reproduction
    expect(reproducePresentationCalculation(value).status).toBe('not_configured')
  })
})
