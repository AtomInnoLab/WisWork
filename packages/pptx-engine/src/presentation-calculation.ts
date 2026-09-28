import type { PresentationPlan } from './presentation-plan'

/** Round the decimal form of a finite JS result, with ties away from zero. */
function roundDecimal(value: number, places: number): number {
  const match = /^(\d+)(?:\.(\d+))?(?:e([+-]?\d+))?$/i.exec(Math.abs(value).toString())
  if (!match) throw new Error('invalid_arithmetic')
  const mantissa = BigInt(match[1]! + (match[2] ?? ''))
  const shift = Number(match[3] ?? 0) - (match[2]?.length ?? 0) + places
  const rounded =
    shift >= 0
      ? mantissa * 10n ** BigInt(shift)
      : (mantissa + 10n ** BigInt(-shift) / 2n) / 10n ** BigInt(-shift)
  if (rounded > BigInt(Number.MAX_SAFE_INTEGER)) throw new Error('invalid_arithmetic')
  return Math.sign(value) * (Number(rounded) / 10 ** places)
}

export interface CalculationResult {
  claimId: string
  status:
    'not_configured' | 'reproduced' | 'mismatch' | 'unsupported_expression' | 'invalid_arithmetic'
  actual?: number
  expected?: number
  tolerance?: number
  scope: 'arithmetic_only'
}

/** Bounded arithmetic grammar. Never evaluates JavaScript or resolves object properties. */
export function reproducePresentationCalculation(
  claim: PresentationPlan['claims'][number],
): CalculationResult {
  const result: CalculationResult = {
    claimId: claim.id,
    status: 'not_configured',
    scope: 'arithmetic_only',
  }
  const configuration = claim.calculation?.reproduction
  if (!configuration) return result
  result.expected = configuration.expected
  const unsupported = (): never => {
    throw new Error('unsupported_expression')
  }
  const arithmetic = (value: number): number => {
    if (!Number.isFinite(value) || Math.abs(value) > 1e12) throw new Error('invalid_arithmetic')
    return value
  }
  try {
    const formula = claim.calculation!.formula
    const tokens: string[] = []
    const lex =
      /\s*(?:(\d+(?:\.\d*)?(?:[eE][+-]?\d+)?|\.\d+(?:[eE][+-]?\d+)?)|([A-Za-z][A-Za-z0-9_]*)|([+*/(),-]))/y
    let offset = 0
    while (offset < formula.length) {
      if (!formula.slice(offset).trim()) break
      lex.lastIndex = offset
      const match = lex.exec(formula)
      if (!match || tokens.length >= 256) unsupported()
      tokens.push(match![1] ?? match![2] ?? match![3]!)
      offset = lex.lastIndex
    }
    const bindings = new Map(configuration.bindings.map((binding) => [binding.name, binding.value]))
    const used = new Set<string>()
    let cursor = 0,
      operations = 0
    const operation = (): void => {
      if (++operations > 128) unsupported()
    }
    const primary = (depth: number): number => {
      if (depth > 32) unsupported()
      const token = tokens[cursor++]
      if (token === '+' || token === '-') {
        operation()
        return arithmetic((token === '-' ? -1 : 1) * primary(depth))
      }
      if (token === '(') {
        const value = expression(depth + 1)
        if (tokens[cursor++] !== ')') unsupported()
        return value
      }
      if (token === 'round' && tokens[cursor] === '(') {
        operation()
        cursor++
        const value = expression(depth + 1)
        if (tokens[cursor++] !== ',') unsupported()
        const places = expression(depth + 1)
        if (tokens[cursor++] !== ')' || !Number.isInteger(places) || places < 0 || places > 6)
          unsupported()
        return arithmetic(roundDecimal(value, places))
      }
      if (token && /^(?:\d|\.)/.test(token)) return arithmetic(Number(token))
      if (token && bindings.has(token)) {
        used.add(token)
        return arithmetic(bindings.get(token)!)
      }
      return unsupported()
    }
    const product = (depth: number): number => {
      let value = primary(depth)
      while (tokens[cursor] === '*' || tokens[cursor] === '/') {
        const operator = tokens[cursor++]
        operation()
        const right = primary(depth)
        if (operator === '/' && right === 0) throw new Error('invalid_arithmetic')
        value = arithmetic(operator === '*' ? value * right : value / right)
      }
      return value
    }
    const expression = (depth: number): number => {
      let value = product(depth)
      while (tokens[cursor] === '+' || tokens[cursor] === '-') {
        const operator = tokens[cursor++]
        operation()
        const right = product(depth)
        value = arithmetic(operator === '+' ? value + right : value - right)
      }
      return value
    }
    const actual = expression(0)
    if (cursor !== tokens.length || used.size !== bindings.size) unsupported()
    const tolerance =
      8 * Number.EPSILON * Math.max(1, Math.abs(actual), Math.abs(configuration.expected))
    return {
      ...result,
      status: Math.abs(actual - configuration.expected) <= tolerance ? 'reproduced' : 'mismatch',
      actual,
      tolerance,
    }
  } catch (error) {
    return {
      ...result,
      status:
        error instanceof Error && error.message === 'invalid_arithmetic'
          ? 'invalid_arithmetic'
          : 'unsupported_expression',
    }
  }
}
