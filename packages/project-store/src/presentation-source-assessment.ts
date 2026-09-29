import {
  PROFESSIONAL_CONTEXT_SCHEMA,
  parsePresentationProfessionalContext,
  type PresentationProfessionalContext,
} from './presentation-professional-context'
import { canonicalPresentationValue } from './presentation-canonical'
/** Historical Agent opinions; literal basis never authenticates a source or its claims. */
export interface PresentationSourceAssessment {
  scope: string
  authority: {
    outcome: 'appropriate_for_claim' | 'insufficient_authority' | 'uncertain'
    sourceTier: 'primary' | 'authoritative_secondary' | 'secondary' | 'unverified'
    reason: string
  }
  timeliness: {
    outcome: 'current_for_claim' | 'historical_only' | 'superseded' | 'uncertain'
    referenceDate: string
    claimAsOf?: string
    sourceAsOf?: string
    reason: string
  }
  jurisdiction?: {
    claimJurisdiction: string
    outcome: 'applicable' | 'mismatch' | 'uncertain'
    reason: string
  }
  professional?: {
    context: PresentationProfessionalContext
    checks: {
      aspect: 'conclusion_scope' | 'qualifications' | 'comparability' | 'forecast'
      outcome: 'consistent' | 'conflict' | 'uncertain' | 'not_applicable'
      reason: string
    }[]
  }
  basis: { offset: number; text: string }[]
}
interface Schema {
  type?: string
  anyOf?: Schema[]
  minItems?: number
  properties?: Record<string, Schema>
  required?: string[]
  additionalProperties?: boolean
  minLength?: number
  maxLength?: number
  enum?: string[]
  pattern?: string
  minimum?: number
  maximum?: number
  maxItems?: number
  items?: Schema
}
const text = (maxLength: number) => ({ type: 'string', minLength: 1, maxLength })
const choice = (...values: string[]) => ({ type: 'string', enum: values })
const object = (properties: Record<string, Schema>, required = Object.keys(properties)) => ({
  type: 'object',
  additionalProperties: false,
  properties,
  required,
})
export const PRESENTATION_SOURCE_ASSESSMENT_SCHEMA = object(
  {
    scope: text(400),
    authority: object({
      outcome: choice('appropriate_for_claim', 'insufficient_authority', 'uncertain'),
      sourceTier: choice('primary', 'authoritative_secondary', 'secondary', 'unverified'),
      reason: text(600),
    }),
    timeliness: object(
      {
        outcome: choice('current_for_claim', 'historical_only', 'superseded', 'uncertain'),
        referenceDate: { ...text(10), pattern: '^[0-9]{4}-[0-9]{2}-[0-9]{2}$' },
        claimAsOf: text(100),
        sourceAsOf: text(100),
        reason: text(600),
      },
      ['outcome', 'referenceDate', 'reason'],
    ),
    jurisdiction: object({
      claimJurisdiction: text(400),
      outcome: choice('applicable', 'mismatch', 'uncertain'),
      reason: text(600),
    }),
    professional: object({
      context: PROFESSIONAL_CONTEXT_SCHEMA,
      checks: {
        type: 'array',
        minItems: 2,
        maxItems: 2,
        items: object({
          aspect: choice('conclusion_scope', 'qualifications', 'comparability', 'forecast'),
          outcome: choice('consistent', 'conflict', 'uncertain', 'not_applicable'),
          reason: text(600),
        }),
      },
    }),
    basis: {
      type: 'array',
      maxItems: 4,
      items: object({ offset: { type: 'integer', minimum: 0, maximum: 1000000 }, text: text(600) }),
    },
  },
  ['scope', 'authority', 'timeliness', 'basis'],
)
function invalid(): never {
  throw new Error('source_assessment_invalid')
}
function validate(value: unknown, schema: Schema, literal = false): void {
  if (schema.anyOf) {
    try {
      parsePresentationProfessionalContext(value)
    } catch {
      invalid()
    }
    return
  }
  if (schema.type === 'object') {
    if (!value || typeof value !== 'object' || Array.isArray(value)) invalid()
    const v = value as Record<string, unknown>
    if (
      Object.keys(v).some((k) => !Object.hasOwn(schema.properties!, k)) ||
      schema.required!.some((k: string) => !Object.hasOwn(v, k))
    )
      invalid()
    for (const [k, vv] of Object.entries(v))
      validate(vv, schema.properties![k]!, literal || k === 'basis')
  } else if (schema.type === 'array') {
    if (
      !Array.isArray(value) ||
      value.length > schema.maxItems! ||
      value.length < (schema.minItems ?? 0)
    )
      invalid()
    value.forEach((v) => validate(v, schema.items!, literal))
  } else if (schema.type === 'integer') {
    if (
      !Number.isSafeInteger(value) ||
      Number(value) < schema.minimum! ||
      Number(value) > schema.maximum!
    )
      invalid()
  } else {
    if (
      typeof value !== 'string' ||
      (!literal && !value.trim()) ||
      value.length < (schema.minLength ?? 0) ||
      value.length > (schema.maxLength ?? Infinity) ||
      (schema.enum && !schema.enum.includes(value)) ||
      (schema.pattern && !new RegExp(schema.pattern).test(value))
    )
      invalid()
    if (
      !literal &&
      // eslint-disable-next-line no-control-regex
      /[^\u0009\u000a\u000d\u0020-\ud7ff\ue000-\ufffd\u{10000}-\u{10ffff}]/u.test(value)
    )
      invalid()
  }
}
export function parsePresentationSourceAssessment(value: unknown): PresentationSourceAssessment {
  validate(value, PRESENTATION_SOURCE_ASSESSMENT_SCHEMA)
  const result = value as PresentationSourceAssessment
  const date = result.timeliness.referenceDate
  if (
    !Number.isFinite(Date.parse(date)) ||
    new Date(date).toISOString().slice(0, 10) !== date ||
    new TextEncoder().encode(JSON.stringify(value)).length > 16 * 1024
  )
    invalid()
  if (
    new Set(result.basis.map((b) => JSON.stringify([b.offset, b.text]))).size !==
    result.basis.length
  )
    invalid()
  if (
    !result.basis.length &&
    (result.authority.outcome === 'appropriate_for_claim' ||
      result.timeliness.outcome === 'current_for_claim' ||
      result.jurisdiction?.outcome === 'applicable')
  )
    invalid()
  if (result.professional) {
    const { context, checks } = result.professional
    const required =
      context.domain === 'finance'
        ? ['comparability', 'forecast']
        : ['conclusion_scope', 'qualifications']
    if (
      new Set(checks.map((c) => c.aspect)).size !== 2 ||
      checks.some(
        (c) =>
          !required.includes(c.aspect) ||
          (c.outcome === 'not_applicable' && c.aspect !== 'forecast') ||
          (c.outcome !== 'uncertain' && !result.basis.length),
      )
    )
      invalid()
  }
  return structuredClone(result)
}
export function assertPresentationSourceAssessmentBasis(
  value: PresentationSourceAssessment,
  window: { offset: number; text: string },
): void {
  const assessment = parsePresentationSourceAssessment(value)
  if (!Number.isSafeInteger(window.offset) || window.offset < 0 || typeof window.text !== 'string')
    invalid()
  for (const basis of assessment.basis) {
    const start = basis.offset - window.offset
    if (
      start < 0 ||
      start + basis.text.length > window.text.length ||
      window.text.slice(start, start + basis.text.length) !== basis.text
    )
      invalid()
  }
}

/** Bind historical opinions to the complete frozen declared context, without authenticating it. */
export function assertPresentationProfessionalAssessmentContext(
  value: PresentationSourceAssessment,
  context?: PresentationProfessionalContext,
): void {
  const assessment = parsePresentationSourceAssessment(value)
  if (!assessment.professional) return
  try {
    if (
      !context ||
      canonicalPresentationValue(assessment.professional.context) !==
        canonicalPresentationValue(parsePresentationProfessionalContext(context))
    )
      invalid()
  } catch {
    invalid()
  }
}
