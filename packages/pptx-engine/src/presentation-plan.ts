import type { PresentationClaim, PresentationDeck, PresentationStyle } from './presentation'
import { PRESENTATION_DECK_SCHEMA, PRESENTATION_TEXT_BUDGET } from './presentation'
import { type Schema, text, number, choice, array, object, id, valid } from './presentation-schema'

/** Mandatory deck JSON may consume 192 KiB of the 256 KiB compile transport.
 * The remaining 64 KiB accommodates content/geometry and the request envelope;
 * optional content and assets still require the normal compile transport checks.
 */
export const PRESENTATION_PLAN_COMPILED_BYTE_BUDGET = 192 * 1024

/** Durable planning metadata; source URIs are never fetched or treated as verified evidence. */
export interface PresentationPlan {
  version: 1
  projectId: string
  title: string
  brief: {
    objective: string
    audience: string
    language: string
    minutes: number
    requiredContent: string[]
    constraints: string[]
  }
  sources: {
    id: string
    title: string
    uri: string
    locator?: string
    excerpt: string
    asOf?: string
  }[]
  claims: {
    id: string
    statement: string
    type: 'fact' | 'quote' | 'calculation' | 'judgment' | 'assumption'
    sourceIds: string[]
    confidence: 'high' | 'medium' | 'low'
    reviewStatus: 'needs_review'
    asOf?: string
    jurisdiction?: string
    calculation?: {
      formula: string
      inputs: string[]
      unit?: string
      currency?: string
      reproduction?: {
        bindings: { name: string; inputIndex: number; value: number; sourceId: string }[]
        expected: number
      }
    }
  }[]
  style: PresentationStyle
  slides: {
    id: string
    title: string
    purpose: string
    claimIds: string[]
    layout: 'cover' | 'content' | 'comparison' | 'process' | 'chart' | 'summary'
    requiredAssets: string[]
    acceptanceCriteria: string[]
  }[]
}

/** Tool discovery and runtime validation share bounded structural rules. */
export const PRESENTATION_PLAN_SCHEMA: Schema = object({
  version: { type: 'number', enum: [1] },
  projectId: id,
  title: text(300, 1),
  brief: object({
    objective: text(4000, 1),
    audience: text(1000, 1),
    language: text(80, 1),
    minutes: number(0.1, 1440),
    requiredContent: array(text(2000, 1), 100),
    constraints: array(text(2000, 1), 100),
  }),
  sources: array(
    object(
      {
        id,
        title: text(300, 1),
        uri: text(500, 1),
        locator: text(200),
        excerpt: text(12000),
        asOf: text(100, 1),
      },
      ['id', 'title', 'uri', 'excerpt'],
    ),
    256,
  ),
  claims: array(
    object(
      {
        id,
        statement: text(12000, 1),
        type: choice('fact', 'quote', 'calculation', 'judgment', 'assumption'),
        sourceIds: array(id, 3),
        confidence: choice('high', 'medium', 'low'),
        reviewStatus: choice('needs_review'),
        asOf: text(100, 1),
        jurisdiction: text(300, 1),
        calculation: object(
          {
            formula: text(2000, 1),
            inputs: array(text(1000, 1), 32, 1),
            unit: text(100, 1),
            currency: text(100, 1),
            reproduction: object({
              bindings: array(
                object({
                  name: { ...text(32, 1), pattern: '^[A-Za-z][A-Za-z0-9_]*$' },
                  inputIndex: number(0, 31),
                  value: number(-1e12, 1e12),
                  sourceId: id,
                }),
                32,
                1,
              ),
              expected: number(-1e12, 1e12),
            }),
          },
          ['formula', 'inputs'],
        ),
      },
      ['id', 'statement', 'type', 'sourceIds', 'confidence', 'reviewStatus'],
    ),
    256,
  ),
  style: PRESENTATION_DECK_SCHEMA.properties!.style!,
  slides: array(
    object({
      id,
      title: text(300, 1),
      purpose: text(2000, 1),
      claimIds: array(id, 32),
      layout: choice('cover', 'content', 'comparison', 'process', 'chart', 'summary'),
      requiredAssets: array(text(1000, 1), 32),
      acceptanceCriteria: array(text(2000, 1), 32),
    }),
    32,
    1,
  ),
})

function reject(reason: string): never {
  throw new Error(`presentation_plan_invalid:${reason}`)
}
function unique(values: string[], label: string): void {
  if (new Set(values).size !== values.length) reject(`duplicate_${label}`)
}

export function parsePresentationPlan(input: unknown): PresentationPlan {
  if (!valid(input, PRESENTATION_PLAN_SCHEMA)) reject('schema')
  const plan = input as PresentationPlan
  if (new TextEncoder().encode(JSON.stringify(plan)).byteLength > 192 * 1024) reject('size_budget')
  unique(
    plan.sources.map((source) => source.id),
    'source',
  )
  unique(
    plan.claims.map((claim) => claim.id),
    'claim',
  )
  unique(
    plan.slides.map((slide) => slide.id),
    'slide',
  )
  const sourceIds = new Set(plan.sources.map((source) => source.id))
  const claimIds = new Set(plan.claims.map((claim) => claim.id))
  for (const claim of plan.claims) {
    unique(claim.sourceIds, 'source_reference')
    if (claim.sourceIds.some((source) => !sourceIds.has(source))) reject('source_reference')
    if (['fact', 'quote', 'calculation'].includes(claim.type) && !claim.sourceIds.length)
      reject('source_required')
    if (claim.type === 'calculation' && !claim.calculation) reject('calculation_required')
    const reproduction = claim.calculation?.reproduction
    if (reproduction) {
      if (
        claim.type !== 'calculation' ||
        reproduction.bindings.length !== claim.calculation!.inputs.length
      )
        reject('reproduction_inputs')
      unique(
        reproduction.bindings.map((binding) => binding.name),
        'binding_name',
      )
      unique(
        reproduction.bindings.map((binding) => String(binding.inputIndex)),
        'binding_index',
      )
      for (const binding of reproduction.bindings) {
        if (
          !Number.isInteger(binding.inputIndex) ||
          binding.inputIndex >= claim.calculation!.inputs.length ||
          ['prototype', 'constructor', '__proto__'].includes(binding.name) ||
          !claim.sourceIds.includes(binding.sourceId)
        )
          reject('reproduction_binding')
      }
    }
  }
  for (const slide of plan.slides) {
    unique(slide.claimIds, 'claim_reference')
    if (slide.claimIds.some((claim) => !claimIds.has(claim))) reject('claim_reference')
  }
  const claims = mappedClaims(plan)
  const requiredText = claims.reduce(
    (sum, claim) => sum + claim.text.length + claim.source.length + (claim.locator?.length ?? 0),
    plan.slides.reduce((sum, slide) => sum + slide.title.length, 0),
  )
  if (requiredText > PRESENTATION_TEXT_BUDGET) reject('text_budget')
  const mandatoryDeck = {
    version: 1,
    id: plan.projectId,
    title: plan.title,
    style: plan.style,
    assets: [],
    claims,
    slides: plan.slides.map(({ id, title, claimIds }) => ({ id, title, claimIds, elements: [] })),
  }
  if (
    new TextEncoder().encode(JSON.stringify(mandatoryDeck)).byteLength >
    PRESENTATION_PLAN_COMPILED_BYTE_BUDGET
  )
    reject('compiled_byte_budget')
  return structuredClone(plan)
}

export function presentationPlanClaims(plan: PresentationPlan): PresentationClaim[] {
  return mappedClaims(parsePresentationPlan(plan))
}

function mappedClaims(plan: PresentationPlan): PresentationClaim[] {
  const sources = new Map(plan.sources.map((source) => [source.id, source]))
  return plan.claims.map((claim) => {
    const evidence = claim.sourceIds.map((id) => sources.get(id)!)
    const locator = evidence
      .map((source) => source.locator)
      .filter((value) => value?.trim())
      .join(' ; ')
    return {
      id: claim.id,
      text: claim.statement,
      source: evidence.length
        ? evidence.map((source) => source.uri).join(' ; ')
        : `未核验：${claim.type}`,
      ...(locator ? { locator } : {}),
    }
  })
}

/** Deterministic identity/content binding only; this makes no visual or factual QA claim. */
export function assertDeckMatchesPresentationPlan(
  deck: PresentationDeck,
  plan: PresentationPlan,
): void {
  const expectedClaims = presentationPlanClaims(plan)
  const mismatch = (reason: string): never => {
    throw new Error(`presentation_plan_mismatch:${reason}`)
  }
  if (deck.id !== plan.projectId || deck.title !== plan.title) mismatch('project')
  for (const key of ['fontFace', 'background', 'textColor', 'accentColor'] as const)
    if (deck.style[key] !== plan.style[key]) mismatch('style')
  if (deck.slides.length !== plan.slides.length) mismatch('slides')
  for (const [index, slide] of deck.slides.entries()) {
    const planned = plan.slides[index]!
    if (
      slide.id !== planned.id ||
      slide.title !== planned.title ||
      JSON.stringify(slide.claimIds ?? []) !== JSON.stringify(planned.claimIds)
    )
      mismatch('slide')
  }
  if (deck.claims.length !== expectedClaims.length) mismatch('claims')
  for (const [index, claim] of deck.claims.entries()) {
    const expected = expectedClaims[index]!
    if (
      claim.id !== expected.id ||
      claim.text !== expected.text ||
      claim.source !== expected.source ||
      claim.locator !== expected.locator
    )
      mismatch('claim')
  }
}
