import {
  parsePresentationAcceptanceContract,
  type PresentationAcceptanceContract,
  type PresentationAcceptanceCheck,
  type SafeScalar,
  type SupportedProperty,
  type SupportedRole,
} from '@wiswork/presentation-verification'

export type SlidesAcceptanceElement = {
  /** Opaque durable creation identity, never the renderer's runtime source id. */
  targetToken: string
  role?: Exclude<SupportedRole, 'background'>
  locked: boolean
  properties: Partial<Record<SupportedProperty, SafeScalar>>
}

export type SlidesAcceptanceAuthority = {
  documentToken: string
  sessionToken: string
  revision: string
  /** Pre-edit revision proved by the mutation receipt; omitted for compile-time inspection. */
  baseRevision?: string
  slides: Array<{
    number: number
    slideToken: string
    backgroundColor?: string
    elements: SlidesAcceptanceElement[]
  }>
  /** Durable targets reported by the proved mutation receipt, when verifying scope. */
  mutatedTargetTokens?: string[]
}

type PropertyChange = {
  kind: 'set_property'
  slides: number[]
  role?: Exclude<SupportedRole, 'background'>
  targetToken?: string
  property: SupportedProperty
  value: SafeScalar
  tolerance?: number
}

export type SlidesAcceptanceIntent = {
  taskId: string
  affectedSlides: number[]
  changes: Array<
    | PropertyChange
    | { kind: 'set_background'; slides: number[]; color: string }
    | {
        kind: 'match_reference'
        slides: number[]
        referenceSlide?: number
        role: Exclude<SupportedRole, 'background'>
        properties: SupportedProperty[]
        tolerance: number
      }
  >
  maxCorrectionPasses: 0 | 1 | 2
}

export type SlidesAcceptanceCompileResult =
  | { status: 'compiled'; contract: PresentationAcceptanceContract }
  | { status: 'unchanged'; taskId: string; affectedSlides: number[] }
  | {
      status: 'needs_clarification'
      code: 'missing_reference' | 'ambiguous_reference' | 'ambiguous_target' | 'target_missing'
      slide?: number
      role?: Exclude<SupportedRole, 'background'>
    }
  | {
      status: 'needs_user'
      code: 'locked_target'
      slide: number
      targetToken: string
    }

export type SlidesDeterministicResult = {
  checkId: string
  status: 'pass' | 'fail' | 'unavailable'
  code?:
    'target_missing' | 'stale_revision' | 'unsupported_check' | 'value_mismatch' | 'scope_mismatch'
}

type CheckWithoutId = PresentationAcceptanceCheck extends infer Check
  ? Check extends PresentationAcceptanceCheck
    ? Omit<Check, 'id'>
    : never
  : never

const deterministicProperties = new Set<SupportedProperty>([
  'text',
  'color',
  'x',
  'y',
  'width',
  'height',
  'fill_color',
  'stroke_color',
  'background_color',
])
const geometryProperties = new Set<SupportedProperty>(['x', 'y', 'width', 'height'])

const normalizeValue = (property: SupportedProperty, value: SafeScalar): SafeScalar =>
  property.endsWith('color') && typeof value === 'string' ? value.toUpperCase() : value

function slideAt(authority: SlidesAcceptanceAuthority, number: number) {
  return authority.slides.find((slide) => slide.number === number)
}

function roleResolution(
  authority: SlidesAcceptanceAuthority,
  slide: number,
  role: Exclude<SupportedRole, 'background'>,
) {
  return slideAt(authority, slide)?.elements.filter((element) => element.role === role) ?? []
}

function targetResolution(
  authority: SlidesAcceptanceAuthority,
  slide: number,
  change: Pick<PropertyChange, 'role' | 'targetToken'>,
) {
  const elements = slideAt(authority, slide)?.elements ?? []
  return change.targetToken
    ? elements.filter((element) => element.targetToken === change.targetToken)
    : change.role
      ? elements.filter((element) => element.role === change.role)
      : []
}

export function compileSlidesAcceptance(
  intent: SlidesAcceptanceIntent,
  authority: SlidesAcceptanceAuthority,
): SlidesAcceptanceCompileResult {
  const affected = new Set(intent.affectedSlides)
  for (const change of intent.changes)
    for (const slide of change.slides)
      if (!affected.has(slide))
        throw new TypeError(`Slide ${slide} is outside frozen affected scope`)

  // Resolve references first so an ambiguous style source produces the material question.
  for (const change of intent.changes) {
    if (change.kind !== 'match_reference') continue
    if (change.referenceSlide === undefined)
      return { status: 'needs_clarification', code: 'missing_reference', role: change.role }
    const candidates = roleResolution(authority, change.referenceSlide, change.role)
    if (candidates.length !== 1)
      return {
        status: 'needs_clarification',
        code: candidates.length ? 'ambiguous_reference' : 'target_missing',
        slide: change.referenceSlide,
        role: change.role,
      }
  }

  const checks: PresentationAcceptanceCheck[] = []
  const referenceSlides = new Set<number>()
  let changed = false
  const addCheck = (check: CheckWithoutId) =>
    checks.push({
      ...check,
      id: `check-${String(checks.length + 1).padStart(3, '0')}`,
    } as PresentationAcceptanceCheck)

  for (const change of intent.changes) {
    if (change.kind === 'set_background') {
      for (const slide of change.slides) {
        const current = slideAt(authority, slide)
        if (!current) return { status: 'needs_clarification', code: 'target_missing', slide }
        const expected = normalizeValue('background_color', change.color)
        changed ||= normalizeValue('background_color', current.backgroundColor ?? null) !== expected
        addCheck({
          kind: 'element_property',
          slide,
          roleOrTarget: { kind: 'role', role: 'background' },
          property: 'background_color',
          expected,
        })
      }
      continue
    }
    if (change.kind === 'set_property') {
      for (const slide of change.slides) {
        const candidates = targetResolution(authority, slide, change)
        if (candidates.length !== 1)
          return {
            status: 'needs_clarification',
            code: candidates.length ? 'ambiguous_target' : 'target_missing',
            slide,
            ...(change.role ? { role: change.role } : {}),
          }
        const target = candidates[0]!
        if (target.locked)
          return {
            status: 'needs_user',
            code: 'locked_target',
            slide,
            targetToken: target.targetToken,
          }
        const expected = normalizeValue(change.property, change.value)
        changed ||=
          normalizeValue(change.property, target.properties[change.property] ?? null) !== expected
        addCheck({
          kind: 'element_property',
          slide,
          roleOrTarget: { kind: 'target', targetToken: target.targetToken },
          property: change.property,
          expected,
        })
      }
      continue
    }
    const referenceSlide = change.referenceSlide!
    referenceSlides.add(referenceSlide)
    const reference = roleResolution(authority, referenceSlide, change.role)[0]!
    for (const slide of change.slides) {
      const candidates = roleResolution(authority, slide, change.role)
      if (candidates.length !== 1)
        return {
          status: 'needs_clarification',
          code: candidates.length ? 'ambiguous_target' : 'target_missing',
          slide,
          role: change.role,
        }
      const target = candidates[0]!
      if (target.locked)
        return {
          status: 'needs_user',
          code: 'locked_target',
          slide,
          targetToken: target.targetToken,
        }
      for (const property of change.properties) {
        const expected = reference.properties[property] ?? null
        const actual = target.properties[property] ?? null
        changed ||= !valuesMatch(property, actual, expected, change.tolerance)
        addCheck({
          kind: 'element_property',
          slide,
          roleOrTarget: { kind: 'target', targetToken: target.targetToken },
          property,
          expected,
        })
      }
    }
  }
  if (!changed)
    return {
      status: 'unchanged',
      taskId: intent.taskId,
      affectedSlides: [...intent.affectedSlides],
    }
  return {
    status: 'compiled',
    contract: parsePresentationAcceptanceContract({
      version: 1,
      taskId: intent.taskId,
      documentToken: authority.documentToken,
      sessionToken: authority.sessionToken,
      baseRevision: authority.revision,
      affectedSlides: [...intent.affectedSlides],
      referenceSlides: [...referenceSlides],
      checks,
      maxCorrectionPasses: intent.maxCorrectionPasses,
    }),
  }
}

function valuesMatch(
  property: SupportedProperty,
  actual: SafeScalar,
  expected: SafeScalar,
  tolerance = geometryProperties.has(property) ? 0.5 : 0,
): boolean {
  if (typeof actual === 'number' && typeof expected === 'number')
    return Math.abs(actual - expected) <= tolerance
  return normalizeValue(property, actual) === normalizeValue(property, expected)
}

export function verifySlidesAcceptance(
  rawContract: PresentationAcceptanceContract,
  authority: SlidesAcceptanceAuthority,
): SlidesDeterministicResult[] {
  const contract = parsePresentationAcceptanceContract(rawContract)
  const stale =
    authority.documentToken !== contract.documentToken ||
    authority.sessionToken !== contract.sessionToken ||
    (authority.baseRevision ?? authority.revision) !== contract.baseRevision
  if (stale)
    return contract.checks.map(({ id }) => ({
      checkId: id,
      status: 'fail',
      code: 'stale_revision',
    }))

  const allowedTargets = new Set(
    contract.checks.flatMap((check) =>
      check.kind === 'element_property' && check.roleOrTarget.kind === 'target'
        ? [check.roleOrTarget.targetToken]
        : check.kind === 'element_property' &&
            check.roleOrTarget.kind === 'role' &&
            check.roleOrTarget.role === 'background'
          ? [slideAt(authority, check.slide)?.slideToken ?? '']
          : [],
    ),
  )
  if (authority.mutatedTargetTokens?.some((target) => !allowedTargets.has(target)))
    return contract.checks.map(({ id }) => ({
      checkId: id,
      status: 'fail',
      code: 'scope_mismatch',
    }))

  return contract.checks.map((check): SlidesDeterministicResult => {
    if (check.kind !== 'element_property' || !deterministicProperties.has(check.property))
      return { checkId: check.id, status: 'unavailable', code: 'unsupported_check' }
    let actual: SafeScalar | undefined
    if (check.roleOrTarget.kind === 'role') {
      if (check.roleOrTarget.role !== 'background')
        return { checkId: check.id, status: 'unavailable', code: 'unsupported_check' }
      actual = slideAt(authority, check.slide)?.backgroundColor
    } else {
      const targetToken = check.roleOrTarget.targetToken
      const target = slideAt(authority, check.slide)?.elements.find(
        (element) => element.targetToken === targetToken,
      )
      if (!target) return { checkId: check.id, status: 'fail', code: 'target_missing' }
      actual = target.properties[check.property]
    }
    if (actual === undefined)
      return { checkId: check.id, status: 'unavailable', code: 'unsupported_check' }
    return valuesMatch(check.property, actual, check.expected)
      ? { checkId: check.id, status: 'pass' }
      : { checkId: check.id, status: 'fail', code: 'value_mismatch' }
  })
}
