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
  textMatches?: Record<string, { targetToken: string; matches: boolean; proof: string }>
  slides: Array<{
    number: number
    slideToken: string
    backgroundColor?: string
    elements: SlidesAcceptanceElement[]
  }>
}

type PropertyChange = {
  kind: 'set_property'
  slides: number[]
  role?: Exclude<SupportedRole, 'background'>
  targetToken?: string
  property: SupportedProperty
  value: SafeScalar
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
  | {
      status: 'compiled'
      contract: PresentationAcceptanceContract
      plannedMutationTargets: string[]
    }
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
const roles = new Set(['title', 'body', 'emphasis'])
const properties = new Set([
  'text',
  'color',
  'font_size',
  'font_family',
  'bold',
  'italic',
  'x',
  'y',
  'width',
  'height',
  'fill_color',
  'stroke_color',
  'background_color',
])

function strictObject(
  value: unknown,
  keys: readonly string[],
  name: string,
): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new TypeError(`${name} must be a plain object`)
  if (Object.getPrototypeOf(value) !== Object.prototype)
    throw new TypeError(`${name} must be a plain object`)
  const record = value as Record<string, unknown>
  for (const key of Reflect.ownKeys(record)) {
    if (typeof key !== 'string' || !keys.includes(key))
      throw new TypeError(`${name} has an unknown field`)
    const descriptor = Object.getOwnPropertyDescriptor(record, key)
    if (!descriptor?.enumerable || descriptor.get || descriptor.set)
      throw new TypeError(`${name} has an unsafe field`)
  }
  return record
}

function strictArray(value: unknown, maximum: number, name: string): unknown[] {
  if (
    !Array.isArray(value) ||
    Object.getPrototypeOf(value) !== Array.prototype ||
    value.length > maximum
  )
    throw new TypeError(`${name} is invalid`)
  for (let index = 0; index < value.length; index += 1)
    if (!Object.prototype.hasOwnProperty.call(value, index))
      throw new TypeError(`${name} has holes`)
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || (key !== 'length' && !/^(0|[1-9][0-9]*)$/.test(key)))
      throw new TypeError(`${name} has an unsafe field`)
    const descriptor = Object.getOwnPropertyDescriptor(value, key)
    if (
      !descriptor ||
      descriptor.get ||
      descriptor.set ||
      (key !== 'length' && !descriptor.enumerable)
    )
      throw new TypeError(`${name} has an unsafe field`)
  }
  return value
}

const identifier = (value: unknown, name: string): string => {
  if (typeof value !== 'string' || !/^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/.test(value))
    throw new TypeError(`${name} is invalid`)
  return value
}
const slideNumbers = (value: unknown, name: string): number[] =>
  strictArray(value, 50, name).map((item) => {
    if (!Number.isSafeInteger(item) || (item as number) < 1 || (item as number) > 10_000)
      throw new TypeError(`${name} is invalid`)
    return item as number
  })

export function parseSlidesAcceptanceIntent(value: unknown): SlidesAcceptanceIntent {
  const root = strictObject(
    value,
    ['taskId', 'affectedSlides', 'changes', 'maxCorrectionPasses'],
    'intent',
  )
  const changes = strictArray(root.changes, 50, 'intent.changes').map((item, index) => {
    const base = strictObject(
      item,
      [
        'kind',
        'slides',
        'role',
        'targetToken',
        'property',
        'value',
        'tolerance',
        'color',
        'referenceSlide',
        'properties',
      ],
      `change ${index}`,
    )
    const slides = slideNumbers(base.slides, `change ${index}.slides`)
    if (base.kind === 'set_background') {
      if (
        Object.keys(base).some((key) => !['kind', 'slides', 'color'].includes(key)) ||
        typeof base.color !== 'string' ||
        !/^#[0-9A-Fa-f]{6}$/.test(base.color)
      )
        throw new TypeError(`change ${index} is invalid`)
      return { kind: 'set_background' as const, slides, color: base.color }
    }
    if (base.kind === 'match_reference') {
      if (
        Object.keys(base).some(
          (key) =>
            !['kind', 'slides', 'referenceSlide', 'role', 'properties', 'tolerance'].includes(key),
        )
      )
        throw new TypeError(`change ${index} is invalid`)
      const role = base.role
      const props = strictArray(base.properties, 16, `change ${index}.properties`)
      if (
        !roles.has(String(role)) ||
        !Number.isSafeInteger(base.referenceSlide) ||
        typeof base.tolerance !== 'number' ||
        !Number.isFinite(base.tolerance) ||
        base.tolerance < 0 ||
        props.length === 0 ||
        new Set(props).size !== props.length ||
        props.some((property) => !properties.has(String(property)))
      )
        throw new TypeError(`change ${index} is invalid`)
      return {
        kind: 'match_reference' as const,
        slides,
        referenceSlide: base.referenceSlide as number,
        role: role as Exclude<SupportedRole, 'background'>,
        properties: props as SupportedProperty[],
        tolerance: base.tolerance,
      }
    }
    if (
      base.kind !== 'set_property' ||
      Object.keys(base).some(
        (key) =>
          !['kind', 'slides', 'role', 'targetToken', 'property', 'value', 'tolerance'].includes(
            key,
          ),
      ) ||
      !properties.has(String(base.property)) ||
      (base.role === undefined) === (base.targetToken === undefined)
    )
      throw new TypeError(`change ${index} is invalid`)
    if (base.tolerance !== undefined) throw new TypeError(`change ${index} is invalid`)
    if (base.role !== undefined && !roles.has(String(base.role)))
      throw new TypeError(`change ${index} is invalid`)
    if (base.targetToken !== undefined) identifier(base.targetToken, `change ${index}.targetToken`)
    if (
      !(base.value === null || ['string', 'number', 'boolean'].includes(typeof base.value)) ||
      (typeof base.value === 'number' && !Number.isFinite(base.value)) ||
      (base.tolerance !== undefined &&
        (typeof base.tolerance !== 'number' ||
          !Number.isFinite(base.tolerance) ||
          base.tolerance < 0))
    )
      throw new TypeError(`change ${index}.value is invalid`)
    return {
      kind: 'set_property' as const,
      slides,
      ...(base.role === undefined
        ? { targetToken: base.targetToken as string }
        : { role: base.role as Exclude<SupportedRole, 'background'> }),
      property: base.property as SupportedProperty,
      value: base.value as SafeScalar,
      ...(base.tolerance === undefined ? {} : { tolerance: base.tolerance as number }),
    }
  })
  if (changes.length === 0) throw new TypeError('intent.changes is empty')
  if (![0, 1, 2].includes(Number(root.maxCorrectionPasses)))
    throw new TypeError('maxCorrectionPasses is invalid')
  const affectedSlides = slideNumbers(root.affectedSlides, 'affectedSlides')
  if (affectedSlides.length === 0 || new Set(affectedSlides).size !== affectedSlides.length)
    throw new TypeError('affectedSlides is invalid')
  return {
    taskId: identifier(root.taskId, 'taskId'),
    affectedSlides,
    changes,
    maxCorrectionPasses: root.maxCorrectionPasses as 0 | 1 | 2,
  }
}

export function parseSlidesAcceptanceAuthority(value: unknown): SlidesAcceptanceAuthority {
  const root = strictObject(
    value,
    ['documentToken', 'sessionToken', 'revision', 'baseRevision', 'textMatches', 'slides'],
    'authority',
  )
  const slides = strictArray(root.slides, 100, 'authority.slides').map((item, index) => {
    const slide = strictObject(
      item,
      ['number', 'slideToken', 'backgroundColor', 'elements'],
      `authority.slides[${index}]`,
    )
    const elements = strictArray(slide.elements, 500, `authority.slides[${index}].elements`).map(
      (raw, elementIndex) => {
        const element = strictObject(
          raw,
          ['targetToken', 'role', 'locked', 'properties'],
          `element ${elementIndex}`,
        )
        const facts = strictObject(
          element.properties,
          [...properties],
          `element ${elementIndex}.properties`,
        )
        for (const value of Object.values(facts))
          if (
            !(value === null || ['string', 'number', 'boolean'].includes(typeof value)) ||
            (typeof value === 'number' && !Number.isFinite(value)) ||
            (typeof value === 'string' && value.length > 256)
          )
            throw new TypeError('element property is invalid')
        if (element.role !== undefined && !roles.has(String(element.role)))
          throw new TypeError('element role is invalid')
        if (typeof element.locked !== 'boolean') throw new TypeError('element locked is invalid')
        return {
          targetToken: identifier(element.targetToken, 'targetToken'),
          ...(element.role === undefined
            ? {}
            : { role: element.role as Exclude<SupportedRole, 'background'> }),
          locked: element.locked,
          properties: facts as Partial<Record<SupportedProperty, SafeScalar>>,
        }
      },
    )
    return {
      number: slideNumbers([slide.number], 'slide number')[0]!,
      slideToken: identifier(slide.slideToken, 'slideToken'),
      ...(slide.backgroundColor === undefined
        ? {}
        : typeof slide.backgroundColor === 'string' &&
            /^#[0-9A-Fa-f]{6}$/.test(slide.backgroundColor)
          ? { backgroundColor: slide.backgroundColor }
          : (() => {
              throw new TypeError('backgroundColor is invalid')
            })()),
      elements,
    }
  })
  const revision = String(root.revision)
  if (!/^sha256:[0-9a-f]{64}$/.test(revision)) throw new TypeError('revision is invalid')
  if (root.baseRevision !== undefined && !/^sha256:[0-9a-f]{64}$/.test(String(root.baseRevision)))
    throw new TypeError('baseRevision is invalid')
  const slideNumbersSeen = slides.map((slide) => slide.number)
  if (new Set(slideNumbersSeen).size !== slideNumbersSeen.length)
    throw new TypeError('slide numbers must be unique')
  let textMatches: SlidesAcceptanceAuthority['textMatches']
  if (root.textMatches !== undefined) {
    const record = strictObject(
      root.textMatches,
      Object.keys(root.textMatches as object),
      'textMatches',
    )
    textMatches = {}
    for (const [checkId, raw] of Object.entries(record)) {
      identifier(checkId, 'text check id')
      const match = strictObject(raw, ['targetToken', 'matches', 'proof'], 'text match')
      if (
        typeof match.matches !== 'boolean' ||
        typeof match.proof !== 'string' ||
        !/^sha256:[0-9a-f]{64}$/.test(match.proof)
      )
        throw new TypeError('text match is invalid')
      textMatches[checkId] = {
        targetToken: identifier(match.targetToken, 'text match target'),
        matches: match.matches,
        proof: match.proof,
      }
    }
  }
  return {
    documentToken: identifier(root.documentToken, 'documentToken'),
    sessionToken: identifier(root.sessionToken, 'sessionToken'),
    revision,
    ...(root.baseRevision === undefined ? {} : { baseRevision: String(root.baseRevision) }),
    ...(textMatches ? { textMatches } : {}),
    slides,
  }
}

const normalizeValue = (property: SupportedProperty, value: SafeScalar): SafeScalar =>
  property.endsWith('color') && typeof value === 'string' ? value.toUpperCase() : value

function slideAt(authority: SlidesAcceptanceAuthority, number: number) {
  return authority.slides.find((slide) => slide.number === number)
}

function roleResolution(authority: SlidesAcceptanceAuthority, slide: number, role: SupportedRole) {
  if (role === 'background') return []
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
  rawIntent: SlidesAcceptanceIntent,
  rawAuthority: SlidesAcceptanceAuthority,
): SlidesAcceptanceCompileResult {
  const intent = parseSlidesAcceptanceIntent(rawIntent)
  const authority = parseSlidesAcceptanceAuthority(rawAuthority)
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
  const plannedMutationTargets = new Set<string>()
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
        const needsChange =
          normalizeValue('background_color', current.backgroundColor ?? null) !== expected
        changed ||= needsChange
        if (needsChange) plannedMutationTargets.add(current.slideToken)
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
        const actual = target.properties[change.property]
        const checkId = `check-${String(checks.length + 1).padStart(3, '0')}`
        const textMatch = change.property === 'text' ? authority.textMatches?.[checkId] : undefined
        const needsChange =
          change.property === 'text'
            ? !textMatch || textMatch.targetToken !== target.targetToken || !textMatch.matches
            : actual == null ||
              !deterministicProperties.has(change.property) ||
              normalizeValue(change.property, actual) !== expected
        changed ||= needsChange
        if (needsChange) plannedMutationTargets.add(target.targetToken)
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
      const reference = roleResolution(authority, referenceSlide, change.role)[0]!
      const needsChange = change.properties.some((property) => {
        const actual = target.properties[property]
        const expected = reference.properties[property]
        return (
          actual == null ||
          expected == null ||
          !deterministicProperties.has(property) ||
          !valuesMatch(property, actual, expected, change.tolerance)
        )
      })
      changed ||= needsChange
      if (needsChange) plannedMutationTargets.add(target.targetToken)
      addCheck({
        kind: 'reference_match',
        slide,
        referenceSlide,
        role: change.role,
        properties: change.properties,
        tolerance: change.tolerance,
      })
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
    plannedMutationTargets: [...plannedMutationTargets],
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
  rawAuthority: SlidesAcceptanceAuthority,
  proof:
    | { mode: 'prewrite' }
    | {
        mode: 'postwrite'
        mutatedTargetTokens: string[]
        plannedMutationTargets: string[]
      },
): SlidesDeterministicResult[] {
  const contract = parsePresentationAcceptanceContract(rawContract)
  const authority = parseSlidesAcceptanceAuthority(rawAuthority)
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

  const requiredTargets = new Set(
    contract.checks.flatMap((check) =>
      check.kind === 'element_property' && check.roleOrTarget.kind === 'target'
        ? [check.roleOrTarget.targetToken]
        : check.kind === 'element_property' &&
            check.roleOrTarget.kind === 'role' &&
            check.roleOrTarget.role === 'background'
          ? [slideAt(authority, check.slide)?.slideToken ?? '']
          : check.kind === 'reference_match'
            ? roleResolution(authority, check.slide, check.role).map((target) => target.targetToken)
            : [],
    ),
  )
  if (proof.mode === 'postwrite') {
    const plannedTargets = proof.plannedMutationTargets.map((target) =>
      identifier(target, 'plannedMutationTarget'),
    )
    const proved = new Set(
      proof.mutatedTargetTokens.map((target) => identifier(target, 'mutatedTargetToken')),
    )
    if (
      proved.size !== proof.mutatedTargetTokens.length ||
      new Set(plannedTargets).size !== plannedTargets.length ||
      plannedTargets.some((target) => !requiredTargets.has(target)) ||
      proved.size !== plannedTargets.length ||
      [...proved].some((target) => !plannedTargets.includes(target))
    )
      return contract.checks.map(({ id }) => ({
        checkId: id,
        status: 'fail',
        code: 'scope_mismatch',
      }))
  }
  return contract.checks.map((check): SlidesDeterministicResult => {
    if (check.kind === 'reference_match') {
      const target = roleResolution(authority, check.slide, check.role)
      const reference = roleResolution(authority, check.referenceSlide, check.role)
      if (target.length !== 1 || reference.length !== 1)
        return { checkId: check.id, status: 'fail', code: 'target_missing' }
      for (const property of check.properties) {
        if (!deterministicProperties.has(property))
          return { checkId: check.id, status: 'unavailable', code: 'unsupported_check' }
        const actual = target[0]!.properties[property]
        const expected = reference[0]!.properties[property]
        if (actual == null || expected == null)
          return { checkId: check.id, status: 'unavailable', code: 'unsupported_check' }
        if (!valuesMatch(property, actual, expected, check.tolerance))
          return { checkId: check.id, status: 'fail', code: 'value_mismatch' }
      }
      return { checkId: check.id, status: 'pass' }
    }
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
    if (check.property === 'text') {
      const textMatch = authority.textMatches?.[check.id]
      if (
        !textMatch ||
        check.roleOrTarget.kind !== 'target' ||
        textMatch.targetToken !== check.roleOrTarget.targetToken
      )
        return { checkId: check.id, status: 'unavailable', code: 'unsupported_check' }
      return textMatch.matches
        ? { checkId: check.id, status: 'pass' }
        : { checkId: check.id, status: 'fail', code: 'value_mismatch' }
    }
    if (actual == null)
      return { checkId: check.id, status: 'unavailable', code: 'unsupported_check' }
    return valuesMatch(check.property, actual, check.expected)
      ? { checkId: check.id, status: 'pass' }
      : { checkId: check.id, status: 'fail', code: 'value_mismatch' }
  })
}
