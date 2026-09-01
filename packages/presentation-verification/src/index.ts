const unsafeKeys = new Set(['__proto__', 'prototype', 'constructor'])
const identifierPattern = /^[A-Za-z0-9][A-Za-z0-9._:-]*$/
const digestPattern = /^sha256:[0-9a-f]{64}$/
const colorPattern = /^#[0-9A-Fa-f]{6}$/
const arrayIndexPattern = /^(0|[1-9][0-9]*)$/

export const PRESENTATION_VERIFICATION_LIMITS = Object.freeze({
  maxChecks: 50,
  maxAffectedSlides: 50,
  maxReferenceSlides: 50,
  maxScreenshotsPerPass: 8,
  maxVisualRequestBytes: 2 * 1024 * 1024,
  maxCorrectionPasses: 2,
  maxObservations: 64,
  maxFixIntents: 50,
  maxIdentifierLength: 128,
  maxSafeStringLength: 256,
  maxPropertiesPerCheck: 16,
  maxRenderRulesPerCheck: 16,
  maxReceiptIds: 50,
})

export type PresentationVerificationFlags = Readonly<{
  planning: boolean
  verifiedCompletion: boolean
  visualReview: boolean
  autoCorrection: boolean
}>

export function presentationVerificationFlags(
  env: Record<string, string | undefined>,
  prefix = 'WISWORK_PRESENTATION_',
): PresentationVerificationFlags {
  const flag = (name: string, fallback: boolean) => {
    const value = env[`${prefix}${name}`]
    if (value === undefined || value === '') return fallback
    if (value === '1') return true
    if (value === '0') return false
    throw new Error('invalid_presentation_verification_flags')
  }
  return Object.freeze({
    planning: flag('PLANNING', true),
    verifiedCompletion: flag('VERIFIED_COMPLETION', true),
    visualReview: flag('VISUAL_REVIEW', true),
    // Correction remains opt-in while golden evaluation is rolling out.
    autoCorrection: flag('AUTO_CORRECTION', false),
  })
}

export type SupportedRole = 'title' | 'body' | 'emphasis' | 'background'
export type SupportedProperty =
  | 'text'
  | 'color'
  | 'font_size'
  | 'font_family'
  | 'bold'
  | 'italic'
  | 'x'
  | 'y'
  | 'width'
  | 'height'
  | 'fill_color'
  | 'stroke_color'
  | 'background_color'
export type RenderRule =
  | 'no_overflow'
  | 'no_overlap'
  | 'within_slide'
  | 'legible'
  | 'consistent_alignment'
  | 'reference_similarity'
export type SafeScalar = string | number | boolean | null
export type TargetRef =
  { kind: 'role'; role: SupportedRole } | { kind: 'target'; targetToken: string }

export type PresentationAcceptanceCheck =
  | {
      id: string
      kind: 'element_property'
      slide: number
      roleOrTarget: TargetRef
      property: SupportedProperty
      expected: SafeScalar
    }
  | {
      id: string
      kind: 'reference_match'
      slide: number
      referenceSlide: number
      role: SupportedRole
      properties: SupportedProperty[]
      tolerance: number
    }
  | { id: string; kind: 'render_quality'; slide: number; rules: RenderRule[] }

export type PresentationAcceptanceContract = {
  version: 1
  taskId: string
  documentToken: string
  sessionToken: string
  baseRevision: string
  affectedSlides: number[]
  referenceSlides: number[]
  checks: PresentationAcceptanceCheck[]
  maxCorrectionPasses: 0 | 1 | 2
}

export type SafeObservation = {
  code:
    | 'overflow'
    | 'overlap'
    | 'out_of_bounds'
    | 'illegible'
    | 'misaligned'
    | 'reference_mismatch'
    | 'review_unavailable'
  severity: 'info' | 'warning' | 'error'
  checkId?: string
  slide?: number
}

export type SafeFixIntent = {
  checkId: string
  kind: 'set_property'
  roleOrTarget: TargetRef
  property: SupportedProperty
  value: SafeScalar
}

export type VisualReviewResult = {
  status: 'pass' | 'needs_fix' | 'cannot_verify'
  failedCheckIds: string[]
  observations: SafeObservation[]
  fixIntents: SafeFixIntent[]
}

export type PresentationCompletionReceipt = {
  version: 1
  taskId: string
  status: 'verified' | 'applied_unverified' | 'needs_user' | 'failed' | 'unchanged'
  mutationReceiptIds: string[]
  passedCheckIds: string[]
  failedCheckIds: string[]
  unavailableCheckIds: string[]
  correctionPasses: number
  affectedSlides: number[]
  rollbackId?: string
  safeCode?: SafeCompletionCode
}

export type SafeCompletionCode =
  | 'screenshot_unavailable'
  | 'review_unavailable'
  | 'verification_invalid'
  | 'cancelled_after_apply'
  | 'office_state_uncertain'
  | 'stale_authority'
  | 'confirmation_required'
  | 'unsupported_check'
  | 'cancelled'
  | 'mutation_failed'
  | 'visual_disabled'

export type PresentationRenderingFacts = {
  contractDigest: string
  revision: string
  screenshots: Array<{
    slide: number
    role: 'affected' | 'reference'
    roles?: Array<'affected' | 'reference'>
    mediaToken: string
    bytes: number
  }>
  deterministicResults: Array<{
    checkId: string
    status: 'pass' | 'fail' | 'unavailable'
    code?:
      | 'target_missing'
      | 'stale_revision'
      | 'unsupported_check'
      | 'value_mismatch'
      | 'scope_mismatch'
  }>
}

function fail(message: string): never {
  throw new TypeError(`Invalid presentation verification data: ${message}`)
}

const readObject = (value: unknown, name: string): Record<string, unknown> => {
  if (typeof value !== 'object' || value === null || Array.isArray(value))
    fail(`${name} must be an object`)
  const record = value as Record<string, unknown>
  const prototype = Object.getPrototypeOf(record)
  if (prototype !== Object.prototype && prototype !== null) fail(`${name} must be a plain object`)
  for (const key of Reflect.ownKeys(record)) {
    if (typeof key !== 'string') fail(`${name} contains an unknown symbol field`)
    if (unsafeKeys.has(key)) fail(`${name} contains unsafe key ${key}`)
    const descriptor = Object.getOwnPropertyDescriptor(record, key)
    if (descriptor === undefined || descriptor.get !== undefined || descriptor.set !== undefined)
      fail(`${name} contains an accessor field`)
    if (!descriptor.enumerable) fail(`${name} contains a hidden field`)
  }
  return record
}

const exactKeys = (
  record: Record<string, unknown>,
  keys: readonly string[],
  name: string,
): void => {
  const allowed = new Set(keys)
  for (const key of Object.keys(record))
    if (!allowed.has(key)) fail(`${name} contains unknown field ${key}`)
}

const readArray = (value: unknown, name: string, maximum: number, minimum = 0): unknown[] => {
  if (!Array.isArray(value)) fail(`${name} must be a standard array`)
  if (Object.getPrototypeOf(value) !== Array.prototype) fail(`${name} must be a standard array`)
  const lengthDescriptor = Object.getOwnPropertyDescriptor(value, 'length')
  if (
    lengthDescriptor === undefined ||
    lengthDescriptor.get !== undefined ||
    lengthDescriptor.set !== undefined ||
    !Number.isSafeInteger(lengthDescriptor.value) ||
    lengthDescriptor.value < minimum ||
    lengthDescriptor.value > maximum
  )
    fail(`${name} length is out of bounds`)
  const descriptors: PropertyDescriptor[] = []
  for (const key of Reflect.ownKeys(value)) {
    if (key === 'length') continue
    if (typeof key !== 'string') fail(`${name} contains an unknown symbol field`)
    if (unsafeKeys.has(key)) fail(`${name} contains unsafe key ${key}`)
    if (!arrayIndexPattern.test(key)) fail(`${name} contains extra field ${key}`)
    const index = Number(key)
    if (index >= lengthDescriptor.value || String(index) !== key)
      fail(`${name} contains an invalid index`)
    const descriptor = Object.getOwnPropertyDescriptor(value, key)
    if (descriptor === undefined || descriptor.get !== undefined || descriptor.set !== undefined)
      fail(`${name} contains an accessor field`)
    if (!descriptor.enumerable) fail(`${name} contains a hidden field`)
    descriptors[index] = descriptor
  }
  return Array.from({ length: lengthDescriptor.value }, (_, index) => {
    const descriptor = descriptors[index]
    if (descriptor === undefined) fail(`${name} must not contain array holes`)
    return descriptor.value
  })
}

const readEnum = <T extends string>(value: unknown, values: readonly T[], name: string): T => {
  if (typeof value !== 'string' || !values.includes(value as T)) fail(`${name} is unknown`)
  return value as T
}

const readIdentifier = (value: unknown, name: string): string => {
  if (
    typeof value !== 'string' ||
    value.length === 0 ||
    value.length > PRESENTATION_VERIFICATION_LIMITS.maxIdentifierLength ||
    !identifierPattern.test(value) ||
    unsafeKeys.has(value)
  )
    fail(`${name} is unsafe`)
  return value
}

const readDigest = (value: unknown, name: string): string => {
  if (typeof value !== 'string' || !digestPattern.test(value))
    fail(`${name} must be a SHA-256 digest`)
  return value
}

const readInteger = (value: unknown, name: string, minimum: number, maximum: number): number => {
  if (!Number.isSafeInteger(value) || (value as number) < minimum || (value as number) > maximum)
    fail(`${name} is out of bounds`)
  return value as number
}

const readFinite = (value: unknown, name: string, minimum: number, maximum: number): number => {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < minimum || value > maximum)
    fail(`${name} is out of bounds`)
  return value
}

const readSafeScalar = (value: unknown, name: string): SafeScalar => {
  if (value === null || typeof value === 'boolean') return value
  if (typeof value === 'number') return readFinite(value, name, -1_000_000, 1_000_000)
  if (typeof value === 'string') {
    if (value.length > PRESENTATION_VERIFICATION_LIMITS.maxSafeStringLength)
      fail(`${name} is out of bounds`)
    return value
  }
  fail(`${name} must be a safe scalar`)
}

const readUniqueIdentifiers = (value: unknown, name: string, maximum: number): string[] => {
  const result = readArray(value, name, maximum).map((item, index) =>
    readIdentifier(item, `${name}[${index}]`),
  )
  if (new Set(result).size !== result.length) fail(`${name} must be unique`)
  return result
}

const readSlides = (value: unknown, name: string, maximum: number): number[] => {
  const result = readArray(value, name, maximum).map((item, index) =>
    readInteger(item, `${name}[${index}]`, 1, 100_000),
  )
  if (new Set(result).size !== result.length) fail(`${name} must be unique`)
  return result
}

const roles = ['title', 'body', 'emphasis', 'background'] as const
const properties = [
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
] as const
const renderRules = [
  'no_overflow',
  'no_overlap',
  'within_slide',
  'legible',
  'consistent_alignment',
  'reference_similarity',
] as const

const readTarget = (value: unknown, name: string): TargetRef => {
  const record = readObject(value, name)
  const kind = readEnum(record.kind, ['role', 'target'] as const, `${name}.kind`)
  if (kind === 'role') {
    exactKeys(record, ['kind', 'role'], name)
    return { kind, role: readEnum(record.role, roles, `${name}.role`) }
  }
  exactKeys(record, ['kind', 'targetToken'], name)
  return { kind, targetToken: readIdentifier(record.targetToken, `${name}.targetToken`) }
}

const readProperty = (value: unknown, name: string): SupportedProperty =>
  readEnum(value, properties, name)

const readPropertyValue = (
  value: unknown,
  property: SupportedProperty,
  name: string,
): SafeScalar => {
  const parsed = readSafeScalar(value, name)
  if (
    (property.endsWith('color') || property === 'color') &&
    (typeof parsed !== 'string' || !colorPattern.test(parsed))
  )
    fail(`${name} must be a color`)
  if (['font_size', 'x', 'y', 'width', 'height'].includes(property) && typeof parsed !== 'number')
    fail(`${name} must be numeric`)
  if (['bold', 'italic'].includes(property) && typeof parsed !== 'boolean')
    fail(`${name} must be boolean`)
  if (['text', 'font_family'].includes(property) && typeof parsed !== 'string')
    fail(`${name} must be a string`)
  return parsed
}

const readCheck = (value: unknown, index: number): PresentationAcceptanceCheck => {
  const name = `checks[${index}]`
  const record = readObject(value, name)
  const kind = readEnum(
    record.kind,
    ['element_property', 'reference_match', 'render_quality'] as const,
    `${name}.kind`,
  )
  const id = readIdentifier(record.id, `${name}.id`)
  const slide = readInteger(record.slide, `${name}.slide`, 1, 100_000)
  if (kind === 'element_property') {
    exactKeys(record, ['id', 'kind', 'slide', 'roleOrTarget', 'property', 'expected'], name)
    const property = readProperty(record.property, `${name}.property`)
    const expected = readPropertyValue(record.expected, property, `${name}.expected`)
    return {
      id,
      kind,
      slide,
      roleOrTarget: readTarget(record.roleOrTarget, `${name}.roleOrTarget`),
      property,
      expected,
    }
  }
  if (kind === 'reference_match') {
    exactKeys(
      record,
      ['id', 'kind', 'slide', 'referenceSlide', 'role', 'properties', 'tolerance'],
      name,
    )
    const parsedProperties = readArray(
      record.properties,
      `${name}.properties`,
      PRESENTATION_VERIFICATION_LIMITS.maxPropertiesPerCheck,
      1,
    ).map((item, propertyIndex) => readProperty(item, `${name}.properties[${propertyIndex}]`))
    if (new Set(parsedProperties).size !== parsedProperties.length)
      fail(`${name}.properties must be unique`)
    return {
      id,
      kind,
      slide,
      referenceSlide: readInteger(record.referenceSlide, `${name}.referenceSlide`, 1, 100_000),
      role: readEnum(record.role, roles, `${name}.role`),
      properties: parsedProperties,
      tolerance: readFinite(record.tolerance, `${name}.tolerance`, 0, 1_000),
    }
  }
  exactKeys(record, ['id', 'kind', 'slide', 'rules'], name)
  const rules = readArray(
    record.rules,
    `${name}.rules`,
    PRESENTATION_VERIFICATION_LIMITS.maxRenderRulesPerCheck,
    1,
  ).map((item, ruleIndex) => readEnum(item, renderRules, `${name}.rules[${ruleIndex}]`))
  if (new Set(rules).size !== rules.length) fail(`${name}.rules must be unique`)
  return { id, kind, slide, rules }
}

export const parsePresentationAcceptanceContract = (
  value: unknown,
): PresentationAcceptanceContract => {
  const record = readObject(value, 'contract')
  exactKeys(
    record,
    [
      'version',
      'taskId',
      'documentToken',
      'sessionToken',
      'baseRevision',
      'affectedSlides',
      'referenceSlides',
      'checks',
      'maxCorrectionPasses',
    ],
    'contract',
  )
  if (record.version !== 1) fail('contract.version must be 1')
  const affectedSlides = readSlides(
    record.affectedSlides,
    'contract.affectedSlides',
    PRESENTATION_VERIFICATION_LIMITS.maxAffectedSlides,
  )
  const referenceSlides = readSlides(
    record.referenceSlides,
    'contract.referenceSlides',
    PRESENTATION_VERIFICATION_LIMITS.maxReferenceSlides,
  )
  const checks = readArray(
    record.checks,
    'contract.checks',
    PRESENTATION_VERIFICATION_LIMITS.maxChecks,
    1,
  ).map(readCheck)
  if (new Set(checks.map((check) => check.id)).size !== checks.length)
    fail('contract check ids must be unique')
  const affected = new Set(affectedSlides)
  const references = new Set(referenceSlides)
  for (const check of checks) {
    if (!affected.has(check.slide)) fail(`check ${check.id} slide is outside affected scope`)
    if (check.kind === 'reference_match' && !references.has(check.referenceSlide))
      fail(`check ${check.id} reference slide is outside reference scope`)
  }
  return {
    version: 1,
    taskId: readIdentifier(record.taskId, 'contract.taskId'),
    documentToken: readIdentifier(record.documentToken, 'contract.documentToken'),
    sessionToken: readIdentifier(record.sessionToken, 'contract.sessionToken'),
    baseRevision: readDigest(record.baseRevision, 'contract.baseRevision'),
    affectedSlides,
    referenceSlides,
    checks,
    maxCorrectionPasses: readInteger(
      record.maxCorrectionPasses,
      'contract.maxCorrectionPasses',
      0,
      2,
    ) as 0 | 1 | 2,
  }
}

const canonicalContract = (contract: PresentationAcceptanceContract): unknown => ({
  ...contract,
  affectedSlides: [...contract.affectedSlides].sort((a, b) => a - b),
  referenceSlides: [...contract.referenceSlides].sort((a, b) => a - b),
  checks: [...contract.checks]
    .map((check) =>
      check.kind === 'reference_match'
        ? { ...check, properties: [...check.properties].sort() }
        : check.kind === 'render_quality'
          ? { ...check, rules: [...check.rules].sort() }
          : check,
    )
    .sort((a, b) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0)),
})

const canonicalJson = (value: unknown): string => {
  if (value === null || typeof value !== 'object') return JSON.stringify(value)
  if (Array.isArray(value)) return `[${value.map(canonicalJson).join(',')}]`
  return `{${Object.keys(value as Record<string, unknown>)
    .sort()
    .map(
      (key) => `${JSON.stringify(key)}:${canonicalJson((value as Record<string, unknown>)[key])}`,
    )
    .join(',')}}`
}

export const digestPresentationAcceptanceContract = async (value: unknown): Promise<string> => {
  const bytes = new TextEncoder().encode(
    canonicalJson(canonicalContract(parsePresentationAcceptanceContract(value))),
  )
  const digest = await globalThis.crypto.subtle.digest('SHA-256', bytes)
  return `sha256:${Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('')}`
}

export const parseVisualReviewResult = (value: unknown): VisualReviewResult => {
  const record = readObject(value, 'visualReview')
  exactKeys(record, ['status', 'failedCheckIds', 'observations', 'fixIntents'], 'visualReview')
  const status = readEnum(
    record.status,
    ['pass', 'needs_fix', 'cannot_verify'] as const,
    'visualReview.status',
  )
  const failedCheckIds = readUniqueIdentifiers(
    record.failedCheckIds,
    'visualReview.failedCheckIds',
    PRESENTATION_VERIFICATION_LIMITS.maxChecks,
  )
  const observations = readArray(
    record.observations,
    'visualReview.observations',
    PRESENTATION_VERIFICATION_LIMITS.maxObservations,
  ).map((item, index) => {
    const name = `visualReview.observations[${index}]`
    const observation = readObject(item, name)
    exactKeys(observation, ['code', 'severity', 'checkId', 'slide'], name)
    const checkId = Object.hasOwn(observation, 'checkId')
      ? readIdentifier(observation.checkId, `${name}.checkId`)
      : undefined
    const slide = Object.hasOwn(observation, 'slide')
      ? readInteger(observation.slide, `${name}.slide`, 1, 100_000)
      : undefined
    return {
      code: readEnum(
        observation.code,
        [
          'overflow',
          'overlap',
          'out_of_bounds',
          'illegible',
          'misaligned',
          'reference_mismatch',
          'review_unavailable',
        ] as const,
        `${name}.code`,
      ),
      severity: readEnum(
        observation.severity,
        ['info', 'warning', 'error'] as const,
        `${name}.severity`,
      ),
      ...(checkId === undefined ? {} : { checkId }),
      ...(slide === undefined ? {} : { slide }),
    }
  })
  const fixIntents = readArray(
    record.fixIntents,
    'visualReview.fixIntents',
    PRESENTATION_VERIFICATION_LIMITS.maxFixIntents,
  ).map((item, index) => {
    const name = `visualReview.fixIntents[${index}]`
    const intent = readObject(item, name)
    exactKeys(intent, ['checkId', 'kind', 'roleOrTarget', 'property', 'value'], name)
    if (intent.kind !== 'set_property') fail(`${name}.kind is unknown`)
    const property = readProperty(intent.property, `${name}.property`)
    return {
      checkId: readIdentifier(intent.checkId, `${name}.checkId`),
      kind: 'set_property' as const,
      roleOrTarget: readTarget(intent.roleOrTarget, `${name}.roleOrTarget`),
      property,
      value: readPropertyValue(intent.value, property, `${name}.value`),
    }
  })
  if (status === 'pass' && (failedCheckIds.length !== 0 || fixIntents.length !== 0))
    fail('visual pass must not report failures or fixes')
  if (status === 'needs_fix' && (failedCheckIds.length === 0 || fixIntents.length === 0))
    fail('needs_fix requires failures and fixes')
  if (status === 'cannot_verify' && fixIntents.length !== 0)
    fail('cannot_verify must not propose fixes')
  if (fixIntents.some((intent) => !failedCheckIds.includes(intent.checkId)))
    fail('fix intent is not bound to a failed check')
  return { status, failedCheckIds, observations, fixIntents }
}

export const parsePresentationRenderingFacts = (value: unknown): PresentationRenderingFacts => {
  const record = readObject(value, 'renderingFacts')
  exactKeys(
    record,
    ['contractDigest', 'revision', 'screenshots', 'deterministicResults'],
    'renderingFacts',
  )
  let totalBytes = 0
  const screenshots = readArray(
    record.screenshots,
    'renderingFacts.screenshots',
    PRESENTATION_VERIFICATION_LIMITS.maxScreenshotsPerPass,
  ).map((item, index) => {
    const name = `renderingFacts.screenshots[${index}]`
    const screenshot = readObject(item, name)
    exactKeys(screenshot, ['slide', 'role', 'roles', 'mediaToken', 'bytes'], name)
    const bytes = readInteger(
      screenshot.bytes,
      `${name}.bytes`,
      1,
      PRESENTATION_VERIFICATION_LIMITS.maxVisualRequestBytes,
    )
    totalBytes += bytes
    const role = readEnum(screenshot.role, ['affected', 'reference'] as const, `${name}.role`)
    const roles =
      screenshot.roles === undefined
        ? undefined
        : readArray(screenshot.roles, `${name}.roles`, 2).map((value, roleIndex) =>
            readEnum(value, ['affected', 'reference'] as const, `${name}.roles[${roleIndex}]`),
          )
    if (
      roles &&
      (roles.length < 1 || new Set(roles).size !== roles.length || !roles.includes(role))
    )
      fail(`${name}.roles must be unique and include role`)
    return {
      slide: readInteger(screenshot.slide, `${name}.slide`, 1, 100_000),
      role,
      ...(roles ? { roles } : {}),
      mediaToken: readIdentifier(screenshot.mediaToken, `${name}.mediaToken`),
      bytes,
    }
  })
  const deterministicResults = readArray(
    record.deterministicResults,
    'renderingFacts.deterministicResults',
    PRESENTATION_VERIFICATION_LIMITS.maxChecks,
  ).map((item, index) => {
    const name = `renderingFacts.deterministicResults[${index}]`
    const result = readObject(item, name)
    exactKeys(result, ['checkId', 'status', 'code'], name)
    const status = readEnum(
      result.status,
      ['pass', 'fail', 'unavailable'] as const,
      `${name}.status`,
    )
    const code = Object.hasOwn(result, 'code')
      ? readEnum(
          result.code,
          [
            'target_missing',
            'stale_revision',
            'unsupported_check',
            'value_mismatch',
            'scope_mismatch',
          ] as const,
          `${name}.code`,
        )
      : undefined
    if (status === 'pass' && code !== undefined) fail(`${name} pass must not include a code`)
    if (status !== 'pass' && code === undefined) fail(`${name} failure requires a code`)
    return {
      checkId: readIdentifier(result.checkId, `${name}.checkId`),
      status,
      ...(code === undefined ? {} : { code }),
    }
  })
  if (
    new Set(deterministicResults.map((result) => result.checkId)).size !==
    deterministicResults.length
  )
    fail('deterministic result check ids must be unique')
  const facts: PresentationRenderingFacts = {
    contractDigest: readDigest(record.contractDigest, 'renderingFacts.contractDigest'),
    revision: readDigest(record.revision, 'renderingFacts.revision'),
    screenshots,
    deterministicResults,
  }
  const serializedBytes = new TextEncoder().encode(JSON.stringify(facts)).byteLength
  if (serializedBytes + totalBytes > PRESENTATION_VERIFICATION_LIMITS.maxVisualRequestBytes)
    fail('renderingFacts serialized visual request is out of bounds')
  return facts
}

export const parsePresentationCompletionReceipt = (
  value: unknown,
  contractValue: unknown,
): PresentationCompletionReceipt => {
  const contract = parsePresentationAcceptanceContract(contractValue)
  const record = readObject(value, 'receipt')
  exactKeys(
    record,
    [
      'version',
      'taskId',
      'status',
      'mutationReceiptIds',
      'passedCheckIds',
      'failedCheckIds',
      'unavailableCheckIds',
      'correctionPasses',
      'affectedSlides',
      'rollbackId',
      'safeCode',
    ],
    'receipt',
  )
  if (record.version !== 1) fail('receipt.version must be 1')
  const status = readEnum(
    record.status,
    ['verified', 'applied_unverified', 'needs_user', 'failed', 'unchanged'] as const,
    'receipt.status',
  )
  const mutationReceiptIds = readUniqueIdentifiers(
    record.mutationReceiptIds,
    'receipt.mutationReceiptIds',
    PRESENTATION_VERIFICATION_LIMITS.maxReceiptIds,
  )
  const passedCheckIds = readUniqueIdentifiers(
    record.passedCheckIds,
    'receipt.passedCheckIds',
    PRESENTATION_VERIFICATION_LIMITS.maxChecks,
  )
  const failedCheckIds = readUniqueIdentifiers(
    record.failedCheckIds,
    'receipt.failedCheckIds',
    PRESENTATION_VERIFICATION_LIMITS.maxChecks,
  )
  const unavailableCheckIds = readUniqueIdentifiers(
    record.unavailableCheckIds,
    'receipt.unavailableCheckIds',
    PRESENTATION_VERIFICATION_LIMITS.maxChecks,
  )
  const allChecks = [...passedCheckIds, ...failedCheckIds, ...unavailableCheckIds]
  if (new Set(allChecks).size !== allChecks.length)
    fail('receipt check accounting must be disjoint')
  const expectedCheckIds = new Set(contract.checks.map((check) => check.id))
  if (
    allChecks.length !== expectedCheckIds.size ||
    allChecks.some((checkId) => !expectedCheckIds.has(checkId))
  )
    fail('receipt check accounting must exactly match the contract')
  const rollbackId = Object.hasOwn(record, 'rollbackId')
    ? readIdentifier(record.rollbackId, 'receipt.rollbackId')
    : undefined
  const safeCode = Object.hasOwn(record, 'safeCode')
    ? readEnum(
        record.safeCode,
        [
          'screenshot_unavailable',
          'review_unavailable',
          'verification_invalid',
          'cancelled_after_apply',
          'office_state_uncertain',
          'stale_authority',
          'confirmation_required',
          'unsupported_check',
          'cancelled',
          'mutation_failed',
          'visual_disabled',
        ] as const,
        'receipt.safeCode',
      )
    : undefined
  const taskId = readIdentifier(record.taskId, 'receipt.taskId')
  if (taskId !== contract.taskId) fail('receipt taskId does not match the contract')
  const affectedSlides = readSlides(
    record.affectedSlides,
    'receipt.affectedSlides',
    PRESENTATION_VERIFICATION_LIMITS.maxAffectedSlides,
  )
  const affectedSet = new Set(affectedSlides)
  if (
    affectedSet.size !== contract.affectedSlides.length ||
    contract.affectedSlides.some((slide) => !affectedSet.has(slide))
  )
    fail('receipt affected slides do not match the contract')
  const correctionPasses = readInteger(
    record.correctionPasses,
    'receipt.correctionPasses',
    0,
    contract.maxCorrectionPasses,
  )
  const unprovedCount = failedCheckIds.length + unavailableCheckIds.length
  if (correctionPasses > 0 && mutationReceiptIds.length === 0)
    fail('receipt correction passes require mutation evidence')
  if (rollbackId !== undefined && mutationReceiptIds.length === 0)
    fail('receipt rollback requires mutation evidence')
  if (status === 'verified' && (unprovedCount !== 0 || mutationReceiptIds.length === 0))
    fail('verified receipt requires every check passed and mutation evidence')
  if (status === 'verified' && safeCode !== undefined)
    fail('verified receipt must not contain a safeCode')
  if (
    status === 'applied_unverified' &&
    (mutationReceiptIds.length === 0 || unavailableCheckIds.length === 0)
  )
    fail('applied_unverified receipt requires mutation evidence and unavailable checks')
  if (
    status === 'applied_unverified' &&
    safeCode !== 'screenshot_unavailable' &&
    safeCode !== 'unsupported_check' &&
    safeCode !== 'review_unavailable' &&
    safeCode !== 'verification_invalid' &&
    safeCode !== 'stale_authority' &&
    safeCode !== 'cancelled_after_apply' &&
    safeCode !== 'visual_disabled'
  )
    fail('applied_unverified receipt has an incoherent status safeCode')
  if (status === 'failed' && (mutationReceiptIds.length !== 0 || unprovedCount === 0))
    fail('failed receipt requires no mutation evidence and at least one unproved check')
  if (
    status === 'failed' &&
    safeCode !== 'mutation_failed' &&
    safeCode !== 'cancelled' &&
    safeCode !== 'stale_authority' &&
    safeCode !== 'office_state_uncertain'
  )
    fail('failed receipt has an incoherent status safeCode')
  if (
    status === 'unchanged' &&
    (mutationReceiptIds.length !== 0 || unprovedCount !== 0 || correctionPasses !== 0)
  )
    fail('unchanged receipt requires all checks passed without mutations or corrections')
  if (status === 'unchanged' && safeCode !== undefined && safeCode !== 'cancelled')
    fail('unchanged receipt has an incoherent status safeCode')
  if (
    status === 'needs_user' &&
    (unprovedCount === 0 ||
      (safeCode !== 'confirmation_required' && safeCode !== 'unsupported_check'))
  )
    fail('needs_user receipt has an incoherent status safeCode or no unproved checks')
  return {
    version: 1,
    taskId,
    status,
    mutationReceiptIds,
    passedCheckIds,
    failedCheckIds,
    unavailableCheckIds,
    correctionPasses,
    affectedSlides,
    ...(rollbackId === undefined ? {} : { rollbackId }),
    ...(safeCode === undefined ? {} : { safeCode }),
  }
}

export type PresentationCompletionFacts = {
  status: PresentationCompletionReceipt['status']
  affectedSlides: number[]
  passedCount: number
  failedCount: number
  unavailableCount: number
  correctionPasses: number
  rollbackAvailable: boolean
  safeCode?: SafeCompletionCode
}

export const renderPresentationCompletionFacts = (
  value: unknown,
  contract: unknown,
): PresentationCompletionFacts => {
  const receipt = parsePresentationCompletionReceipt(value, contract)
  return {
    status: receipt.status,
    affectedSlides: [...receipt.affectedSlides],
    passedCount: receipt.passedCheckIds.length,
    failedCount: receipt.failedCheckIds.length,
    unavailableCount: receipt.unavailableCheckIds.length,
    correctionPasses: receipt.correctionPasses,
    rollbackAvailable: receipt.rollbackId !== undefined,
    ...(receipt.safeCode === undefined ? {} : { safeCode: receipt.safeCode }),
  }
}

export type PresentationGoldenCase = Readonly<{
  id: string
  scenario:
    | 'multi_slide_consistency'
    | 'already_correct'
    | 'ambiguous_emphasis'
    | 'screenshot_unavailable'
    | 'visual_safe_correction'
    | 'unsafe_fix'
    | 'cancel_pre'
    | 'cancel_post'
    | 'session_replacement'
    | 'stale_authority'
    | 'scope_expansion'
    | 'privacy'
  slideCount: number
}>

/** Scenario inputs only. Host outcomes must be derived by executing production adapters. */
export const PRESENTATION_GOLDEN_CASES: readonly PresentationGoldenCase[] = [
  {
    id: 'consistent-pages-6-8',
    scenario: 'multi_slide_consistency',
    slideCount: 8,
  },
  { id: 'already-correct', scenario: 'already_correct', slideCount: 1 },
  {
    id: 'ambiguous-emphasis',
    scenario: 'ambiguous_emphasis',
    slideCount: 6,
  },
  {
    id: 'capture-missing',
    scenario: 'screenshot_unavailable',
    slideCount: 1,
  },
  {
    id: 'one-safe-correction',
    scenario: 'visual_safe_correction',
    slideCount: 1,
  },
  { id: 'unsafe-fix', scenario: 'unsafe_fix', slideCount: 1 },
  { id: 'cancel-before-write', scenario: 'cancel_pre', slideCount: 1 },
  {
    id: 'cancel-after-write',
    scenario: 'cancel_post',
    slideCount: 1,
  },
  {
    id: 'session-replaced',
    scenario: 'session_replacement',
    slideCount: 1,
  },
  {
    id: 'authority-stale',
    scenario: 'stale_authority',
    slideCount: 1,
  },
  { id: 'scope-expanded', scenario: 'scope_expansion', slideCount: 1 },
  { id: 'private-content', scenario: 'privacy', slideCount: 1 },
]

export const PRESENTATION_CONSISTENCY_GOLDEN = Object.freeze({
  documentToken: 'golden-document',
  sessionToken: 'golden-session',
  baseRevision: `sha256:${'1'.repeat(64)}`,
  reference: { slide: 6, title: { x: 72, y: 48, width: 816, height: 54 } },
  colors: { title: '#2457A7', body: '#172033', emphasis: '#18A0A6' },
  pages: Array.from({ length: 8 }, (_, index) => ({
    slide: index + 1,
    shapes: ['title', 'body', 'emphasis'].map((role) => ({
      targetToken: `slide-${index + 1}:${role}`,
      role,
      color: '#000000',
      ...(role === 'title' ? { x: index + 1 === 6 ? 72 : 80, y: 48, width: 816, height: 54 } : {}),
    })),
  })),
  operations: [6, 7, 8].flatMap((slide) => [
    { slide, targetToken: `slide-${slide}:title`, property: 'color', value: '#2457A7' },
    { slide, targetToken: `slide-${slide}:body`, property: 'color', value: '#172033' },
    { slide, targetToken: `slide-${slide}:emphasis`, property: 'color', value: '#18A0A6' },
    ...(slide === 6
      ? []
      : [
          { slide, targetToken: `slide-${slide}:title`, property: 'x', value: 72 },
          { slide, targetToken: `slide-${slide}:title`, property: 'y', value: 48 },
          { slide, targetToken: `slide-${slide}:title`, property: 'width', value: 816 },
          { slide, targetToken: `slide-${slide}:title`, property: 'height', value: 54 },
        ]),
  ]),
  expectedCheckCount: 17,
})

export type PresentationTelemetryEvent = Readonly<{
  host: 'pc' | 'office'
  phase: 'plan' | 'dispatch' | 'deterministic' | 'visual' | 'correction' | 'complete'
  outcome: 'success' | 'unchanged' | 'needs_user' | 'unverified' | 'failed' | 'cancelled'
  code:
    | 'ready'
    | 'verified'
    | 'unchanged'
    | 'needs_user'
    | 'applied_unverified'
    | 'failed'
    | SafeCompletionCode
  count: number
  durationMs: number
}>

/** Exact allow-list: prompts, content, identifiers, receipts, fingerprints and images cannot pass. */
export function parsePresentationTelemetryEvent(value: unknown): PresentationTelemetryEvent {
  const record = readObject(value, 'telemetry event')
  const allowed = new Set(['host', 'phase', 'outcome', 'code', 'count', 'durationMs'])
  if (Object.keys(record).some((key) => !allowed.has(key)))
    fail('telemetry event contains private fields')
  const host = record.host
  const phase = record.phase
  const code = record.code
  const outcome = record.outcome
  const count = record.count
  const durationMs = record.durationMs
  if (host !== 'pc' && host !== 'office') fail('invalid telemetry host')
  if (
    !['plan', 'dispatch', 'deterministic', 'visual', 'correction', 'complete'].includes(
      String(phase),
    )
  )
    fail('invalid telemetry phase')
  const outcomes = ['success', 'unchanged', 'needs_user', 'unverified', 'failed', 'cancelled']
  if (!outcomes.includes(String(outcome))) fail('invalid telemetry outcome')
  const codes = [
    'ready',
    'verified',
    'unchanged',
    'needs_user',
    'applied_unverified',
    'failed',
    'screenshot_unavailable',
    'review_unavailable',
    'verification_invalid',
    'cancelled_after_apply',
    'office_state_uncertain',
    'stale_authority',
    'confirmation_required',
    'unsupported_check',
    'cancelled',
    'mutation_failed',
    'visual_disabled',
  ]
  if (!codes.includes(String(code))) fail('invalid telemetry code')
  if (!Number.isSafeInteger(count) || (count as number) < 0 || (count as number) > 1000)
    fail('invalid telemetry count')
  if (
    !Number.isSafeInteger(durationMs) ||
    (durationMs as number) < 0 ||
    (durationMs as number) > 3_600_000
  )
    fail('invalid telemetry duration')
  return {
    host,
    phase: phase as PresentationTelemetryEvent['phase'],
    outcome: outcome as PresentationTelemetryEvent['outcome'],
    code: code as PresentationTelemetryEvent['code'],
    count: count as number,
    durationMs: durationMs as number,
  }
}

export function presentationCompletionTelemetry(
  host: PresentationTelemetryEvent['host'],
  facts: Pick<PresentationCompletionFacts, 'status' | 'affectedSlides' | 'safeCode'>,
  durationMs: number,
): PresentationTelemetryEvent {
  const outcome =
    facts.status === 'verified'
      ? 'success'
      : facts.status === 'applied_unverified'
        ? 'unverified'
        : facts.status
  return parsePresentationTelemetryEvent({
    host,
    phase: 'complete',
    outcome,
    code: facts.safeCode ?? facts.status,
    count: facts.affectedSlides.length,
    durationMs: Math.max(0, Math.round(durationMs)),
  })
}

export function emitPresentationTelemetry(
  sink: ((event: PresentationTelemetryEvent) => void) | undefined,
  event: unknown,
): void {
  if (!sink) return
  try {
    sink(parsePresentationTelemetryEvent(event))
  } catch {
    // Diagnostics are observational and never alter mutation or receipt truth.
  }
}
