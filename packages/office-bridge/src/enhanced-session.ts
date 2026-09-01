export type OfficeEnhancedHost = 'office-word' | 'office-excel' | 'office-powerpoint'

export interface OfficeEnhancedSessionStatement {
  readonly version: 1
  readonly runtime_mode: 'enhanced'
  readonly runtime_instance: string
  readonly component_version: string
  readonly host: OfficeEnhancedHost
  readonly raw_office: boolean
  readonly expires_at: number
  readonly policy_generation: number
  readonly session_generation: number
}

export interface OfficeRuntimeSelectionContext {
  readonly host: OfficeEnhancedHost
  readonly now: number
  readonly componentVersion: string
  readonly sessionGeneration: number
  readonly minimumPolicyGeneration: number
}

export type OfficeRuntimeSelection =
  | Readonly<{ mode: 'standard'; rawOfficeAllowed: false }>
  | Readonly<{ mode: 'enhanced'; rawOfficeAllowed: boolean; runtimeInstance: string }>

const KEYS = [
  'version',
  'runtime_mode',
  'runtime_instance',
  'component_version',
  'host',
  'raw_office',
  'expires_at',
  'policy_generation',
  'session_generation',
] as const
const HOSTS = new Set<OfficeEnhancedHost>(['office-word', 'office-excel', 'office-powerpoint'])
const opaque = (value: unknown): value is string =>
  typeof value === 'string' && /^[A-Za-z0-9_-]{16,128}$/.test(value)
const generation = (value: unknown): value is number =>
  Number.isSafeInteger(value) && (value as number) >= 0

export function parseOfficeEnhancedSessionStatement(
  value: unknown,
): Readonly<OfficeEnhancedSessionStatement> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error('invalid_office_enhanced_statement')
  }
  const record = value as Record<string, unknown>
  if (
    Object.keys(record).length !== KEYS.length ||
    !KEYS.every((key) => Object.hasOwn(record, key)) ||
    record.version !== 1 ||
    record.runtime_mode !== 'enhanced' ||
    !opaque(record.runtime_instance) ||
    typeof record.component_version !== 'string' ||
    !/^\d+\.\d+\.\d+$/.test(record.component_version) ||
    !HOSTS.has(record.host as OfficeEnhancedHost) ||
    typeof record.raw_office !== 'boolean' ||
    !Number.isSafeInteger(record.expires_at) ||
    (record.expires_at as number) <= 0 ||
    !generation(record.policy_generation) ||
    !generation(record.session_generation)
  ) {
    throw new Error('invalid_office_enhanced_statement')
  }
  return Object.freeze({ ...(record as unknown as OfficeEnhancedSessionStatement) })
}

export function selectOfficeAgentRuntime(
  value: unknown,
  context: OfficeRuntimeSelectionContext,
): OfficeRuntimeSelection {
  if (!value) return Object.freeze({ mode: 'standard', rawOfficeAllowed: false })
  let statement: Readonly<OfficeEnhancedSessionStatement>
  try {
    statement = parseOfficeEnhancedSessionStatement(value)
  } catch {
    return Object.freeze({ mode: 'standard', rawOfficeAllowed: false })
  }
  if (
    statement.host !== context.host ||
    statement.expires_at <= context.now ||
    statement.component_version !== context.componentVersion ||
    statement.session_generation !== context.sessionGeneration ||
    statement.policy_generation < context.minimumPolicyGeneration
  ) {
    return Object.freeze({ mode: 'standard', rawOfficeAllowed: false })
  }
  return Object.freeze({
    mode: 'enhanced',
    rawOfficeAllowed: statement.raw_office,
    runtimeInstance: statement.runtime_instance,
  })
}
