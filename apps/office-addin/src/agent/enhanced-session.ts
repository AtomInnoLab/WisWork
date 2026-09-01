import type { OfficeHost } from '../office-document.js'

export interface OfficeEnhancedStatement {
  readonly version: 1
  readonly runtime_mode: 'enhanced'
  readonly runtime_instance: string
  readonly component_version: '0.147.0'
  readonly host: `office-${Exclude<OfficeHost, 'unknown'>}`
  readonly raw_office: boolean
  readonly expires_at: number
  readonly policy_generation: number
  readonly session_generation: number
}

export function rawOfficeCapabilities(statement: OfficeEnhancedStatement): Readonly<{
  rawJs: boolean
  rawOoxml: boolean
}> {
  return Object.freeze({
    rawJs: statement.raw_office,
    // Only PowerPoint has a reviewed bounded package transaction in this release.
    rawOoxml: statement.raw_office && statement.host === 'office-powerpoint',
  })
}

const exact = (value: Record<string, unknown>, keys: readonly string[]) =>
  Object.keys(value).length === keys.length && keys.every((key) => Object.hasOwn(value, key))

export function parseOfficeEnhancedStatement(value: unknown): OfficeEnhancedStatement {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new Error('invalid_enhanced_state')
  const record = value as Record<string, unknown>
  if (
    !exact(record, [
      'version',
      'runtime_mode',
      'runtime_instance',
      'component_version',
      'host',
      'raw_office',
      'expires_at',
      'policy_generation',
      'session_generation',
    ]) ||
    record.version !== 1 ||
    record.runtime_mode !== 'enhanced' ||
    typeof record.runtime_instance !== 'string' ||
    !/^[A-Za-z0-9_-]{16,128}$/.test(record.runtime_instance) ||
    record.component_version !== '0.147.0' ||
    !['office-word', 'office-excel', 'office-powerpoint'].includes(String(record.host)) ||
    typeof record.raw_office !== 'boolean' ||
    !Number.isSafeInteger(record.expires_at) ||
    Number(record.expires_at) <= 0 ||
    !Number.isSafeInteger(record.policy_generation) ||
    Number(record.policy_generation) < 0 ||
    !Number.isSafeInteger(record.session_generation) ||
    Number(record.session_generation) < 0
  )
    throw new Error('invalid_enhanced_state')
  return Object.freeze({ ...(record as unknown as OfficeEnhancedStatement) })
}
