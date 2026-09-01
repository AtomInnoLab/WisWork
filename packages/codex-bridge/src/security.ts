import {
  ALLOWED_ENHANCED_CAPABILITIES,
  SAFE_RUNTIME_ERROR_CODES,
  type CapabilityDeclaration,
  type EnhancedCapability,
  type SafeRuntimeError,
} from './types'
import {
  ENHANCED_HOSTS,
  parseEnhancedRolloutPolicy,
  type EnhancedHost,
  type EnhancedRolloutPolicy,
} from '@wiswork/agent-runtime'

const MAX_CAPABILITIES = 32
const MAX_TOKEN_LENGTH = 64

function strictRecord(value: unknown, keys: readonly string[]): Record<string, unknown> {
  if (typeof value !== 'object' || value === null) throw new TypeError('invalid_runtime_input')
  const prototype = Object.getPrototypeOf(value)
  if (prototype !== Object.prototype && prototype !== null)
    throw new TypeError('invalid_runtime_input')
  if (Object.getOwnPropertySymbols(value).length > 0) throw new TypeError('invalid_runtime_input')
  const descriptors = Object.getOwnPropertyDescriptors(value)
  if (Object.keys(descriptors).some((key) => !keys.includes(key)))
    throw new TypeError('invalid_runtime_input')
  for (const descriptor of Object.values(descriptors)) {
    if (!('value' in descriptor)) throw new TypeError('invalid_runtime_input')
  }
  return Object.fromEntries(
    Object.entries(descriptors).map(([key, descriptor]) => [key, descriptor.value]),
  )
}

function strictArray(value: unknown, maximumLength: number): unknown[] {
  if (!Array.isArray(value) || Object.getPrototypeOf(value) !== Array.prototype) {
    throw new TypeError('invalid_runtime_input')
  }
  if (value.length > maximumLength || Object.getOwnPropertySymbols(value).length > 0) {
    throw new TypeError('invalid_runtime_input')
  }
  const descriptors = Object.getOwnPropertyDescriptors(value)
  for (let index = 0; index < value.length; index += 1) {
    const descriptor = descriptors[String(index)]
    if (!descriptor || !('value' in descriptor)) throw new TypeError('invalid_runtime_input')
  }
  const allowedKeys = new Set([
    'length',
    ...Array.from({ length: value.length }, (_, index) => String(index)),
  ])
  if (Object.keys(descriptors).some((key) => !allowedKeys.has(key))) {
    throw new TypeError('invalid_runtime_input')
  }
  return Array.from({ length: value.length }, (_, index) => descriptors[String(index)].value)
}

export type EnhancedCapabilityAuthorization = Readonly<{
  host: EnhancedHost
  policy: EnhancedRolloutPolicy
  declaration: CapabilityDeclaration
}>

export function parseEnhancedCapabilityAuthorization(
  value: unknown,
): EnhancedCapabilityAuthorization {
  const record = strictRecord(value, ['host', 'policy', 'declaration'])
  if (
    typeof record.host !== 'string' ||
    record.host.length > MAX_TOKEN_LENGTH ||
    !ENHANCED_HOSTS.includes(record.host as EnhancedHost)
  )
    throw new TypeError('invalid_enhanced_host')
  const host = record.host as EnhancedHost
  const policy = parseEnhancedRolloutPolicy(record.policy)
  const declaration = parseCapabilityDeclaration(record.declaration)
  if (!policy.globalEnabled || !policy.hosts[host]) throw new TypeError('capability_not_authorized')
  if (
    declaration.capabilities.includes('raw-office-proposal') &&
    (!host.startsWith('office-') || !policy.rawOfficeEnabled)
  )
    throw new TypeError('capability_not_authorized')
  return { host, policy, declaration }
}

export function parseCapabilityDeclaration(value: unknown): CapabilityDeclaration {
  const record = strictRecord(value, ['capabilities'])
  const capabilities = strictArray(record.capabilities, MAX_CAPABILITIES).map((capability) => {
    if (
      typeof capability !== 'string' ||
      capability.length > MAX_TOKEN_LENGTH ||
      !ALLOWED_ENHANCED_CAPABILITIES.includes(capability as EnhancedCapability)
    )
      throw new TypeError('invalid_capability')
    return capability as EnhancedCapability
  })
  return { capabilities }
}

export function parseSafeRuntimeError(value: unknown): SafeRuntimeError {
  const record = strictRecord(value, ['code'])
  if (
    typeof record.code !== 'string' ||
    record.code.length > MAX_TOKEN_LENGTH ||
    !SAFE_RUNTIME_ERROR_CODES.includes(record.code as SafeRuntimeError['code'])
  )
    throw new TypeError('invalid_runtime_error')
  return { code: record.code as SafeRuntimeError['code'] }
}
