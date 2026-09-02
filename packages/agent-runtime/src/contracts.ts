export const ENHANCED_HOSTS = [
  'latex',
  'slides',
  'docs',
  'sheets',
  'office-word',
  'office-excel',
  'office-powerpoint',
] as const

export type EnhancedHost = (typeof ENHANCED_HOSTS)[number]
export const ENHANCED_CAPABILITIES = [
  'semantic-read',
  'transaction-proposal',
  'bounded-render-facts',
  'raw-office-proposal',
] as const
export type EnhancedCapability = (typeof ENHANCED_CAPABILITIES)[number]
export type AgentRuntimeMode = 'standard' | 'enhanced'
export type AgentRuntimeHost = EnhancedHost
export type RuntimeSelection = Readonly<{ requested: AgentRuntimeMode; active: AgentRuntimeMode }>
export type EnhancedRolloutPolicy = Readonly<{
  globalEnabled: boolean
  hosts: Readonly<Record<EnhancedHost, boolean>>
  rawOfficeEnabled: boolean
}>

export const COMPONENT_STATUSES = [
  'not_installed',
  'downloading',
  'verifying',
  'installed_restart_required',
  'ready',
  'update_available',
  'removal_restart_required',
  'blocked_by_policy',
  'unsupported_platform',
  'failed_safe',
] as const
export type EnhancedComponentStatus = (typeof COMPONENT_STATUSES)[number]

function strictRecord(value: unknown, keys: readonly string[]): Record<string, unknown> {
  if (value === undefined) return {}
  if (typeof value !== 'object' || value === null) throw new TypeError('invalid_contract')
  const prototype = Object.getPrototypeOf(value)
  if (prototype !== Object.prototype && prototype !== null) throw new TypeError('invalid_contract')
  if (Object.getOwnPropertySymbols(value).length > 0) throw new TypeError('invalid_contract')
  const descriptors = Object.getOwnPropertyDescriptors(value)
  if (Object.keys(descriptors).some((key) => !keys.includes(key)))
    throw new TypeError('invalid_contract')
  for (const descriptor of Object.values(descriptors)) {
    if (!('value' in descriptor)) throw new TypeError('invalid_contract')
  }
  return Object.fromEntries(
    Object.entries(descriptors).map(([key, descriptor]) => [key, descriptor.value]),
  )
}

function runtimeMode(value: unknown): AgentRuntimeMode {
  if (value !== 'standard' && value !== 'enhanced') throw new TypeError('invalid_runtime_mode')
  return value
}

export function parseRuntimeSelection(value: unknown): RuntimeSelection {
  if (value === undefined) return { requested: 'standard', active: 'standard' }
  const record = strictRecord(value, ['requested', 'active'])
  if (!Object.hasOwn(record, 'requested') || !Object.hasOwn(record, 'active')) {
    throw new TypeError('invalid_runtime_selection')
  }
  return { requested: runtimeMode(record.requested), active: runtimeMode(record.active) }
}

export function parseEnhancedRolloutPolicy(value: unknown): EnhancedRolloutPolicy {
  const record = strictRecord(value, ['globalEnabled', 'hosts', 'rawOfficeEnabled'])
  if (typeof record.globalEnabled !== 'boolean' || typeof record.rawOfficeEnabled !== 'boolean') {
    throw new TypeError('invalid_rollout_policy')
  }
  const hosts = strictRecord(record.hosts, ENHANCED_HOSTS)
  if (Object.keys(hosts).length !== ENHANCED_HOSTS.length)
    throw new TypeError('invalid_rollout_policy')
  const parsedHosts = Object.fromEntries(
    ENHANCED_HOSTS.map((host) => {
      if (typeof hosts[host] !== 'boolean') throw new TypeError('invalid_rollout_policy')
      return [host, hosts[host]]
    }),
  ) as Record<EnhancedHost, boolean>
  return {
    globalEnabled: record.globalEnabled,
    hosts: parsedHosts,
    rawOfficeEnabled: record.rawOfficeEnabled,
  }
}

export function parseEnhancedComponentStatus(value: unknown): EnhancedComponentStatus {
  if (
    typeof value !== 'string' ||
    value.length > 64 ||
    !COMPONENT_STATUSES.includes(value as EnhancedComponentStatus)
  ) {
    throw new TypeError('invalid_component_status')
  }
  return value as EnhancedComponentStatus
}

export function shouldStartEnhancedRuntime(
  selection: RuntimeSelection,
  policy: EnhancedRolloutPolicy,
  host: EnhancedHost,
): boolean {
  return selection.active === 'enhanced' && policy.globalEnabled && policy.hosts[host]
}

export interface EnhancedPolicySnapshot {
  readonly generation: number
  readonly host: EnhancedHost
  readonly policy: EnhancedRolloutPolicy
  readonly capabilities: readonly EnhancedCapability[]
}
export function parseEnhancedCapabilities(value: unknown): EnhancedCapability[] {
  if (
    !Array.isArray(value) ||
    Object.getPrototypeOf(value) !== Array.prototype ||
    value.length > ENHANCED_CAPABILITIES.length ||
    Object.getOwnPropertySymbols(value).length > 0
  )
    throw new TypeError('invalid_enhanced_policy')
  const descriptors = Object.getOwnPropertyDescriptors(value)
  const allowedKeys = new Set([
    'length',
    ...Array.from({ length: value.length }, (_, index) => String(index)),
  ])
  if (Object.keys(descriptors).some((key) => !allowedKeys.has(key)))
    throw new TypeError('invalid_enhanced_policy')
  const capabilities: EnhancedCapability[] = []
  for (let index = 0; index < value.length; index += 1) {
    const descriptor = descriptors[String(index)]
    if (
      !descriptor ||
      !('value' in descriptor) ||
      !ENHANCED_CAPABILITIES.includes(descriptor.value as EnhancedCapability) ||
      capabilities.includes(descriptor.value as EnhancedCapability)
    )
      throw new TypeError('invalid_enhanced_policy')
    capabilities.push(descriptor.value as EnhancedCapability)
  }
  return capabilities
}
