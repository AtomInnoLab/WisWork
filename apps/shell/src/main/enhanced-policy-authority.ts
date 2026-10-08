import {
  ENHANCED_HOSTS,
  parseEnhancedCapabilities,
  parseEnhancedRolloutPolicy,
  type EnhancedHost,
  type EnhancedPolicySnapshot,
} from '@wiswork/agent-runtime'

declare const policyGrantBrand: unique symbol
export interface ShellEnhancedPolicyGrant {
  readonly [policyGrantBrand]: true
}

export function createShellEnhancedPolicyAuthority(currentGeneration: () => number) {
  const ledger = new WeakMap<object, EnhancedPolicySnapshot & { consumed: boolean }>()
  return Object.freeze({
    issue(value: {
      generation: number
      host: EnhancedHost
      policy: unknown
      capabilities: unknown
    }): ShellEnhancedPolicyGrant {
      const policy = parseEnhancedRolloutPolicy(value.policy)
      const capabilities = parseEnhancedCapabilities(value.capabilities)
      if (
        !Number.isSafeInteger(value.generation) ||
        value.generation < 0 ||
        currentGeneration() !== value.generation ||
        !ENHANCED_HOSTS.includes(value.host) ||
        !policy.globalEnabled ||
        !policy.hosts[value.host] ||
        (capabilities.includes('raw-office-proposal') &&
          (!value.host.startsWith('office-') || !policy.rawOfficeEnabled))
      )
        throw new Error('enhanced_policy_denied')
      const grant = Object.freeze(Object.create(null)) as ShellEnhancedPolicyGrant
      ledger.set(grant, {
        generation: value.generation,
        host: value.host,
        policy,
        capabilities: Object.freeze(capabilities),
        consumed: false,
      })
      return grant
    },
    consume(grant: ShellEnhancedPolicyGrant): EnhancedPolicySnapshot {
      const entry = ledger.get(grant as object)
      if (!entry) throw new Error('invalid_enhanced_policy_handle')
      if (entry.consumed) throw new Error('enhanced_policy_consumed')
      entry.consumed = true
      if (currentGeneration() !== entry.generation) throw new Error('enhanced_policy_stale')
      return Object.freeze({
        generation: entry.generation,
        host: entry.host,
        policy: entry.policy,
        capabilities: entry.capabilities,
      })
    },
  })
}
