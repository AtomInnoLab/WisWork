import { ENHANCED_HOSTS } from '@wiswork/agent-runtime'
import { describe, expect, it } from 'vitest'
import { createShellEnhancedPolicyAuthority } from '../src/main/enhanced-policy-authority'

describe('Shell Enhanced policy authority', () => {
  it('issues one-use opaque grants and revokes them on generation change', () => {
    let generation = 1
    const authority = createShellEnhancedPolicyAuthority(() => generation)
    const policy = {
      globalEnabled: true,
      rawOfficeEnabled: false,
      hosts: Object.fromEntries(ENHANCED_HOSTS.map((host) => [host, true])),
    }
    const grant = authority.issue({
      generation,
      host: 'docs',
      policy,
      capabilities: ['semantic-read'],
    })
    expect(Object.keys(grant)).toEqual([])
    expect(authority.consume(grant)).toMatchObject({ generation: 1, host: 'docs' })
    expect(() => authority.consume(grant)).toThrow('enhanced_policy_consumed')
    const stale = authority.issue({
      generation,
      host: 'docs',
      policy,
      capabilities: ['semantic-read'],
    })
    generation += 1
    expect(() => authority.consume(stale)).toThrow('enhanced_policy_stale')
  })
})
