import { describe, expect, it } from 'vitest'
import {
  ENHANCED_HOSTS,
  parseEnhancedComponentStatus,
  parseEnhancedRolloutPolicy,
  parseRuntimeSelection,
  shouldStartEnhancedRuntime,
  createEnhancedPolicyIssuer,
} from '../src/contracts'
import * as publicRuntime from '../src/index'

describe('runtime selection', () => {
  it('defaults exclusively to Standard', () => {
    const selection = parseRuntimeSelection(undefined)
    expect(selection).toEqual({ requested: 'standard', active: 'standard' })
    expect(
      shouldStartEnhancedRuntime(
        selection,
        {
          globalEnabled: true,
          hosts: Object.fromEntries(ENHANCED_HOSTS.map((host) => [host, true])) as Record<
            (typeof ENHANCED_HOSTS)[number],
            boolean
          >,
          rawOfficeEnabled: true,
        },
        'latex',
      ),
    ).toBe(false)
  })

  it('rejects unknown fields, symbols, prototypes, accessors, and oversized values', () => {
    const attacks: unknown[] = [
      { requested: 'standard', active: 'standard', surprise: true },
      Object.create({ requested: 'standard', active: 'standard' }),
      Object.defineProperty({ active: 'standard' }, 'requested', { get: () => 'enhanced' }),
      Object.assign({ requested: 'standard', active: 'standard' }, { [Symbol('x')]: true }),
      { requested: 'x'.repeat(65), active: 'standard' },
    ]
    for (const attack of attacks) expect(() => parseRuntimeSelection(attack)).toThrowError()
  })

  it('strictly parses component status', () => {
    expect(parseEnhancedComponentStatus('not_installed')).toBe('not_installed')
    expect(() => parseEnhancedComponentStatus('x'.repeat(65))).toThrowError()
    expect(() => parseEnhancedComponentStatus('launching')).toThrowError()
  })
})

describe('closed enhanced policy issuer', () => {
  it('does not expose the issuer from the package public root', () => {
    expect('createEnhancedPolicyIssuer' in publicRuntime).toBe(false)
  })
  it('issues opaque one-use grants bound to generation, host and capabilities', () => {
    let generation = 7
    const issuer = createEnhancedPolicyIssuer(() => generation)
    const policy = {
      globalEnabled: true,
      hosts: Object.fromEntries(ENHANCED_HOSTS.map((host) => [host, true])),
      rawOfficeEnabled: false,
    }
    const handle = issuer.issue({
      generation,
      host: 'docs',
      policy,
      capabilities: ['semantic-read'],
    })
    expect(Object.keys(handle)).toEqual([])
    expect(issuer.consume(handle)).toMatchObject({
      generation: 7,
      host: 'docs',
      capabilities: ['semantic-read'],
    })
    expect(() => issuer.consume(handle)).toThrow('enhanced_policy_consumed')

    const stale = issuer.issue({
      generation,
      host: 'docs',
      policy,
      capabilities: ['semantic-read'],
    })
    generation = 8
    expect(() => issuer.consume(stale)).toThrow('enhanced_policy_stale')
  })

  it('rejects cross-issuer, forged and raw Office policy escalation', () => {
    const a = createEnhancedPolicyIssuer(() => 1),
      b = createEnhancedPolicyIssuer(() => 1)
    const policy = {
      globalEnabled: true,
      hosts: Object.fromEntries(ENHANCED_HOSTS.map((host) => [host, true])),
      rawOfficeEnabled: false,
    }
    const handle = a.issue({ generation: 1, host: 'docs', policy, capabilities: ['semantic-read'] })
    expect(() => b.consume(handle)).toThrow('enhanced_policy_issuer_mismatch')
    expect(() => a.consume({} as never)).toThrow('invalid_enhanced_policy_handle')
    expect(() =>
      a.issue({
        generation: 1,
        host: 'office-word',
        policy,
        capabilities: ['raw-office-proposal'],
      }),
    ).toThrow('enhanced_policy_denied')
  })
})

describe('rollout policy', () => {
  it('defines all seven independently kill-switchable hosts and raw Office separately', () => {
    expect(ENHANCED_HOSTS).toEqual([
      'latex',
      'slides',
      'docs',
      'sheets',
      'office-word',
      'office-excel',
      'office-powerpoint',
    ])
    const hosts = Object.fromEntries(ENHANCED_HOSTS.map((host) => [host, host !== 'slides']))
    expect(
      parseEnhancedRolloutPolicy({ globalEnabled: true, hosts, rawOfficeEnabled: false }),
    ).toEqual({
      globalEnabled: true,
      hosts,
      rawOfficeEnabled: false,
    })
  })

  it('fails closed for partial, unknown, accessor, prototype, and oversized input', () => {
    const hosts = Object.fromEntries(ENHANCED_HOSTS.map((host) => [host, true]))
    const attacks: unknown[] = [
      { globalEnabled: true, hosts: { latex: true }, rawOfficeEnabled: true },
      { globalEnabled: true, hosts: { ...hosts, unknown: true }, rawOfficeEnabled: true },
      Object.assign(Object.create({ globalEnabled: true }), { hosts, rawOfficeEnabled: true }),
      Object.defineProperty({ hosts, rawOfficeEnabled: true }, 'globalEnabled', {
        get: () => true,
      }),
      { globalEnabled: true, hosts, rawOfficeEnabled: true, note: 'x'.repeat(10_000) },
    ]
    for (const attack of attacks) expect(() => parseEnhancedRolloutPolicy(attack)).toThrowError()
  })
})
