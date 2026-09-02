import { describe, expect, it } from 'vitest'
import {
  ENHANCED_HOSTS,
  parseEnhancedComponentStatus,
  parseEnhancedRolloutPolicy,
  parseRuntimeSelection,
  shouldStartEnhancedRuntime,
} from '../src/contracts'

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

describe('seven-host emergency rollback', () => {
  const hosts = Object.fromEntries(ENHANCED_HOSTS.map((host) => [host, true])) as Record<
    (typeof ENHANCED_HOSTS)[number],
    boolean
  >
  const selection = { requested: 'enhanced', active: 'enhanced' } as const

  it('global and each independent host switch restore the Standard path', () => {
    const enabled = { globalEnabled: true, hosts, rawOfficeEnabled: true }
    for (const host of ENHANCED_HOSTS) {
      expect(shouldStartEnhancedRuntime(selection, enabled, host)).toBe(true)
      expect(
        shouldStartEnhancedRuntime(
          selection,
          { ...enabled, hosts: { ...hosts, [host]: false } },
          host,
        ),
      ).toBe(false)
    }
    for (const host of ENHANCED_HOSTS)
      expect(
        shouldStartEnhancedRuntime(selection, { ...enabled, globalEnabled: false }, host),
      ).toBe(false)
  })

  it('raw Office rollback is independent from ordinary Enhanced host availability', () => {
    const policy = parseEnhancedRolloutPolicy({
      globalEnabled: true,
      hosts,
      rawOfficeEnabled: false,
    })
    for (const host of ['office-word', 'office-excel', 'office-powerpoint'] as const)
      expect(shouldStartEnhancedRuntime(selection, policy, host)).toBe(true)
    expect(policy.rawOfficeEnabled).toBe(false)
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
