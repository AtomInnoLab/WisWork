import { describe, expect, it } from 'vitest'
import {
  ALLOWED_ENHANCED_CAPABILITIES,
  DENIED_ENHANCED_CAPABILITIES,
  parseEnhancedCapabilityAuthorization,
  parseCapabilityDeclaration,
  parseSafeRuntimeError,
} from '../src'

describe('enhanced capability policy', () => {
  it('architecturally denies ambient and bypass capabilities', () => {
    expect(DENIED_ENHANCED_CAPABILITIES).toEqual([
      'shell',
      'arbitrary-filesystem',
      'git',
      'browser-control',
      'free-network',
      'direct-document-write',
    ])
    expect(ALLOWED_ENHANCED_CAPABILITIES).not.toEqual(
      expect.arrayContaining([...DENIED_ENHANCED_CAPABILITIES]),
    )
  })

  it('accepts only exact, bounded, data-only declarations', () => {
    expect(
      parseCapabilityDeclaration({ capabilities: ['semantic-read', 'transaction-proposal'] }),
    ).toEqual({ capabilities: ['semantic-read', 'transaction-proposal'] })

    for (const attack of [
      { capabilities: ['shell'] },
      { capabilities: ['semantic-read'], extra: true },
      { capabilities: Array(33).fill('semantic-read') },
      { capabilities: ['x'.repeat(65)] },
      Object.assign(Object.create({ capabilities: [] }), {}),
      Object.defineProperty({}, 'capabilities', { get: () => [] }),
      { capabilities: Object.defineProperty([], '0', { get: () => 'semantic-read' }) },
      { capabilities: Object.assign(['semantic-read'], { '01': 'semantic-read' }) },
      { capabilities: Object.assign(['semantic-read'], { '00': 'semantic-read' }) },
      { capabilities: Object.assign(['semantic-read'], { '2': 'semantic-read' }) },
      Object.assign({ capabilities: [] }, { [Symbol('x')]: true }),
    ])
      expect(() => parseCapabilityDeclaration(attack)).toThrowError()
  })

  it('authorizes capabilities only for an enabled host and keeps raw Office Office-only', () => {
    const hosts = {
      latex: true,
      slides: true,
      docs: true,
      sheets: true,
      'office-word': true,
      'office-excel': true,
      'office-powerpoint': true,
    }
    const policy = { globalEnabled: true, hosts, rawOfficeEnabled: true }
    expect(
      parseEnhancedCapabilityAuthorization({
        host: 'office-word',
        policy,
        declaration: { capabilities: ['semantic-read', 'raw-office-proposal'] },
      }),
    ).toEqual({
      host: 'office-word',
      policy,
      declaration: { capabilities: ['semantic-read', 'raw-office-proposal'] },
    })

    for (const attack of [
      { host: 'latex', policy, declaration: { capabilities: ['raw-office-proposal'] } },
      {
        host: 'office-word',
        policy: { ...policy, rawOfficeEnabled: false },
        declaration: { capabilities: ['raw-office-proposal'] },
      },
      {
        host: 'office-word',
        policy: { ...policy, hosts: { ...hosts, 'office-word': false } },
        declaration: { capabilities: ['semantic-read'] },
      },
      {
        host: 'office-word',
        policy: { ...policy, globalEnabled: false },
        declaration: { capabilities: ['semantic-read'] },
      },
    ])
      expect(() => parseEnhancedCapabilityAuthorization(attack)).toThrowError()
  })

  it('parses safe errors without retaining unknown or oversized details', () => {
    expect(parseSafeRuntimeError({ code: 'runtime_unavailable' })).toEqual({
      code: 'runtime_unavailable',
    })
    for (const attack of [
      { code: 'runtime_unavailable', prompt: 'secret' },
      { code: 'x'.repeat(65) },
      Object.defineProperty({}, 'code', { get: () => 'runtime_unavailable' }),
    ])
      expect(() => parseSafeRuntimeError(attack)).toThrowError()
  })
})
