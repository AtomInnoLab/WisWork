import { describe, expect, it } from 'vitest'
import {
  ALLOWED_ENHANCED_CAPABILITIES,
  DENIED_ENHANCED_CAPABILITIES,
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
      Object.assign({ capabilities: [] }, { [Symbol('x')]: true }),
    ])
      expect(() => parseCapabilityDeclaration(attack)).toThrowError()
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
