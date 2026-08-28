import { describe, expect, it } from 'vitest'

import { sanitizeOfficeRelayStatus } from '../src/preload/office-relay-status'

describe('Office relay preload status validation', () => {
  it.each(['disconnected:binding_not_remembered', 'error:binding_lifecycle'] as const)(
    'allows the persistent pairing diagnostic %s',
    (status) => expect(sanitizeOfficeRelayStatus(status)).toBe(status),
  )

  it('fails closed for unknown renderer input', () => {
    expect(sanitizeOfficeRelayStatus('attacker-status')).toBe('disconnected')
  })
})
