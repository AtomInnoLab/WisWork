import { describe, expect, it, vi } from 'vitest'

import {
  completeOfficeRelayOAuthLogin,
  logoutWithOfficeRelay,
  officePairingResumeEnabled,
  startOfficeRelayPersistence,
  syncOfficeRelayAccountSafely,
} from '../src/main/office-relay-runtime'

describe('Office relay runtime integration', () => {
  it('parses the exact persistent-pairing environment kill switch fail closed', () => {
    expect(officePairingResumeEnabled({})).toBe(true)
    expect(officePairingResumeEnabled({ WISWORK_OFFICE_PAIRING_RESUME: '1' })).toBe(true)
    expect(officePairingResumeEnabled({ WISWORK_OFFICE_PAIRING_RESUME: '0' })).toBe(false)
    for (const value of ['', 'false', 'true', '2'])
      expect(() => officePairingResumeEnabled({ WISWORK_OFFICE_PAIRING_RESUME: value })).toThrow(
        'invalid_office_pairing_resume',
      )
  })

  it('always performs auth logout after binding invalidation fails and reports the binding diagnostic', async () => {
    const order: string[] = []
    let suspended = false
    const diagnostic = vi.fn()
    const result = await logoutWithOfficeRelay({
      invalidate: async () => {
        suspended = true
        order.push('invalidate')
        throw new Error('disk_failure')
      },
      logout: async () => {
        expect(suspended).toBe(true)
        order.push('logout')
        return true
      },
      onBindingFailure: diagnostic,
    })
    expect(result).toBe(true)
    expect(order).toEqual(['invalidate', 'logout'])
    expect(diagnostic).toHaveBeenCalledOnce()
  })

  it('suspends persistent relay and lets bootstrap continue when startup storage sync fails', async () => {
    const suspend = vi.fn()
    const diagnostic = vi.fn()
    await expect(
      startOfficeRelayPersistence({
        sync: async () => {
          throw new Error('secure_storage_unavailable')
        },
        suspend,
        onBindingFailure: diagnostic,
      }),
    ).resolves.toBe(false)
    expect(suspend).toHaveBeenCalledWith('binding_lifecycle')
    expect(diagnostic).toHaveBeenCalledOnce()
  })

  it('continues bootstrap even when diagnostic shutdown itself throws', async () => {
    await expect(
      startOfficeRelayPersistence({
        sync: async () => {
          throw new Error('invalid_schema')
        },
        suspend: () => {
          throw new Error('partial_pool')
        },
        onBindingFailure: () => {
          throw new Error('renderer_gone')
        },
      }),
    ).resolves.toBe(false)
  })

  it('returns the real account status while isolating accountStatus binding storage failure', async () => {
    const account = { loggedIn: true, userId: 'account-one' }
    const suspend = vi.fn()
    const diagnostic = vi.fn()
    await expect(
      syncOfficeRelayAccountSafely({
        getAccountStatus: async () => account,
        syncBindingLifecycle: async () => {
          throw new Error('secure_storage_unavailable')
        },
        suspend,
        onBindingFailure: diagnostic,
      }),
    ).resolves.toBe(account)
    expect(suspend).toHaveBeenCalledWith('binding_lifecycle')
    expect(diagnostic).toHaveBeenCalledOnce()
  })

  it('keeps OAuth login successful after binding sync failure but propagates auth failure', async () => {
    const account = { loggedIn: true, userId: 'account-two' }
    const consumeCallback = vi.fn(async () => undefined)
    const getAccountStatus = vi.fn(async () => account)
    const diagnostic = vi.fn()
    await expect(
      completeOfficeRelayOAuthLogin({
        consumeCallback,
        getAccountStatus,
        syncBindingLifecycle: async () => {
          throw new Error('invalid_binding_schema')
        },
        suspend: vi.fn(),
        onBindingFailure: diagnostic,
      }),
    ).resolves.toBe(account)
    expect(consumeCallback).toHaveBeenCalledOnce()
    expect(diagnostic).toHaveBeenCalledOnce()

    const authFailure = new Error('invalid_oauth_state')
    await expect(
      completeOfficeRelayOAuthLogin({
        consumeCallback: async () => {
          throw authFailure
        },
        getAccountStatus,
        syncBindingLifecycle: vi.fn(async () => undefined),
        suspend: vi.fn(),
        onBindingFailure: vi.fn(),
      }),
    ).rejects.toBe(authFailure)
  })
})
