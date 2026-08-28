import { describe, expect, it, vi } from 'vitest'

import {
  completeOfficeRelayOAuthLogin,
  createOfficeRelayActivationFence,
  invalidateOfficeRelayBindingForCurrentAccount,
  lockOfficeRelayPersistence,
  logoutWithOfficeRelay,
  logoutOfficeRelaySession,
  officePairingResumeEnabled,
  shutdownOfficeRelaySession,
  startOfficeRelayPersistence,
  saveOfficeRelayBindingForCurrentAccount,
  syncOfficeRelaySession,
  syncOfficeRelayAccountSafely,
  terminalOfficeRelayAuthLoss,
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

  it('degrades to ordinary pairing and lets bootstrap continue when startup storage sync fails', async () => {
    const fallbackToOrdinary = vi.fn(async () => undefined)
    const diagnostic = vi.fn()
    await expect(
      startOfficeRelayPersistence({
        sync: async () => {
          throw new Error('secure_storage_unavailable')
        },
        fallbackToOrdinary,
        onBindingFailure: diagnostic,
      }),
    ).resolves.toBe(false)
    expect(fallbackToOrdinary).toHaveBeenCalledOnce()
    expect(diagnostic).toHaveBeenCalledOnce()
  })

  it('continues bootstrap even when diagnostic shutdown itself throws', async () => {
    await expect(
      startOfficeRelayPersistence({
        sync: async () => {
          throw new Error('invalid_schema')
        },
        fallbackToOrdinary: async () => {
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
    const fallbackToOrdinary = vi.fn(async () => undefined)
    const diagnostic = vi.fn()
    await expect(
      syncOfficeRelayAccountSafely({
        getAccountStatus: async () => account,
        syncBindingLifecycle: async () => {
          throw new Error('secure_storage_unavailable')
        },
        fallbackToOrdinary,
        onBindingFailure: diagnostic,
      }),
    ).resolves.toBe(account)
    expect(fallbackToOrdinary).toHaveBeenCalledWith(account)
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
        fallbackToOrdinary: vi.fn(async () => undefined),
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
        fallbackToOrdinary: vi.fn(async () => undefined),
        onBindingFailure: vi.fn(),
      }),
    ).rejects.toBe(authFailure)
  })

  it('closes ordinary one-time relay sessions without a persistent lifecycle', async () => {
    const suspend = vi.fn()
    const pool = { suspend }

    await logoutOfficeRelaySession(null, pool)
    await terminalOfficeRelayAuthLoss(null, pool)
    shutdownOfficeRelaySession(null, pool)

    expect(suspend.mock.calls).toEqual([['logout'], ['auth_required'], ['shutdown']])
  })

  it('delegates session closure to the durable lifecycle when it exists', async () => {
    const lifecycle = {
      syncAccount: vi.fn(async () => undefined),
      logout: vi.fn(async () => undefined),
      terminalAuthLoss: vi.fn(async () => undefined),
      shutdown: vi.fn(),
    }
    const suspend = vi.fn()
    const pool = { suspend }

    await logoutOfficeRelaySession(lifecycle, pool)
    await terminalOfficeRelayAuthLoss(lifecycle, pool)
    shutdownOfficeRelaySession(lifecycle, pool)

    expect(lifecycle.logout).toHaveBeenCalledOnce()
    expect(lifecycle.terminalAuthLoss).toHaveBeenCalledOnce()
    expect(lifecycle.shutdown).toHaveBeenCalledOnce()
    expect(suspend).not.toHaveBeenCalled()
  })

  it('reactivates flag-off ordinary pairing only after authenticated account sync', async () => {
    const pool = { suspend: vi.fn(), activate: vi.fn() }

    await syncOfficeRelaySession({ loggedIn: true, userId: 'account-one' }, false, null, pool)
    expect(pool.activate).toHaveBeenCalledOnce()
    expect(pool.suspend).not.toHaveBeenCalled()

    await syncOfficeRelaySession({ loggedIn: false }, false, null, pool)
    expect(pool.suspend).toHaveBeenCalledWith('auth_required')
  })

  it('uses durable account sync without activating ordinary mode while persistence is healthy', async () => {
    const lifecycle = {
      syncAccount: vi.fn(async () => undefined),
      logout: vi.fn(async () => undefined),
      terminalAuthLoss: vi.fn(async () => undefined),
      shutdown: vi.fn(),
    }
    const pool = { suspend: vi.fn(), activate: vi.fn() }

    await syncOfficeRelaySession({ loggedIn: true, userId: 'account-one' }, true, lifecycle, pool)
    expect(lifecycle.syncAccount).toHaveBeenCalledOnce()
    expect(pool.activate).not.toHaveBeenCalled()
  })

  it('locks failed logout persistence while keeping ordinary relay suspended until a fresh login sync', async () => {
    let persistenceAvailable = true
    const pool = { suspend: vi.fn(), activate: vi.fn() }
    lockOfficeRelayPersistence({
      disable: () => {
        persistenceAvailable = false
      },
      pool,
      reason: 'logout',
    })
    expect(persistenceAvailable).toBe(false)
    expect(pool.suspend).toHaveBeenCalledWith('logout')
    expect(pool.activate).not.toHaveBeenCalled()

    await syncOfficeRelaySession({ loggedIn: true, userId: 'fresh-account' }, false, null, pool)
    expect(pool.activate).toHaveBeenCalledOnce()
  })

  it('does not let a pre-logout account check reactivate ordinary pairing after the fence locks', async () => {
    const fence = createOfficeRelayActivationFence()
    const pool = { suspend: vi.fn(), activate: vi.fn() }
    let release!: (account: { loggedIn: boolean }) => void
    const account = new Promise<{ loggedIn: boolean }>((resolve) => {
      release = resolve
    })
    const stale = fence.settle({
      expectedEpoch: fence.snapshot(),
      allowUnlock: true,
      getValidAccountStatus: () => account,
      pool,
      suspendedReason: 'logout',
    })

    fence.lock()
    release({ loggedIn: true })
    await expect(stale).resolves.toBe(false)
    expect(pool.activate).not.toHaveBeenCalled()
    expect(pool.suspend).not.toHaveBeenCalled()
  })

  it('does not let an old account check suspend a newer OAuth activation', async () => {
    const fence = createOfficeRelayActivationFence()
    const pool = { suspend: vi.fn(), activate: vi.fn() }
    let releaseOld!: (account: { loggedIn: boolean }) => void
    const oldAccount = new Promise<{ loggedIn: boolean }>((resolve) => {
      releaseOld = resolve
    })
    const oldSettle = fence.settle({
      expectedEpoch: fence.snapshot(),
      allowUnlock: true,
      getValidAccountStatus: () => oldAccount,
      pool,
      suspendedReason: 'auth_required',
    })

    fence.lock()
    await expect(
      fence.settle({
        expectedEpoch: fence.snapshot(),
        allowUnlock: true,
        getValidAccountStatus: async () => ({ loggedIn: true }),
        pool,
        suspendedReason: 'auth_required',
      }),
    ).resolves.toBe(true)
    releaseOld({ loggedIn: true })
    await expect(oldSettle).resolves.toBe(false)

    expect(pool.activate).toHaveBeenCalledOnce()
    expect(pool.suspend).not.toHaveBeenCalled()
  })

  it('ignores an old account-check rejection after a newer OAuth activation', async () => {
    const fence = createOfficeRelayActivationFence()
    const pool = { suspend: vi.fn(), activate: vi.fn() }
    let rejectOld!: (error: Error) => void
    const oldAccount = new Promise<{ loggedIn: boolean }>((_, reject) => {
      rejectOld = reject
    })
    const oldSettle = fence.settle({
      expectedEpoch: fence.snapshot(),
      allowUnlock: true,
      getValidAccountStatus: () => oldAccount,
      pool,
      suspendedReason: 'auth_required',
    })

    fence.lock()
    await fence.settle({
      expectedEpoch: fence.snapshot(),
      allowUnlock: true,
      getValidAccountStatus: async () => ({ loggedIn: true }),
      pool,
      suspendedReason: 'auth_required',
    })
    rejectOld(new Error('stale_network_failure'))
    await expect(oldSettle).resolves.toBe(false)

    expect(pool.activate).toHaveBeenCalledOnce()
    expect(pool.suspend).not.toHaveBeenCalled()
  })

  it('keeps storage callbacks suspended and unlocks only after an authorized fresh account check', async () => {
    const fence = createOfficeRelayActivationFence()
    const pool = { suspend: vi.fn(), activate: vi.fn() }
    fence.lock()
    const expectedEpoch = fence.snapshot()

    await expect(
      fence.settle({
        expectedEpoch,
        allowUnlock: false,
        getValidAccountStatus: async () => ({ loggedIn: true }),
        pool,
        suspendedReason: 'binding_lifecycle',
      }),
    ).resolves.toBe(false)
    expect(pool.activate).not.toHaveBeenCalled()

    await expect(
      fence.settle({
        expectedEpoch,
        allowUnlock: true,
        getValidAccountStatus: async () => ({ loggedIn: true }),
        pool,
        suspendedReason: 'auth_required',
      }),
    ).resolves.toBe(true)
    expect(pool.activate).toHaveBeenCalledOnce()
  })

  it('turns a binding completion that races an account switch into an old-account tombstone', async () => {
    let accountId = 'old-account'
    let releasePut!: () => void
    const stored = new Promise<void>((resolve) => {
      releasePut = resolve
    })
    const put = vi.fn(() => stored)
    const tombstoneAccount = vi.fn(async () => undefined)
    const binding = {
      bindingId: 'binding_word_12345678',
      accountId: 'old-account',
      host: 'Word' as const,
      origin: 'https://office.8-216-134-194.sslip.io',
      capabilities: ['agent.v1'],
      createdAt: 1,
    }
    const saving = saveOfficeRelayBindingForCurrentAccount({
      binding,
      persistenceAvailable: () => true,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: accountId }),
      put,
      tombstoneAccount,
      onStorageFailure: vi.fn(async () => undefined),
    })

    await vi.waitFor(() => expect(put).toHaveBeenCalledWith(binding))
    accountId = 'replacement-account'
    releasePut()
    await expect(saving).rejects.toThrow('stale_binding_account')
    expect(put).toHaveBeenCalledWith(binding)
    expect(tombstoneAccount).toHaveBeenCalledWith('old-account')
  })

  it('tombstones a saved binding when the post-write account check fails', async () => {
    const binding = {
      bindingId: 'binding_word_12345678',
      accountId: 'old-account',
      host: 'Word' as const,
      origin: 'https://office.8-216-134-194.sslip.io',
      capabilities: ['agent.v1'],
      createdAt: 1,
    }
    const getValidAccountStatus = vi
      .fn()
      .mockResolvedValueOnce({ loggedIn: true, userId: 'old-account' })
      .mockRejectedValueOnce(new Error('auth_refresh_failed'))
    const tombstoneAccount = vi.fn(async () => undefined)

    await expect(
      saveOfficeRelayBindingForCurrentAccount({
        binding,
        persistenceAvailable: () => true,
        getValidAccountStatus,
        put: vi.fn(async () => undefined),
        tombstoneAccount,
        onStorageFailure: vi.fn(async () => undefined),
      }),
    ).rejects.toThrow('auth_refresh_failed')
    expect(tombstoneAccount).toHaveBeenCalledWith('old-account')
  })

  it('does not let a stale invalidation remove another account binding', async () => {
    const remove = vi.fn(async () => undefined)
    await invalidateOfficeRelayBindingForCurrentAccount({
      binding: {
        bindingId: 'binding_word_12345678',
        accountId: 'old-account',
        host: 'Word',
        origin: 'https://office.8-216-134-194.sslip.io',
        capabilities: ['agent.v1'],
        createdAt: 1,
      },
      persistenceAvailable: () => true,
      getValidAccountStatus: async () => ({
        loggedIn: true,
        userId: 'replacement-account',
      }),
      remove,
      onStorageFailure: vi.fn(async () => undefined),
    })
    expect(remove).not.toHaveBeenCalled()
  })
})
