import { describe, expect, it, vi } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'

import type {
  OfficeRelayBinding,
  OfficeRelayBindingStore,
  OfficeRelayBindingTombstone,
} from '../src/main/office-relay-binding-store'
import type { OfficeRelayPool } from '../src/main/office-relay-pool'
import { createOfficeRelayLifecycle } from '../src/main/office-relay-lifecycle'

const makeBinding = (accountId: string, suffix: string): OfficeRelayBinding => ({
  bindingId: `binding_${suffix}_12345678`,
  accountId,
  host: 'Word',
  origin: 'https://office.8-216-134-194.sslip.io',
  capabilities: ['agent.v1'],
  createdAt: 1,
})

function harness() {
  let account: { loggedIn: boolean; userId?: string } = {
    loggedIn: true,
    userId: 'account-one',
  }
  const bindings = new Map<string, OfficeRelayBinding[]>([
    ['account-one', [makeBinding('account-one', 'one'), makeBinding('account-one', 'two')]],
    ['account-two', [makeBinding('account-two', 'three')]],
  ])
  const tombstones = new Map<string, OfficeRelayBindingTombstone[]>([
    [
      'account-one',
      [{ bindingId: 'binding_stale_12345678', accountId: 'account-one', createdAt: 1 }],
    ],
  ])
  const order: string[] = []
  const store: OfficeRelayBindingStore = {
    listForAccount: vi.fn(async (accountId) => [...(bindings.get(accountId) ?? [])]),
    listTombstonesForAccount: vi.fn(async (accountId) => [...(tombstones.get(accountId) ?? [])]),
    put: vi.fn(async () => undefined),
    remove: vi.fn(async () => undefined),
    tombstoneAccount: vi.fn(async (accountId) => {
      order.push(`tombstone:${accountId}`)
      const active = bindings.get(accountId) ?? []
      bindings.set(accountId, [])
      tombstones.set(accountId, [
        ...(tombstones.get(accountId) ?? []),
        ...active.map(({ bindingId }) => ({ bindingId, accountId, createdAt: 2 })),
      ])
    }),
    acknowledgeTombstone: vi.fn(async (accountId, bindingId) => {
      order.push(`ack:${bindingId}`)
      tombstones.set(
        accountId,
        (tombstones.get(accountId) ?? []).filter((entry) => entry.bindingId !== bindingId),
      )
    }),
  }
  const pool = {
    claim: vi.fn(async () => undefined),
    resume: vi.fn(async (binding: OfficeRelayBinding) => {
      order.push(`resume:${binding.bindingId}`)
    }),
    revokeBinding: vi.fn(async (bindingId: string, expectedAccountId: string) => {
      order.push(`revoke:${bindingId}:${expectedAccountId}`)
    }),
    approve: vi.fn(async () => false),
    reject: vi.fn(() => false),
    listPending: vi.fn(() => []),
    status: vi.fn(() => 'disconnected' as const),
    revoke: vi.fn((reason?: string) => order.push(`disconnect:${reason}`)),
    suspend: vi.fn((reason?: string) => order.push(`suspend:${reason}`)),
    activate: vi.fn(() => order.push('activate')),
  } satisfies OfficeRelayPool
  const lifecycle = createOfficeRelayLifecycle({
    store,
    pool,
    getValidAccountStatus: async () => account,
  })
  return {
    lifecycle,
    store,
    pool,
    order,
    tombstones,
    setAccount(next: typeof account) {
      account = next
    },
  }
}

describe('Office relay binding lifecycle', () => {
  it('captures the local account before terminal startup refresh loss clears auth', async () => {
    const { store, pool } = harness()
    const lifecycle = createOfficeRelayLifecycle({
      store,
      pool,
      getAccountStatus: async () => ({ loggedIn: true, userId: 'account-one' }),
      getValidAccountStatus: async () => ({ loggedIn: false }),
    })
    await lifecycle.syncAccount()
    expect(store.tombstoneAccount).toHaveBeenCalledWith('account-one')
    expect(pool.suspend).toHaveBeenCalledWith('auth_required')
  })

  it('retries same-account tombstones before resuming every active binding at startup', async () => {
    const { lifecycle, order, tombstones } = harness()
    await lifecycle.syncAccount()
    expect(order).toEqual([
      'revoke:binding_stale_12345678:account-one',
      'ack:binding_stale_12345678',
      'activate',
      'resume:binding_one_12345678',
      'resume:binding_two_12345678',
    ])
    expect(tombstones.get('account-one')).toEqual([])
  })

  it('starts resume work concurrently while enforcing the twelve-slot bound', async () => {
    const { lifecycle, store, pool } = harness()
    ;(store.listForAccount as ReturnType<typeof vi.fn>).mockResolvedValue(
      Array.from({ length: 13 }, (_, index) => makeBinding('account-one', `parallel_${index}`)),
    )
    const releases: Array<() => void> = []
    ;(pool.resume as ReturnType<typeof vi.fn>).mockImplementation(
      () =>
        new Promise<void>((resolve) => {
          releases.push(resolve)
        }),
    )

    const syncing = lifecycle.syncAccount()
    await vi.waitFor(() => expect(pool.resume).toHaveBeenCalledTimes(12))
    expect(releases).toHaveLength(12)
    releases.forEach((release) => release())
    await syncing
  })

  it('does not activate or resume when an auth epoch becomes stale during account sync', async () => {
    const { lifecycle, store, pool } = harness()
    ;(store.listTombstonesForAccount as ReturnType<typeof vi.fn>).mockResolvedValue([])
    let release!: () => void
    ;(store.listForAccount as ReturnType<typeof vi.fn>).mockImplementation(
      () =>
        new Promise<OfficeRelayBinding[]>((resolve) => {
          release = () => resolve([makeBinding('account-one', 'delayed')])
        }),
    )
    let current = true
    const syncing = lifecycle.syncAccount(() => current)
    await vi.waitFor(() => expect(store.listForAccount).toHaveBeenCalledOnce())
    current = false
    release()
    await syncing

    expect(pool.activate).not.toHaveBeenCalled()
    expect(pool.resume).not.toHaveBeenCalled()
  })

  it('drains every tombstone with at most twelve revocations in flight and retains failures', async () => {
    const { lifecycle, store, pool } = harness()
    const records = Array.from({ length: 256 }, (_, index) => ({
      bindingId: `binding_bulk_${index}_12345678`,
      accountId: 'account-one',
      createdAt: index,
    }))
    ;(store.listTombstonesForAccount as ReturnType<typeof vi.fn>).mockResolvedValue(records)
    ;(store.listForAccount as ReturnType<typeof vi.fn>).mockResolvedValue([])
    let inFlight = 0
    let maximumInFlight = 0
    const failed = new Set(
      records.filter((_, index) => index % 17 === 0).map((entry) => entry.bindingId),
    )
    ;(pool.revokeBinding as ReturnType<typeof vi.fn>).mockImplementation(async (bindingId) => {
      inFlight += 1
      maximumInFlight = Math.max(maximumInFlight, inFlight)
      await new Promise((resolve) => setTimeout(resolve, 0))
      inFlight -= 1
      if (failed.has(bindingId)) throw new Error('offline')
    })

    await lifecycle.syncAccount()
    expect(pool.revokeBinding).toHaveBeenCalledTimes(256)
    expect(maximumInFlight).toBeLessThanOrEqual(12)
    expect(store.acknowledgeTombstone).toHaveBeenCalledTimes(256 - failed.size)
    for (const bindingId of failed)
      expect(store.acknowledgeTombstone).not.toHaveBeenCalledWith('account-one', bindingId)
  })

  it('deletes active records before remote logout revocation and retains failed tombstones', async () => {
    const { lifecycle, store, pool, order, tombstones } = harness()
    await lifecycle.syncAccount()
    order.length = 0
    ;(pool.revokeBinding as ReturnType<typeof vi.fn>)
      .mockRejectedValueOnce(new Error('offline'))
      .mockResolvedValue(undefined)

    const loggingOut = lifecycle.logout()
    expect(order[0]).toBe('suspend:logout')
    await loggingOut
    expect(order[0]).toBe('suspend:logout')
    expect(store.tombstoneAccount).toHaveBeenCalledWith('account-one')
    expect(tombstones.get('account-one')!.map((entry) => entry.bindingId)).toEqual([
      'binding_one_12345678',
    ])
  })

  it('never delivers old-account tombstones with replacement credentials and retries on old-account return', async () => {
    const { lifecycle, setAccount, order, tombstones } = harness()
    await lifecycle.syncAccount()
    order.length = 0
    setAccount({ loggedIn: true, userId: 'account-two' })

    await lifecycle.syncAccount()
    expect(order[0]).toBe('suspend:account_switch')
    expect(order.some((entry) => entry.startsWith('revoke:binding_one'))).toBe(false)
    expect(tombstones.get('account-one')?.map((entry) => entry.bindingId)).toEqual([
      'binding_one_12345678',
      'binding_two_12345678',
    ])
    expect(order.at(-1)).toBe('resume:binding_three_12345678')

    order.length = 0
    setAccount({ loggedIn: true, userId: 'account-one' })
    await lifecycle.syncAccount()
    expect(order).toContain('revoke:binding_one_12345678:account-one')
    expect(order).toContain('revoke:binding_two_12345678:account-one')
    expect(tombstones.get('account-one')).toEqual([])
  })

  it('tombstones on terminal auth loss but preserves everything on normal shutdown', async () => {
    const { lifecycle, store, pool, order, setAccount } = harness()
    await lifecycle.syncAccount()
    order.length = 0
    lifecycle.shutdown()
    expect(order).toEqual(['suspend:shutdown'])
    expect(store.tombstoneAccount).not.toHaveBeenCalled()

    setAccount({ loggedIn: false })
    await lifecycle.terminalAuthLoss()
    expect(store.tombstoneAccount).toHaveBeenCalledWith('account-one')
    expect(pool.suspend).toHaveBeenLastCalledWith('auth_required')
  })

  it('suspends sockets and retry loops in finally even when local tombstoning fails', async () => {
    const { lifecycle, store, pool, order } = harness()
    await lifecycle.syncAccount()
    order.length = 0
    ;(store.tombstoneAccount as ReturnType<typeof vi.fn>).mockRejectedValueOnce(
      new Error('disk_failure'),
    )

    await expect(lifecycle.logout()).rejects.toThrow('disk_failure')
    expect(order).toEqual(['suspend:logout', 'suspend:logout', 'suspend:logout'])
    expect(pool.revokeBinding).not.toHaveBeenCalledWith('binding_one_12345678')
  })

  it('blocks account switching before consuming the replacement credentials', () => {
    const source = readFileSync(join(import.meta.dirname, '../src/main/index.ts'), 'utf8')
    const callback = source.slice(source.indexOf('void authDeepLinks.initialize'))
    expect(callback.indexOf("officeRelay?.suspend('account_switch')")).toBeLessThan(
      callback.indexOf('consumeCallback(callback)'),
    )
  })

  it('wires the PC kill switch before creating or syncing persistent binding state', () => {
    const source = readFileSync(join(import.meta.dirname, '../src/main/index.ts'), 'utf8')
    const bootstrap = source.slice(source.indexOf('const officeMessagesProxy'))
    expect(bootstrap).toContain('officePairingResumeEnabled(process.env)')
    expect(bootstrap.indexOf('officePairingResumeEnabled(process.env)')).toBeLessThan(
      bootstrap.indexOf('createElectronOfficeRelayBindingStore'),
    )
    expect(bootstrap).toMatch(
      /const officeRelayBindingStore = persistentPairing\s+\? createElectronOfficeRelayBindingStore/,
    )
    expect(bootstrap).toContain('persistentPairing: () => officeRelayPersistenceAvailable')
    expect(bootstrap).toContain('if (persistentPairing && officeRelayBindingStore)')
    expect(bootstrap).toContain('fallbackToOrdinaryOfficeRelay({')
    expect(bootstrap).toContain('terminalOfficeRelayAuthLoss(officeRelayLifecycle, officeRelay)')
    expect(bootstrap).toContain("officeRelayDiagnostic = 'error:invalid_config'")
  })

  it('settles the selected Enhanced runtime before restoring an Office pairing', () => {
    const source = readFileSync(join(import.meta.dirname, '../src/main/index.ts'), 'utf8')
    const runtimeReady = source.indexOf('await codexRuntime.initialize().catch(() => undefined)')
    const relayCreated = source.indexOf('officeRelay = createOfficeRelayPool({')
    const pairingRestored = source.indexOf('await startOfficeRelayPersistence({')

    expect(runtimeReady).toBeGreaterThan(0)
    expect(runtimeReady).toBeLessThan(relayCreated)
    expect(runtimeReady).toBeLessThan(pairingRestored)
  })

  it('wires one-time relay closure and OAuth reactivation through executable runtime controls', () => {
    const source = readFileSync(join(import.meta.dirname, '../src/main/index.ts'), 'utf8')
    expect(source).toContain(
      'invalidate: () => logoutOfficeRelaySession(officeRelayLifecycle, officeRelay)',
    )
    expect(source).toContain('shutdownOfficeRelaySession(officeRelayLifecycle, officeRelay)')
    expect(source).toContain('syncOfficeRelaySession(')
    expect(source).toContain('officeRelayPersistenceAvailable,')
    expect(source).toMatch(
      /onBindingFailure:[\s\S]*lockOfficeRelayPersistence\([\s\S]*reason: 'logout'/,
    )
    expect(source).toMatch(
      /terminalOfficeRelayAuthLoss[\s\S]*lockOfficeRelayPersistence\([\s\S]*reason: 'auth_required'/,
    )
  })
})
