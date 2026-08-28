import { describe, expect, it, vi } from 'vitest'

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
    revokeBinding: vi.fn(async (bindingId: string) => {
      order.push(`revoke:${bindingId}`)
    }),
    approve: vi.fn(async () => false),
    reject: vi.fn(() => false),
    listPending: vi.fn(() => []),
    status: vi.fn(() => 'disconnected' as const),
    revoke: vi.fn((reason?: string) => order.push(`disconnect:${reason}`)),
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
    expect(pool.revoke).toHaveBeenCalledWith('auth_required')
  })

  it('retries same-account tombstones before resuming every active binding at startup', async () => {
    const { lifecycle, order, tombstones } = harness()
    await lifecycle.syncAccount()
    expect(order).toEqual([
      'revoke:binding_stale_12345678',
      'ack:binding_stale_12345678',
      'resume:binding_one_12345678',
      'resume:binding_two_12345678',
    ])
    expect(tombstones.get('account-one')).toEqual([])
  })

  it('deletes active records before remote logout revocation and retains failed tombstones', async () => {
    const { lifecycle, store, pool, order, tombstones } = harness()
    await lifecycle.syncAccount()
    order.length = 0
    ;(pool.revokeBinding as ReturnType<typeof vi.fn>)
      .mockRejectedValueOnce(new Error('offline'))
      .mockResolvedValue(undefined)

    await lifecycle.logout()
    expect(order[0]).toBe('tombstone:account-one')
    expect(order.at(-1)).toBe('disconnect:logout')
    expect(store.tombstoneAccount).toHaveBeenCalledWith('account-one')
    expect(tombstones.get('account-one')!.map((entry) => entry.bindingId)).toEqual([
      'binding_one_12345678',
    ])
  })

  it('queues old-account revocation on account switch and starts only the new account bindings', async () => {
    const { lifecycle, setAccount, order } = harness()
    await lifecycle.syncAccount()
    order.length = 0
    setAccount({ loggedIn: true, userId: 'account-two' })

    await lifecycle.syncAccount()
    expect(order[0]).toBe('tombstone:account-one')
    expect(order).toContain('disconnect:account_switch')
    expect(order.at(-1)).toBe('resume:binding_three_12345678')
  })

  it('tombstones on terminal auth loss but preserves everything on normal shutdown', async () => {
    const { lifecycle, store, pool, order, setAccount } = harness()
    await lifecycle.syncAccount()
    order.length = 0
    lifecycle.shutdown()
    expect(order).toEqual(['disconnect:shutdown'])
    expect(store.tombstoneAccount).not.toHaveBeenCalled()

    setAccount({ loggedIn: false })
    await lifecycle.terminalAuthLoss()
    expect(store.tombstoneAccount).toHaveBeenCalledWith('account-one')
    expect(pool.revoke).toHaveBeenLastCalledWith('auth_required')
  })
})
