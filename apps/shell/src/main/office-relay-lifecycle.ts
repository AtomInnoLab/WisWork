import type { OfficeRelayBindingStore } from './office-relay-binding-store'
import type { OfficeRelayPool } from './office-relay-pool'

export interface OfficeRelayLifecycle {
  syncAccount(): Promise<void>
  logout(): Promise<void>
  terminalAuthLoss(): Promise<void>
  shutdown(): void
}

export function createOfficeRelayLifecycle(options: {
  store: OfficeRelayBindingStore
  pool: OfficeRelayPool
  getAccountStatus?: () => Promise<{ loggedIn: boolean; userId?: string }>
  getValidAccountStatus(): Promise<{ loggedIn: boolean; userId?: string }>
}): OfficeRelayLifecycle {
  let activeAccountId: string | null = null
  let operation = Promise.resolve()

  const serialize = (action: () => Promise<void>): Promise<void> => {
    const next = operation.then(action, action)
    operation = next.catch(() => undefined)
    return next
  }

  const deliverTombstones = async (accountId: string) => {
    const tombstones = await options.store.listTombstonesForAccount(accountId)
    await Promise.all(
      tombstones.map(async (tombstone) => {
        try {
          await options.pool.revokeBinding(tombstone.bindingId)
          await options.store.acknowledgeTombstone(accountId, tombstone.bindingId)
        } catch {
          // Keep the encrypted tombstone for this same account's next authenticated startup.
        }
      }),
    )
  }

  const invalidate = async (
    accountId: string,
    reason: 'logout' | 'auth_required' | 'account_switch',
  ) => {
    await options.store.tombstoneAccount(accountId)
    await deliverTombstones(accountId)
    options.pool.revoke(reason)
    if (activeAccountId === accountId) activeAccountId = null
  }

  return {
    syncAccount() {
      return serialize(async () => {
        const localAccount = await (options.getAccountStatus ?? options.getValidAccountStatus)()
        if (!activeAccountId && localAccount.loggedIn && localAccount.userId)
          activeAccountId = localAccount.userId
        const account = await options.getValidAccountStatus()
        if (!account.loggedIn || !account.userId) {
          if (activeAccountId) await invalidate(activeAccountId, 'auth_required')
          return
        }
        if (activeAccountId && activeAccountId !== account.userId)
          await invalidate(activeAccountId, 'account_switch')
        activeAccountId = account.userId
        await deliverTombstones(account.userId)
        const bindings = await options.store.listForAccount(account.userId)
        for (const binding of bindings) await options.pool.resume(binding)
      })
    },
    logout() {
      return serialize(async () => {
        const localAccount = await (
          options.getAccountStatus ?? options.getValidAccountStatus
        )().catch(() => ({ loggedIn: false as const }))
        const account =
          activeAccountId ??
          (localAccount.loggedIn && localAccount.userId ? localAccount.userId : null)
        if (account) await invalidate(account, 'logout')
        else options.pool.revoke('logout')
      })
    },
    terminalAuthLoss() {
      return serialize(async () => {
        const account = activeAccountId
        if (account) await invalidate(account, 'auth_required')
        else options.pool.revoke('auth_required')
      })
    },
    shutdown() {
      options.pool.revoke('shutdown')
    },
  }
}
