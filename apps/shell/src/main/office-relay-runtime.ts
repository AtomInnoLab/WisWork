import type { OfficeRelayBinding } from './office-relay-binding-store'

export function officePairingResumeEnabled(env: Record<string, string | undefined>): boolean {
  const value = env.WISWORK_OFFICE_PAIRING_RESUME
  if (value === undefined || value === '1') return true
  if (value === '0') return false
  throw new Error('invalid_office_pairing_resume')
}

interface OfficeRelaySessionLifecycle {
  logout(): Promise<void>
  terminalAuthLoss(): Promise<void>
  shutdown(): void
}

interface OfficeRelaySessionPool {
  suspend(reason: string): void
}

interface OfficeRelaySyncLifecycle extends OfficeRelaySessionLifecycle {
  syncAccount(canActivate?: () => boolean): Promise<void>
}

interface OfficeRelaySyncPool extends OfficeRelaySessionPool {
  activate(): void
}

export interface OfficeRelayActivationFence {
  snapshot(): number
  isCurrent(expectedEpoch: number): boolean
  lock(): void
  settle(options: {
    expectedEpoch: number
    allowUnlock: boolean
    getValidAccountStatus(): Promise<{ loggedIn: boolean }>
    pool: OfficeRelaySyncPool | null
    suspendedReason: string
  }): Promise<boolean>
}

export function createOfficeRelayActivationFence(): OfficeRelayActivationFence {
  let epoch = 0
  let locked = false
  return {
    snapshot: () => epoch,
    isCurrent: (expectedEpoch) => expectedEpoch === epoch,
    lock() {
      epoch += 1
      locked = true
    },
    async settle(options) {
      if (options.expectedEpoch !== epoch) return false
      if (!options.allowUnlock) {
        options.pool?.suspend(options.suspendedReason)
        return false
      }
      let account: { loggedIn: boolean }
      try {
        account = await options.getValidAccountStatus()
      } catch (error) {
        if (options.expectedEpoch !== epoch) return false
        options.pool?.suspend(options.suspendedReason)
        throw error
      }
      if (options.expectedEpoch !== epoch) return false
      if (!account.loggedIn) {
        options.pool?.suspend(options.suspendedReason)
        return false
      }
      locked = false
      options.pool?.activate()
      return !locked
    },
  }
}

export function lockOfficeRelayPersistence(options: {
  disable(): void
  pool: OfficeRelaySessionPool | null
  reason: 'logout' | 'auth_required' | 'binding_lifecycle'
}): void {
  options.disable()
  options.pool?.suspend(options.reason)
}

export async function saveOfficeRelayBindingForCurrentAccount(options: {
  binding: OfficeRelayBinding
  persistenceAvailable(): boolean
  getValidAccountStatus(): Promise<{ loggedIn: boolean; userId?: string }>
  put(binding: OfficeRelayBinding): Promise<void>
  tombstoneAccount(accountId: string): Promise<void>
  onStorageFailure(error: unknown): Promise<void> | void
}): Promise<void> {
  if (!options.persistenceAvailable()) throw new Error('binding_lifecycle_disabled')
  const tombstoneStaleBinding = async () => {
    try {
      await options.put(options.binding)
      await options.tombstoneAccount(options.binding.accountId)
    } catch (error) {
      await options.onStorageFailure(error)
      throw error
    }
    throw new Error('stale_binding_account')
  }
  const before = await options.getValidAccountStatus()
  if (!before.loggedIn || before.userId !== options.binding.accountId)
    return tombstoneStaleBinding()
  try {
    await options.put(options.binding)
  } catch (error) {
    await options.onStorageFailure(error)
    throw error
  }
  let after: { loggedIn: boolean; userId?: string }
  try {
    after = await options.getValidAccountStatus()
  } catch (accountError) {
    try {
      await options.tombstoneAccount(options.binding.accountId)
    } catch (storageError) {
      await options.onStorageFailure(storageError)
      throw storageError
    }
    throw accountError
  }
  if (
    !options.persistenceAvailable() ||
    !after.loggedIn ||
    after.userId !== options.binding.accountId
  ) {
    try {
      await options.tombstoneAccount(options.binding.accountId)
    } catch (error) {
      await options.onStorageFailure(error)
      throw error
    }
    throw new Error('stale_binding_account')
  }
}

export async function invalidateOfficeRelayBindingForCurrentAccount(options: {
  binding: OfficeRelayBinding
  persistenceAvailable(): boolean
  getValidAccountStatus(): Promise<{ loggedIn: boolean; userId?: string }>
  remove(accountId: string, bindingId: string): Promise<void>
  onStorageFailure(error: unknown): Promise<void> | void
}): Promise<void> {
  if (!options.persistenceAvailable()) return
  const account = await options.getValidAccountStatus()
  if (!account.loggedIn || account.userId !== options.binding.accountId) return
  try {
    await options.remove(options.binding.accountId, options.binding.bindingId)
  } catch (error) {
    await options.onStorageFailure(error)
  }
}

export async function syncOfficeRelaySession(
  account: { loggedIn: boolean; userId?: string },
  persistenceAvailable: boolean,
  lifecycle: OfficeRelaySyncLifecycle | null,
  pool: OfficeRelaySyncPool | null,
  canActivate?: () => boolean,
): Promise<void> {
  if (persistenceAvailable && lifecycle) {
    await lifecycle.syncAccount(canActivate)
    return
  }
  if (canActivate && !canActivate()) return
  if (account.loggedIn) pool?.activate()
  else pool?.suspend('auth_required')
}

export function logoutOfficeRelaySession(
  lifecycle: OfficeRelaySessionLifecycle | null,
  pool: OfficeRelaySessionPool | null,
): Promise<void> {
  if (lifecycle) return lifecycle.logout()
  pool?.suspend('logout')
  return Promise.resolve()
}

export function terminalOfficeRelayAuthLoss(
  lifecycle: OfficeRelaySessionLifecycle | null,
  pool: OfficeRelaySessionPool | null,
): Promise<void> {
  if (lifecycle) return lifecycle.terminalAuthLoss()
  pool?.suspend('auth_required')
  return Promise.resolve()
}

export function shutdownOfficeRelaySession(
  lifecycle: OfficeRelaySessionLifecycle | null,
  pool: OfficeRelaySessionPool | null,
): void {
  if (lifecycle) lifecycle.shutdown()
  else pool?.suspend('shutdown')
}

export async function logoutWithOfficeRelay<T>(options: {
  invalidate(): Promise<void>
  logout(): Promise<T>
  onBindingFailure(error: unknown): void
}): Promise<T> {
  try {
    await options.invalidate()
  } catch (error) {
    try {
      options.onBindingFailure(error)
    } catch {
      // Authentication logout must not depend on diagnostic reporting.
    }
  }
  return options.logout()
}

export async function startOfficeRelayPersistence(options: {
  sync(): Promise<void>
  fallbackToOrdinary(): Promise<void> | void
  onBindingFailure(error: unknown): void
}): Promise<boolean> {
  try {
    await options.sync()
    return true
  } catch (error) {
    try {
      await options.fallbackToOrdinary()
    } catch {
      // Startup remains available even if the ordinary-pairing fallback fails closed.
    }
    try {
      options.onBindingFailure(error)
    } catch {
      // Startup remains available with persistent Relay disabled.
    }
    return false
  }
}

export async function syncOfficeRelayAccountSafely<T extends { loggedIn: boolean }>(options: {
  getAccountStatus(): Promise<T>
  syncBindingLifecycle(account: T): Promise<void>
  fallbackToOrdinary(account: T): Promise<void> | void
  onBindingFailure(error: unknown): void
}): Promise<T> {
  // Authentication/account failures are outside the binding boundary and must propagate.
  const account = await options.getAccountStatus()
  try {
    await options.syncBindingLifecycle(account)
  } catch (error) {
    try {
      await options.fallbackToOrdinary(account)
    } catch {
      // The account result remains authoritative even if fallback fails closed.
    }
    try {
      options.onBindingFailure(error)
    } catch {
      // Renderer diagnostics must not change authentication semantics.
    }
  }
  return account
}

export async function completeOfficeRelayOAuthLogin<T extends { loggedIn: boolean }>(options: {
  consumeCallback(): Promise<unknown>
  getAccountStatus(): Promise<T>
  syncBindingLifecycle(account: T): Promise<void>
  fallbackToOrdinary(account: T): Promise<void> | void
  onBindingFailure(error: unknown): void
}): Promise<T> {
  // OAuth validation failures deliberately remain visible to the deep-link queue.
  await options.consumeCallback()
  return syncOfficeRelayAccountSafely(options)
}
