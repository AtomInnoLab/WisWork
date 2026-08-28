export function officePairingResumeEnabled(env: Record<string, string | undefined>): boolean {
  const value = env.WISWORK_OFFICE_PAIRING_RESUME
  if (value === undefined || value === '1') return true
  if (value === '0') return false
  throw new Error('invalid_office_pairing_resume')
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
  suspend(reason: string): void
  onBindingFailure(error: unknown): void
}): Promise<boolean> {
  try {
    await options.sync()
    return true
  } catch (error) {
    try {
      options.suspend('binding_lifecycle')
    } catch {
      // Startup remains available even if a partial pool failed to suspend cleanly.
    }
    try {
      options.onBindingFailure(error)
    } catch {
      // Startup remains available with persistent Relay disabled.
    }
    return false
  }
}

export async function syncOfficeRelayAccountSafely<T>(options: {
  getAccountStatus(): Promise<T>
  syncBindingLifecycle(): Promise<void>
  suspend(reason: string): void
  onBindingFailure(error: unknown): void
}): Promise<T> {
  // Authentication/account failures are outside the binding boundary and must propagate.
  const account = await options.getAccountStatus()
  try {
    await options.syncBindingLifecycle()
  } catch (error) {
    try {
      options.suspend('binding_lifecycle')
    } catch {
      // The account result remains authoritative even if the binding pool is partially built.
    }
    try {
      options.onBindingFailure(error)
    } catch {
      // Renderer diagnostics must not change authentication semantics.
    }
  }
  return account
}

export async function completeOfficeRelayOAuthLogin<T>(options: {
  consumeCallback(): Promise<unknown>
  getAccountStatus(): Promise<T>
  syncBindingLifecycle(): Promise<void>
  suspend(reason: string): void
  onBindingFailure(error: unknown): void
}): Promise<T> {
  // OAuth validation failures deliberately remain visible to the deep-link queue.
  await options.consumeCallback()
  return syncOfficeRelayAccountSafely(options)
}
