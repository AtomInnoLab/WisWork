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
