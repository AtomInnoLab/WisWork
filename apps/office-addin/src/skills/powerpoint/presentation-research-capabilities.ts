export interface PresentationResearchCapabilities {
  available: boolean
  cleanupAvailable: boolean
  historyVersion?: 2
  recoveryAvailable?: true
}
/** Negotiate the additive protocol; default callers keep cleanup negotiation; recovery opt-in adds one read-only fallback stage. */
export async function readPresentationResearchCapabilities(
  request: (operation: string, body?: Record<string, unknown>) => Promise<unknown>,
  options?: { includeRecovery?: true },
): Promise<PresentationResearchCapabilities> {
  const attempts = [
    ...(options?.includeRecovery ? [{ includeCleanup: true, includeRecovery: true }] : []),
    { includeCleanup: true },
    {},
  ]
  let value: unknown
  for (const [index, body] of attempts.entries()) {
    try {
      value = await request('research_capabilities', body)
      break
    } catch (error) {
      if (
        !(error instanceof Error) ||
        error.message !== 'presentation_upgrade_required' ||
        index === attempts.length - 1
      )
        throw error
    }
  }
  const invalid = (): never => {
    throw new Error('presentation_response_invalid')
  }
  if (!value || typeof value !== 'object' || Array.isArray(value)) invalid()
  const caps = value as Record<string, unknown>
  if (caps.version !== 1 || typeof caps.available !== 'boolean') invalid()
  const keys = Object.keys(caps).sort().join(',')
  if (keys === 'available,version')
    return { available: caps.available as boolean, cleanupAvailable: false }
  const recovery = keys === 'available,cleanupAvailable,historyVersions,recoveryAvailable,version'
  if (
    (recovery
      ? !options?.includeRecovery || caps.recoveryAvailable !== true
      : keys !== 'available,cleanupAvailable,historyVersions,version') ||
    caps.available !== true ||
    caps.cleanupAvailable !== true ||
    JSON.stringify(caps.historyVersions) !== '[1,2]'
  )
    invalid()
  return {
    available: true,
    cleanupAvailable: true,
    historyVersion: 2,
    ...(recovery ? { recoveryAvailable: true as const } : {}),
  }
}
