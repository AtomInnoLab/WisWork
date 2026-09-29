export interface PresentationResearchCapabilities {
  available: boolean
  cleanupAvailable: boolean
  historyVersion?: 2
}
/** Negotiate the additive protocol; a legacy rejection gets one read-only fallback. */
export async function readPresentationResearchCapabilities(
  request: (operation: string, body?: Record<string, unknown>) => Promise<unknown>,
): Promise<PresentationResearchCapabilities> {
  let value: unknown
  try {
    value = await request('research_capabilities', { includeCleanup: true })
  } catch (error) {
    if (!(error instanceof Error) || error.message !== 'presentation_upgrade_required') throw error
    value = await request('research_capabilities', {})
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
  if (
    keys !== 'available,cleanupAvailable,historyVersions,version' ||
    caps.available !== true ||
    caps.cleanupAvailable !== true ||
    JSON.stringify(caps.historyVersions) !== '[1,2]'
  )
    invalid()
  return { available: true, cleanupAvailable: true, historyVersion: 2 }
}
