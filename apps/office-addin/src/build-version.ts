const BUILD_ID = /^[A-Za-z0-9_.-]{3,96}$/

export async function deployedBuildId(fetcher: typeof fetch = fetch): Promise<string | undefined> {
  try {
    const response = await fetcher(new URL('/version.json', location.origin), {
      cache: 'no-store',
      credentials: 'same-origin',
    })
    if (!response.ok) return undefined
    const payload: unknown = await response.json()
    if (!payload || typeof payload !== 'object' || Array.isArray(payload)) return undefined
    const buildId = (payload as Record<string, unknown>).buildId
    return typeof buildId === 'string' && BUILD_ID.test(buildId) ? buildId : undefined
  } catch {
    return undefined
  }
}
