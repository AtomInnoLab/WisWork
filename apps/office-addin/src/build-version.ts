const BUILD_ID = /^[A-Za-z0-9_.-]{3,96}$/

export type BuildVersionState =
  | { status: 'checking' }
  | { status: 'current' }
  | { status: 'stale'; buildId: string }

export function resolveBuildVersion(deployed: string | undefined, current: string): BuildVersionState {
  return deployed && deployed !== current
    ? { status: 'stale', buildId: deployed }
    : { status: 'current' }
}

export async function deployedBuildId(fetcher: typeof fetch = fetch): Promise<string | undefined> {
  const controller = new AbortController()
  let timer: ReturnType<typeof setTimeout> | undefined
  try {
    return await Promise.race([
      (async () => {
        const response = await fetcher(new URL('/version.json', location.origin), {
          cache: 'no-store',
          credentials: 'same-origin',
          signal: controller.signal,
        })
        if (!response.ok) return undefined
        const payload: unknown = await response.json()
        if (!payload || typeof payload !== 'object' || Array.isArray(payload)) return undefined
        const buildId = (payload as Record<string, unknown>).buildId
        return typeof buildId === 'string' && BUILD_ID.test(buildId) ? buildId : undefined
      })(),
      new Promise<undefined>((resolve) => {
        timer = setTimeout(() => {
          controller.abort()
          resolve(undefined)
        }, 5_000)
      }),
    ])
  } catch {
    return undefined
  } finally {
    clearTimeout(timer)
  }
}
