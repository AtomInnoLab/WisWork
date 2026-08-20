import { WISWORK_MESSAGES_URL } from '@wiswork/ai-provider'
import type { MessagesProxy, OfficeBridge } from '@wiswork/office-bridge'

export const DEFAULT_OFFICE_BRIDGE_PORT = 43_127
export const DEFAULT_OFFICE_ORIGIN = 'https://office.8-216-134-194.sslip.io'

export function officeBridgeEnabled(env: Record<string, string | undefined>): boolean {
  return env.WISWORK_OFFICE_BRIDGE_ENABLED === '1'
}

export function officeBridgePortFromEnv(env: Record<string, string | undefined>): number {
  const raw = env.WISWORK_OFFICE_BRIDGE_PORT
  if (raw === undefined) return DEFAULT_OFFICE_BRIDGE_PORT
  const port = Number(raw)
  if (!Number.isSafeInteger(port) || port < 1 || port > 65_535)
    throw new Error('invalid_office_bridge_port')
  return port
}

export function officeOriginFromEnv(env: Record<string, string | undefined>): string {
  const origin = env.WISWORK_OFFICE_ORIGIN ?? DEFAULT_OFFICE_ORIGIN
  const parsed = new URL(origin)
  if (parsed.protocol !== 'https:' || parsed.origin !== origin)
    throw new Error('invalid_office_origin')
  return origin
}

export async function logoutAndRevokeOfficeBridge(
  bridge: Pick<OfficeBridge, 'revokeAll'> | null,
  logout: () => Promise<void>,
): Promise<void> {
  bridge?.revokeAll()
  await logout()
}

export async function validAccountStatusOrRevoke(
  bridge: Pick<OfficeBridge, 'revokeAll'>,
  getStatus: () => Promise<{ loggedIn: boolean }>,
): Promise<{ loggedIn: boolean }> {
  try {
    const status = await getStatus()
    if (!status.loggedIn) bridge.revokeAll()
    return status
  } catch (error) {
    if (
      error instanceof Error &&
      (error.message === 'auth_required' ||
        (error as Error & { code?: string }).code === 'auth_required')
    )
      bridge.revokeAll()
    throw error
  }
}

export function createOfficeMessagesProxy(options: {
  fetchWithAuth(request: (accessToken: string) => Promise<Response>): Promise<Response>
  fetch?: typeof fetch
  onTerminalAuthLoss?: () => void
}): MessagesProxy {
  const doFetch = options.fetch ?? fetch
  const streamBody = (body: ReadableStream<Uint8Array> | null): AsyncIterable<Uint8Array> => ({
    async *[Symbol.asyncIterator]() {
      if (!body) return
      const reader = body.getReader()
      try {
        while (true) {
          const result = await reader.read()
          if (result.done) return
          yield result.value
        }
      } finally {
        reader.releaseLock()
      }
    },
  })
  return async ({ body, signal }) => {
    try {
      const upstream = await options.fetchWithAuth((accessToken) =>
        doFetch(WISWORK_MESSAGES_URL, {
          method: 'POST',
          headers: {
            authorization: `Bearer ${accessToken}`,
            'content-type': 'application/json',
            'x-req-location': 'sg',
          },
          body: JSON.stringify(body),
          signal,
        }),
      )
      if (upstream.status === 401) throw new Error('auth_required')
      return {
        status: upstream.status,
        contentType: upstream.headers.get('content-type') ?? undefined,
        body: streamBody(upstream.body),
      }
    } catch (error) {
      if (error instanceof Error && error.message === 'auth_required')
        options.onTerminalAuthLoss?.()
      throw error
    }
  }
}

export async function runOfficeBridgeLifecycle(options: {
  app: { whenReady(): Promise<unknown>; on(name: 'before-quit', listener: () => void): unknown }
  bridge: Pick<OfficeBridge, 'revokeAll' | 'shutdown'>
  start(): Promise<{ stop(): Promise<void> }>
}): Promise<void> {
  await options.app.whenReady()
  const server = await options.start()
  let stopped = false
  options.app.on('before-quit', () => {
    if (stopped) return
    stopped = true
    options.bridge.revokeAll()
    options.bridge.shutdown()
    void server.stop()
  })
}
