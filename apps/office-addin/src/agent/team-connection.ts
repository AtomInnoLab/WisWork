import type { AuthClient } from '@wiswork/auth/browser'
import {
  createOfficeRelaySession,
  type OfficeRelaySession,
  type OfficeRelayStatus,
} from '../relay/session.js'
export interface OfficeTeamConnectionSnapshot {
  phase: 'signed_out' | 'signing_in' | 'pairing' | 'ready' | 'error'
  account?: Readonly<{ loggedIn: boolean; email?: string; userId?: string }>
  relayStatus?: OfficeRelayStatus
  verificationCode?: string
  error?: string
}
export interface OfficeTeamConnection {
  snapshot(): OfficeTeamConnectionSnapshot
  subscribe(listener: () => void): () => void
  signIn(): Promise<void>
  connect(): Promise<void>
  signOut(): Promise<void>
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  dispose(): void
  available(): boolean
}
interface Options {
  auth: AuthClient
  loginDialog(authorizationURL: string, signal?: AbortSignal): Promise<string>
  createSession?(options: {
    capabilities: readonly ['presentation-team.v1']
    getTeamAccessToken: () => Promise<string | null>
  }): OfficeRelaySession
  onUnavailable?(): void
}
const safeErrors = new Set([
  'team_auth_dialog_unavailable',
  'team_auth_dialog_invalid',
  'team_auth_dialog_failed',
  'team_auth_dialog_cancelled',
  'team_auth_dialog_timeout',
  'team_auth_dialog_busy',
  'invalid_callback',
  'invalid_state',
  'callback_reused',
  'callback_expired',
  'auth_required',
  'network_error',
  'relay_cancelled',
  'relay_disconnected',
  'relay_team_auth_unavailable',
  'relay_capability_unavailable',
  'relay_busy',
  'team_connection_account_changed',
  'team_connection_invalid_capability',
  'team_connection_cancelled',
])
function safeError(error: unknown) {
  const code = error instanceof Error ? error.message : ''
  return safeErrors.has(code) ? code : 'team_connection_unavailable'
}
function fail(code: string): never {
  throw Error(code)
}
/** Independent authenticated team channel. Public metadata is not a trusted team author identity. */
export function createOfficeTeamConnection(options: Options): OfficeTeamConnection {
  let epoch = 0,
    disposed = false,
    active: AbortController | undefined
  let state: OfficeTeamConnectionSnapshot = Object.freeze({ phase: 'signed_out' })
  const listeners = new Set<() => void>()
  const publish = (next: OfficeTeamConnectionSnapshot) => {
    state = Object.freeze({
      ...next,
      ...(next.account ? { account: Object.freeze({ ...next.account }) } : {}),
    })
    listeners.forEach((l) => l())
  }
  const invalidate = () => {
    try {
      options.onUnavailable?.()
    } catch {
      /* A cleanup callback cannot retain channel availability. */
    }
  }
  const guard = (captured: number, signal?: AbortSignal) => {
    if (disposed || captured !== epoch || signal?.aborted) fail('team_connection_cancelled')
  }
  const account = async (captured: number, signal?: AbortSignal) => {
    const value = await options.auth.getValidAccountStatus()
    guard(captured, signal)
    if (!value.loggedIn) fail('auth_required')
    if (
      (value.userId !== undefined &&
        (typeof value.userId !== 'string' || !value.userId || value.userId.length > 512)) ||
      (value.email !== undefined && (typeof value.email !== 'string' || value.email.length > 512))
    )
      fail('auth_required')
    return {
      loggedIn: true,
      ...(value.email ? { email: value.email } : {}),
      ...(value.userId ? { userId: value.userId } : {}),
    }
  }
  const stop = (phase: 'signed_out' | 'error', error?: string, clearAccount = false) => {
    epoch++
    active?.abort()
    active = undefined
    publish({
      phase,
      ...(!clearAccount && state.account ? { account: state.account } : {}),
      relayStatus: 'offline',
      ...(error ? { error } : {}),
    })
    invalidate()
    session.disconnect()
  }
  const credentialsLost = async (error = 'auth_required') => {
    stop('signed_out', error, true)
    try {
      await options.auth.logout()
    } catch {
      /* Credentials remain unavailable through this controller. */
    }
  }
  const getTeamAccessToken = async () => {
    const captured = epoch
    if (disposed || !state.account?.loggedIn) return null
    try {
      const current = await account(captured)
      if (current.userId !== state.account?.userId) fail('team_connection_account_changed')
      const token = await options.auth.getAccessToken()
      guard(captured)
      const after = await account(captured)
      if (after.userId !== current.userId) fail('team_connection_account_changed')
      if (typeof token !== 'string' || !token || token.length > 4096) fail('auth_required')
      return token
    } catch (error) {
      if (!disposed && captured === epoch) await credentialsLost(safeError(error))
      return null
    }
  }
  const session = (options.createSession ?? createOfficeRelaySession)({
    capabilities: ['presentation-team.v1'],
    getTeamAccessToken,
  })
  const updateRelay = () => {
    if (disposed || !state.account?.loggedIn) return
    const relay = session.snapshot()
    if (relay.status === 'connected') {
      if (relay.capabilities?.length !== 1 || relay.capabilities[0] !== 'presentation-team.v1') {
        stop('error', 'team_connection_invalid_capability')
        return
      }
      publish({ phase: 'ready', account: state.account, relayStatus: relay.status })
    } else if (['connecting', 'pending', 'waiting_for_pc'].includes(relay.status)) {
      publish({
        phase: 'pairing',
        account: state.account,
        relayStatus: relay.status,
        ...(relay.verificationCode ? { verificationCode: relay.verificationCode } : {}),
      })
    } else if (state.phase === 'ready' || state.phase === 'pairing') {
      // Publish before disconnecting to avoid an offline event recursively invalidating this state.
      stop('error', 'relay_disconnected')
    }
  }
  const unsubscribe = session.subscribe(updateRelay)
  const pair = async (captured: number, signal: AbortSignal) => {
    const publicAccount = await account(captured, signal)
    if (state.account && publicAccount.userId !== state.account.userId)
      fail('team_connection_account_changed')
    publish({ phase: 'pairing', account: publicAccount, relayStatus: 'connecting' })
    await session.connect('powerpoint')
    guard(captured, signal)
    updateRelay()
  }
  const begin = () => {
    if (disposed) fail('team_connection_cancelled')
    if (active) fail('relay_busy')
    const controller = new AbortController()
    active = controller
    return { captured: ++epoch, controller }
  }
  const finishError = async (error: unknown, captured: number) => {
    if (disposed || captured !== epoch) return
    const code = safeError(error)
    if (code === 'auth_required' || code === 'team_connection_account_changed')
      await credentialsLost(code)
    else stop('error', code)
  }
  return {
    snapshot: () => state,
    subscribe(listener) {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    available: () =>
      !disposed &&
      state.phase === 'ready' &&
      session.snapshot().status === 'connected' &&
      session.snapshot().capabilities?.length === 1 &&
      session.snapshot().capabilities?.[0] === 'presentation-team.v1',
    async signIn() {
      const { captured, controller } = begin()
      publish({ phase: 'signing_in' })
      invalidate()
      session.disconnect()
      try {
        await options.auth.logout()
        guard(captured, controller.signal)
        const request = options.auth.createAuthorizationRequest()
        guard(captured, controller.signal)
        const callback = await options.loginDialog(request.url, controller.signal)
        guard(captured, controller.signal)
        await options.auth.consumeCallback(callback)
        guard(captured, controller.signal)
        await pair(captured, controller.signal)
      } catch (error) {
        await finishError(error, captured)
      } finally {
        if (active === controller) active = undefined
      }
    },
    async connect() {
      const { captured, controller } = begin()
      try {
        await pair(captured, controller.signal)
      } catch (error) {
        await finishError(error, captured)
      } finally {
        if (active === controller) active = undefined
      }
    },
    async signOut() {
      stop('signed_out', undefined, true)
      const captured = epoch
      try {
        await options.auth.logout()
      } catch {
        if (!disposed && captured === epoch)
          publish({ phase: 'signed_out', error: 'team_connection_unavailable' })
      }
    },
    async request(body, signal) {
      const captured = epoch
      guard(captured, signal)
      if (state.phase !== 'ready') fail('team_connection_unavailable')
      try {
        const token = await getTeamAccessToken()
        guard(captured, signal)
        if (!token) fail('auth_required')
        const response = await session.capabilityFetch('presentation-team.v1', body, signal)
        guard(captured, signal)
        if (state.phase !== 'ready') fail('team_connection_unavailable')
        return response
      } catch (error) {
        if (captured === epoch && !signal?.aborted) await finishError(error, captured)
        fail(safeError(error))
      }
    },
    dispose() {
      if (disposed) return
      disposed = true
      epoch++
      active?.abort()
      active = undefined
      unsubscribe()
      publish({ phase: 'signed_out' })
      invalidate()
      session.disconnect()
      void options.auth.logout().catch(() => {})
      listeners.clear()
    },
  }
}
