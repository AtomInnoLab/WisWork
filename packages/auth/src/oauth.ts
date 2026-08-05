import { createHash, randomBytes as nodeRandomBytes, timingSafeEqual } from 'node:crypto'
import { DEFAULT_AUTH_CONFIG, type AuthConfig } from './config'
import {
  AuthError,
  parseSessionPayload,
  publicAccountStatus,
  type AccountStatus,
  type AuthSession,
  type SessionStore,
} from './session'

interface PendingTransaction {
  state: string
  verifier: string
  expiresAt: number
  consumed: boolean
  attemptGeneration: number
}

export interface AuthClientOptions {
  store: SessionStore
  fetch?: typeof globalThis.fetch
  config?: Partial<AuthConfig>
  now?: () => number
  randomBytes?: (size: number) => Uint8Array
}

export interface AuthClient {
  createAuthorizationRequest(): { url: string; state: string }
  consumeCallback(url: string): Promise<AuthSession>
  getAccountStatus(): Promise<AccountStatus>
  getAccessToken(): Promise<string | null>
  refresh(): Promise<AuthSession>
  logout(): Promise<void>
  fetchWithAuth(request: (accessToken: string) => Promise<Response>): Promise<Response>
  setNowForTesting(now: () => number): void
}

function base64Url(bytes: Uint8Array): string {
  return Buffer.from(bytes).toString('base64url')
}

function equalSecret(a: string, b: string): boolean {
  const left = Buffer.from(a)
  const right = Buffer.from(b)
  return left.length === right.length && timingSafeEqual(left, right)
}

async function parseResponse(response: Response, now: number): Promise<AuthSession> {
  if (!response.ok) throw new AuthError(response.status === 401 ? 'auth_required' : 'network_error')
  try {
    return parseSessionPayload(await response.json(), now)
  } catch (error) {
    if (error instanceof AuthError) throw error
    throw new AuthError('invalid_response')
  }
}

export function createAuthClient(options: AuthClientOptions): AuthClient {
  const config: AuthConfig = { ...DEFAULT_AUTH_CONFIG, ...options.config }
  const doFetch = options.fetch ?? globalThis.fetch
  const randomBytes = options.randomBytes ?? ((size: number) => nodeRandomBytes(size))
  let now = options.now ?? Date.now
  const pending = new Map<string, PendingTransaction>()
  let refreshFlight: Promise<AuthSession> | null = null
  let loginAttemptGeneration = 0
  let sessionRevision = 0
  let mutationTail: Promise<void> = Promise.resolve()

  const mutateSession = <T>(operation: () => Promise<T>): Promise<T> => {
    const result = mutationTail.then(operation)
    mutationTail = result.then(
      () => undefined,
      () => undefined,
    )
    return result
  }

  const client: AuthClient = {
    createAuthorizationRequest() {
      const state = base64Url(randomBytes(32))
      const verifier = base64Url(randomBytes(32))
      const challenge = createHash('sha256').update(verifier).digest('base64url')
      loginAttemptGeneration += 1
      pending.set(state, {
        state,
        verifier,
        expiresAt: now() + config.transactionTtlMs,
        consumed: false,
        attemptGeneration: loginAttemptGeneration,
      })
      const url = new URL(config.authorizationEndpoint)
      url.search = new URLSearchParams({
        client_id: config.clientId,
        redirect_uri: config.redirectUri,
        response_type: 'code',
        scope: config.scope,
        state,
        code_challenge: challenge,
        code_challenge_method: 'S256',
      }).toString()
      return { url: url.toString(), state }
    },
    async consumeCallback(rawUrl) {
      if (typeof rawUrl !== 'string' || rawUrl.length > 4_096)
        throw new AuthError('invalid_callback')
      let url: URL
      try {
        url = new URL(rawUrl)
      } catch {
        throw new AuthError('invalid_callback')
      }
      if (
        url.protocol !== 'wiswork:' ||
        url.hostname !== 'oauth' ||
        url.pathname !== '/callback' ||
        url.username ||
        url.password ||
        url.port ||
        url.hash
      )
        throw new AuthError('invalid_callback')
      const keys = [...url.searchParams.keys()]
      if (keys.some((key) => key !== 'code' && key !== 'state') || keys.length !== 2)
        throw new AuthError('invalid_callback')
      const code = url.searchParams.get('code')
      const state = url.searchParams.get('state')
      if (!code || code.length > 4_096 || !state || state.length > 512)
        throw new AuthError('invalid_callback')
      const transaction = [...pending.values()].find((candidate) =>
        equalSecret(state, candidate.state),
      )
      if (!transaction) {
        const hasActiveTransaction = [...pending.values()].some((candidate) => !candidate.consumed)
        throw new AuthError(hasActiveTransaction ? 'invalid_state' : 'callback_reused')
      }
      if (transaction.consumed) throw new AuthError('callback_reused')
      if (now() >= transaction.expiresAt) {
        transaction.consumed = true
        throw new AuthError('callback_expired')
      }
      if (transaction.attemptGeneration !== loginAttemptGeneration)
        throw new AuthError('auth_required')
      transaction.consumed = true
      const callbackAttemptGeneration = transaction.attemptGeneration
      const callbackUrl = new URL(config.callbackEndpoint)
      callbackUrl.searchParams.set('code', code)
      callbackUrl.searchParams.set('code_verifier', transaction.verifier)
      callbackUrl.searchParams.set('redirect_uri', config.redirectUri)
      callbackUrl.searchParams.set('client_id', config.clientId)
      const response = await doFetch(callbackUrl)
      const session = await parseResponse(response, now())
      return mutateSession(async () => {
        if (callbackAttemptGeneration !== loginAttemptGeneration)
          throw new AuthError('auth_required')
        await options.store.save(session)
        sessionRevision += 1
        return session
      })
    },
    async getAccountStatus() {
      return publicAccountStatus(await options.store.load())
    },
    async getAccessToken() {
      const session = await options.store.load()
      if (!session) return null
      if (session.expiresAt !== undefined && session.expiresAt <= now() + 60_000)
        return (await client.refresh()).accessToken
      return session.accessToken
    },
    refresh() {
      if (refreshFlight) return refreshFlight
      refreshFlight = (async () => {
        const refreshRevision = sessionRevision
        const current = await options.store.load()
        if (!current) throw new AuthError('auth_required')
        const response = await doFetch(config.refreshEndpoint, {
          method: 'POST',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ refresh_token: current.refreshToken }),
        })
        if (response.status === 401 || response.status === 400) {
          return mutateSession(async () => {
            if (refreshRevision === sessionRevision) {
              await options.store.clear()
              sessionRevision += 1
            }
            throw new AuthError('auth_required')
          })
        }
        const next = await parseResponse(response, now())
        return mutateSession(async () => {
          if (refreshRevision !== sessionRevision) throw new AuthError('auth_required')
          await options.store.save(next)
          return next
        })
      })().finally(() => {
        refreshFlight = null
      })
      return refreshFlight
    },
    async logout() {
      loginAttemptGeneration += 1
      pending.clear()
      return mutateSession(async () => {
        await options.store.clear()
        sessionRevision += 1
      })
    },
    async fetchWithAuth(request) {
      const token = await client.getAccessToken()
      if (!token) throw new AuthError('auth_required')
      let response = await request(token)
      if (response.status !== 401) return response
      const next = await client.refresh()
      response = await request(next.accessToken)
      return response
    },
    setNowForTesting(next) {
      now = next
    },
  }
  return client
}
