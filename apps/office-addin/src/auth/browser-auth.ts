import type { RuntimeConfig } from '../config.js'
import { createPkcePair, createState } from './pkce.js'

const STATE_KEY = 'wiswork.oauth.state'
const VERIFIER_KEY = 'wiswork.oauth.verifier'
const MAX_TOKEN_LENGTH = 16 * 1024

type BrowserAuthErrorCode =
  | 'invalid_callback'
  | 'token_exchange_failed'
  | 'unauthenticated'
  | 'invalid_destination'
  | 'refresh_failed'
  | 'unauthorized'

export class BrowserAuthError extends Error {
  constructor(readonly code: BrowserAuthErrorCode) {
    super(code)
    this.name = 'BrowserAuthError'
  }
}

interface TokenSession {
  accessToken: string
  refreshToken?: string
}

interface BrowserAuthDependencies {
  storage?: Storage
  fetch?: typeof globalThis.fetch
}

export interface BrowserAuth {
  startAuthorization(): Promise<string>
  consumeCallback(callbackUrl: string): Promise<void>
  authenticatedFetch(input: RequestInfo | URL, init?: RequestInit): Promise<Response>
  isAuthenticated(): boolean
  logout(): void
}

function singleParameter(url: URL, name: string): string | undefined {
  const values = url.searchParams.getAll(name)
  return values.length === 1 && values[0] ? values[0] : undefined
}

function exactCallbackLocation(actual: URL, expected: URL): boolean {
  return (
    actual.origin === expected.origin &&
    actual.pathname === expected.pathname &&
    actual.username === expected.username &&
    actual.password === expected.password &&
    actual.hash === ''
  )
}

async function readTokenSession(response: Response): Promise<TokenSession | undefined> {
  if (!response.ok) return undefined
  try {
    const value: unknown = await response.json()
    if (!value || typeof value !== 'object') return undefined
    const accessToken = Reflect.get(value, 'access_token')
    const refreshToken = Reflect.get(value, 'refresh_token')
    if (
      typeof accessToken !== 'string' ||
      accessToken.length === 0 ||
      accessToken.length > MAX_TOKEN_LENGTH
    ) {
      return undefined
    }
    if (
      refreshToken !== undefined &&
      (typeof refreshToken !== 'string' || refreshToken.length > MAX_TOKEN_LENGTH)
    ) {
      return undefined
    }
    return { accessToken, refreshToken }
  } catch {
    return undefined
  }
}

export function createBrowserAuth(
  config: RuntimeConfig,
  dependencies: BrowserAuthDependencies = {},
): BrowserAuth {
  const storage = dependencies.storage ?? globalThis.sessionStorage
  const fetchImplementation = dependencies.fetch ?? globalThis.fetch
  let session: TokenSession | undefined
  let sessionGeneration = 0
  let refreshFlight: { generation: number; promise: Promise<void> } | undefined

  function clearPkce(): void {
    storage.removeItem(STATE_KEY)
    storage.removeItem(VERIFIER_KEY)
  }

  function logout(): void {
    session = undefined
    sessionGeneration += 1
    clearPkce()
  }

  async function exchange(body: URLSearchParams, errorCode: BrowserAuthErrorCode) {
    let response: Response
    try {
      response = await fetchImplementation(config.tokenUrl, {
        method: 'POST',
        headers: new Headers({ 'content-type': 'application/x-www-form-urlencoded' }),
        body,
      })
    } catch {
      throw new BrowserAuthError(errorCode)
    }
    const nextSession = await readTokenSession(response)
    if (!nextSession) throw new BrowserAuthError(errorCode)
    return nextSession
  }

  function refreshSession(generation: number): Promise<void> {
    if (sessionGeneration !== generation) return Promise.resolve()
    if (refreshFlight?.generation === generation) return refreshFlight.promise
    if (!session?.refreshToken) {
      logout()
      return Promise.reject(new BrowserAuthError('refresh_failed'))
    }

    const refreshToken = session.refreshToken
    const promise = (async () => {
      try {
        const refreshed = await exchange(
          new URLSearchParams({
            grant_type: 'refresh_token',
            refresh_token: refreshToken,
            client_id: config.clientId,
          }),
          'refresh_failed',
        )
        if (sessionGeneration === generation) {
          session = { ...refreshed, refreshToken: refreshed.refreshToken ?? refreshToken }
          sessionGeneration += 1
        }
      } catch {
        if (sessionGeneration === generation) logout()
        throw new BrowserAuthError('refresh_failed')
      }
    })()
    refreshFlight = { generation, promise }
    const clearFlight = () => {
      if (refreshFlight?.promise === promise) refreshFlight = undefined
    }
    promise.then(clearFlight, clearFlight)
    return promise
  }

  return {
    async startAuthorization() {
      clearPkce()
      const [{ verifier, challenge, method }, state] = await Promise.all([
        createPkcePair(),
        Promise.resolve(createState()),
      ])
      storage.setItem(STATE_KEY, state)
      storage.setItem(VERIFIER_KEY, verifier)

      const authorizationUrl = new URL(config.authorizationUrl)
      authorizationUrl.searchParams.set('response_type', 'code')
      authorizationUrl.searchParams.set('client_id', config.clientId)
      authorizationUrl.searchParams.set('redirect_uri', config.callbackUrl)
      authorizationUrl.searchParams.set('state', state)
      authorizationUrl.searchParams.set('code_challenge', challenge)
      authorizationUrl.searchParams.set('code_challenge_method', method)
      return authorizationUrl.toString()
    },

    async consumeCallback(callbackUrl) {
      const expectedState = storage.getItem(STATE_KEY)
      const verifier = storage.getItem(VERIFIER_KEY)
      clearPkce()

      let callback: URL
      try {
        callback = new URL(callbackUrl)
      } catch {
        throw new BrowserAuthError('invalid_callback')
      }
      const code = singleParameter(callback, 'code')
      const state = singleParameter(callback, 'state')
      const issuers = callback.searchParams.getAll('iss')
      const issuerIsValid =
        issuers.length === 0 || (issuers.length === 1 && issuers[0] === config.issuer)

      if (
        !expectedState ||
        !verifier ||
        !code ||
        !state ||
        state !== expectedState ||
        !issuerIsValid ||
        !exactCallbackLocation(callback, new URL(config.callbackUrl))
      ) {
        throw new BrowserAuthError('invalid_callback')
      }

      const exchangedSession = await exchange(
        new URLSearchParams({
          grant_type: 'authorization_code',
          code,
          redirect_uri: config.callbackUrl,
          client_id: config.clientId,
          code_verifier: verifier,
        }),
        'token_exchange_failed',
      )
      session = exchangedSession
      sessionGeneration += 1
    },

    async authenticatedFetch(input, init = {}) {
      if (!session) throw new BrowserAuthError('unauthenticated')
      let destination: string
      try {
        destination = input instanceof Request ? input.url : new URL(input).href
      } catch {
        throw new BrowserAuthError('invalid_destination')
      }
      if (destination !== new URL(config.messagesUrl).href) {
        throw new BrowserAuthError('invalid_destination')
      }

      const authorizedRequest = (accessToken: string) => {
        const headers = new Headers(init.headers)
        headers.set('authorization', `Bearer ${accessToken}`)
        return fetchImplementation(input, { ...init, headers })
      }

      const initialGeneration = sessionGeneration
      let response = await authorizedRequest(session.accessToken)
      if (response.status !== 401) return response

      await refreshSession(initialGeneration)
      if (!session) throw new BrowserAuthError('refresh_failed')
      const retryGeneration = sessionGeneration
      response = await authorizedRequest(session.accessToken)
      if (response.status === 401) {
        if (sessionGeneration === retryGeneration) logout()
        throw new BrowserAuthError('unauthorized')
      }
      return response
    },

    isAuthenticated() {
      return session !== undefined
    },

    logout,
  }
}
