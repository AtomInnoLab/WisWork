import { describe, expect, it, vi } from 'vitest'
import { BrowserAuthError, createBrowserAuth } from '../src/auth/browser-auth.js'
import type { RuntimeConfig } from '../src/config.js'

const config: RuntimeConfig = {
  authorizationUrl: 'https://gateway.example/oauth/authorize',
  tokenUrl: 'https://gateway.example/oauth/token',
  callbackUrl: 'https://localhost:3000/oauth/callback',
  clientId: 'office-addin',
  issuer: 'https://gateway.example',
  messagesUrl: 'https://gateway.example/v1/messages',
}

class MemoryStorage implements Storage {
  readonly values = new Map<string, string>()
  get length() {
    return this.values.size
  }
  clear() {
    this.values.clear()
  }
  getItem(key: string) {
    return this.values.get(key) ?? null
  }
  key(index: number) {
    return [...this.values.keys()][index] ?? null
  }
  removeItem(key: string) {
    this.values.delete(key)
  }
  setItem(key: string, value: string) {
    this.values.set(key, value)
  }
}

function tokenResponse(accessToken = 'access-secret', refreshToken = 'refresh-secret') {
  return new Response(JSON.stringify({ access_token: accessToken, refresh_token: refreshToken }), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  })
}

describe('browser OAuth', () => {
  it('starts authorization with S256 and session-only one-time material', async () => {
    const storage = new MemoryStorage()
    const auth = createBrowserAuth(config, { storage, fetch: vi.fn() })

    const authorizationUrl = new URL(await auth.startAuthorization())

    expect(authorizationUrl.origin + authorizationUrl.pathname).toBe(config.authorizationUrl)
    expect(authorizationUrl.searchParams.get('client_id')).toBe(config.clientId)
    expect(authorizationUrl.searchParams.get('redirect_uri')).toBe(config.callbackUrl)
    expect(authorizationUrl.searchParams.get('scope')).toBe('openid profile email offline_access')
    expect(authorizationUrl.searchParams.get('code_challenge_method')).toBe('S256')
    expect(authorizationUrl.searchParams.get('code_challenge')).toMatch(/^[A-Za-z0-9_-]{43}$/)
    expect(authorizationUrl.searchParams.get('state')).toBeTruthy()
    expect(storage.length).toBe(2)
  })

  it('consumes the callback once and exchanges the verifier', async () => {
    const storage = new MemoryStorage()
    const fetch = vi.fn().mockResolvedValue(tokenResponse())
    const auth = createBrowserAuth(config, { storage, fetch })
    const authorizationUrl = new URL(await auth.startAuthorization())
    const state = authorizationUrl.searchParams.get('state')!
    const callback = `${config.callbackUrl}?code=one-time-code&state=${state}&iss=${encodeURIComponent(config.issuer)}`

    await expect(auth.consumeCallback(callback)).resolves.toBeUndefined()
    expect(auth.isAuthenticated()).toBe(true)
    expect(storage.length).toBe(0)
    const body = new URLSearchParams(fetch.mock.calls[0][1].body)
    expect(fetch.mock.calls[0][1].redirect).toBe('error')
    expect(body.get('code')).toBe('one-time-code')
    expect(body.get('code_verifier')).toBeTruthy()
    expect(body.get('client_secret')).toBeNull()
    await expect(auth.consumeCallback(callback)).rejects.toMatchObject({ code: 'invalid_callback' })
  })

  it.each([
    ['missing code', (state: string) => `${config.callbackUrl}?state=${state}`],
    ['duplicate code', (state: string) => `${config.callbackUrl}?code=a&code=b&state=${state}`],
    ['missing state', () => `${config.callbackUrl}?code=a`],
    [
      'duplicate state',
      (state: string) => `${config.callbackUrl}?code=a&state=${state}&state=${state}`,
    ],
    ['wrong state', () => `${config.callbackUrl}?code=a&state=wrong`],
    [
      'wrong issuer',
      (state: string) =>
        `${config.callbackUrl}?code=a&state=${state}&iss=https%3A%2F%2Fevil.example`,
    ],
    [
      'wrong callback',
      (state: string) => `https://evil.example/oauth/callback?code=a&state=${state}`,
    ],
  ])('rejects %s and consumes stored PKCE material', async (_name, callbackFor) => {
    const storage = new MemoryStorage()
    const fetch = vi.fn()
    const auth = createBrowserAuth(config, { storage, fetch })
    const authorizationUrl = new URL(await auth.startAuthorization())
    const callback = callbackFor(authorizationUrl.searchParams.get('state')!)

    await expect(auth.consumeCallback(callback)).rejects.toMatchObject({ code: 'invalid_callback' })
    expect(storage.length).toBe(0)
    expect(fetch).not.toHaveBeenCalled()
  })

  it('refreshes exactly once after 401 and retries with the new token', async () => {
    const storage = new MemoryStorage()
    const fetch = vi
      .fn()
      .mockResolvedValueOnce(tokenResponse())
      .mockResolvedValueOnce(new Response('do not expose me', { status: 401 }))
      .mockResolvedValueOnce(tokenResponse('new-access', 'new-refresh'))
      .mockResolvedValueOnce(new Response('ok', { status: 200 }))
    const auth = createBrowserAuth(config, { storage, fetch })
    const start = new URL(await auth.startAuthorization())
    await auth.consumeCallback(
      `${config.callbackUrl}?code=a&state=${start.searchParams.get('state')}`,
    )

    const response = await auth.authenticatedFetch(config.messagesUrl)

    expect(response.status).toBe(200)
    expect(fetch).toHaveBeenCalledTimes(4)
    expect(fetch.mock.calls[1][1].headers.get('authorization')).toBe('Bearer access-secret')
    expect(fetch.mock.calls[2][1].redirect).toBe('error')
    expect(fetch.mock.calls[3][1].headers.get('authorization')).toBe('Bearer new-access')
  })

  it('shares one refresh exchange across concurrent 401 responses', async () => {
    const storage = new MemoryStorage()
    let releaseRefresh!: (response: Response) => void
    const refreshResponse = new Promise<Response>((resolve) => {
      releaseRefresh = resolve
    })
    const fetch = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      if (String(input) === config.tokenUrl) {
        const body = new URLSearchParams(init?.body as string)
        if (body.get('grant_type') === 'refresh_token') return refreshResponse
        return tokenResponse()
      }
      const authorization = new Headers(init?.headers).get('authorization')
      return new Response('', { status: authorization === 'Bearer new-access' ? 200 : 401 })
    })
    const auth = createBrowserAuth(config, { storage, fetch })
    const start = new URL(await auth.startAuthorization())
    await auth.consumeCallback(
      `${config.callbackUrl}?code=a&state=${start.searchParams.get('state')}`,
    )

    const first = auth.authenticatedFetch(config.messagesUrl)
    const second = auth.authenticatedFetch(config.messagesUrl)
    await vi.waitFor(() => {
      expect(fetch).toHaveBeenCalledTimes(4)
    })
    releaseRefresh(tokenResponse('new-access', 'new-refresh'))

    await expect(Promise.all([first, second])).resolves.toMatchObject([
      { status: 200 },
      { status: 200 },
    ])
    const refreshCalls = fetch.mock.calls.filter((call) => {
      if (String(call[0]) !== config.tokenUrl) return false
      return new URLSearchParams(call[1]?.body as string).get('grant_type') === 'refresh_token'
    })
    expect(refreshCalls).toHaveLength(1)
  })

  it('does not let a stale failed refresh clear a newer login session', async () => {
    const storage = new MemoryStorage()
    let rejectRefresh!: (error: Error) => void
    const refreshResponse = new Promise<Response>((_resolve, reject) => {
      rejectRefresh = reject
    })
    let authorizationExchanges = 0
    const fetch = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      if (String(input) === config.tokenUrl) {
        const body = new URLSearchParams(init?.body as string)
        if (body.get('grant_type') === 'refresh_token') return refreshResponse
        authorizationExchanges += 1
        return authorizationExchanges === 1
          ? tokenResponse('old-access', 'old-refresh')
          : tokenResponse('new-access', 'new-refresh')
      }
      const authorization = new Headers(init?.headers).get('authorization')
      return new Response('', { status: authorization === 'Bearer new-access' ? 200 : 401 })
    })
    const auth = createBrowserAuth(config, { storage, fetch })
    const firstStart = new URL(await auth.startAuthorization())
    await auth.consumeCallback(
      `${config.callbackUrl}?code=first&state=${firstStart.searchParams.get('state')}`,
    )
    const staleRequest = auth.authenticatedFetch(config.messagesUrl)
    await vi.waitFor(() => expect(fetch).toHaveBeenCalledTimes(3))

    const secondStart = new URL(await auth.startAuthorization())
    await auth.consumeCallback(
      `${config.callbackUrl}?code=second&state=${secondStart.searchParams.get('state')}`,
    )
    rejectRefresh(new Error('old refresh failed'))

    await expect(staleRequest).rejects.toEqual(new BrowserAuthError('refresh_failed'))
    expect(auth.isAuthenticated()).toBe(true)
    await expect(auth.authenticatedFetch(config.messagesUrl)).resolves.toMatchObject({
      status: 200,
    })
  })

  it('refuses to attach authorization outside the exact messages endpoint', async () => {
    const storage = new MemoryStorage()
    const fetch = vi.fn().mockResolvedValue(tokenResponse())
    const auth = createBrowserAuth(config, { storage, fetch })
    const start = new URL(await auth.startAuthorization())
    await auth.consumeCallback(
      `${config.callbackUrl}?code=a&state=${start.searchParams.get('state')}`,
    )
    fetch.mockClear()

    await expect(auth.authenticatedFetch('https://attacker.example/collect')).rejects.toEqual(
      new BrowserAuthError('invalid_destination'),
    )
    await expect(auth.authenticatedFetch(`${config.messagesUrl}/another-path`)).rejects.toEqual(
      new BrowserAuthError('invalid_destination'),
    )
    expect(fetch).not.toHaveBeenCalled()
  })

  it('logs out after the retried request returns 401', async () => {
    const storage = new MemoryStorage()
    const fetch = vi
      .fn()
      .mockResolvedValueOnce(tokenResponse())
      .mockResolvedValueOnce(new Response('first upstream secret', { status: 401 }))
      .mockResolvedValueOnce(tokenResponse('new-access', 'new-refresh'))
      .mockResolvedValueOnce(new Response('second upstream secret', { status: 401 }))
    const auth = createBrowserAuth(config, { storage, fetch })
    const authLost = vi.fn()
    auth.subscribeAuthLoss(authLost)
    const start = new URL(await auth.startAuthorization())
    await auth.consumeCallback(
      `${config.callbackUrl}?code=a&state=${start.searchParams.get('state')}`,
    )

    await expect(auth.authenticatedFetch(config.messagesUrl)).rejects.toEqual(
      new BrowserAuthError('unauthorized'),
    )
    expect(auth.isAuthenticated()).toBe(false)
    expect(authLost).toHaveBeenCalledOnce()
    expect(fetch).toHaveBeenCalledTimes(4)
  })

  it('does not follow token endpoint redirects with OAuth request bodies', async () => {
    const storage = new MemoryStorage()
    const fetch = vi.fn().mockResolvedValue(new Response(null, { status: 307 }))
    const auth = createBrowserAuth(config, { storage, fetch })
    const start = new URL(await auth.startAuthorization())

    await expect(
      auth.consumeCallback(`${config.callbackUrl}?code=a&state=${start.searchParams.get('state')}`),
    ).rejects.toEqual(new BrowserAuthError('token_exchange_failed'))
    expect(fetch.mock.calls[0][1].redirect).toBe('error')
  })

  it('uses safe stable errors without exposing token endpoint bodies', async () => {
    const storage = new MemoryStorage()
    const fetch = vi
      .fn()
      .mockResolvedValue(new Response('access-secret upstream-body', { status: 500 }))
    const auth = createBrowserAuth(config, { storage, fetch })
    const start = new URL(await auth.startAuthorization())

    await expect(
      auth.consumeCallback(`${config.callbackUrl}?code=a&state=${start.searchParams.get('state')}`),
    ).rejects.toEqual(new BrowserAuthError('token_exchange_failed'))
    expect(auth.isAuthenticated()).toBe(false)
  })

  it('rejects oversized access and refresh token values', async () => {
    const oversized = 'x'.repeat(16 * 1024 + 1)

    for (const payload of [
      { access_token: oversized, refresh_token: 'refresh' },
      { access_token: 'access', refresh_token: oversized },
    ]) {
      const storage = new MemoryStorage()
      const fetch = vi.fn().mockResolvedValue(
        new Response(JSON.stringify(payload), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        }),
      )
      const auth = createBrowserAuth(config, { storage, fetch })
      const start = new URL(await auth.startAuthorization())

      await expect(
        auth.consumeCallback(
          `${config.callbackUrl}?code=a&state=${start.searchParams.get('state')}`,
        ),
      ).rejects.toEqual(new BrowserAuthError('token_exchange_failed'))
      expect(auth.isAuthenticated()).toBe(false)
    }
  })

  it('clears all session material on logout', async () => {
    const storage = new MemoryStorage()
    const auth = createBrowserAuth(config, {
      storage,
      fetch: vi.fn().mockResolvedValue(tokenResponse()),
    })
    const lost = vi.fn()
    const unsubscribe = auth.subscribeAuthLoss(lost)
    const start = new URL(await auth.startAuthorization())
    await auth.consumeCallback(
      `${config.callbackUrl}?code=a&state=${start.searchParams.get('state')}`,
    )

    auth.logout()
    expect(storage.length).toBe(0)
    expect(auth.isAuthenticated()).toBe(false)
    expect(lost).toHaveBeenCalledOnce()
    unsubscribe()
    auth.logout()
    expect(lost).toHaveBeenCalledOnce()
  })
})
