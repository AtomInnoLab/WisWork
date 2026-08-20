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
    expect(fetch.mock.calls[3][1].headers.get('authorization')).toBe('Bearer new-access')
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
    const start = new URL(await auth.startAuthorization())
    await auth.consumeCallback(
      `${config.callbackUrl}?code=a&state=${start.searchParams.get('state')}`,
    )

    await expect(auth.authenticatedFetch(config.messagesUrl)).rejects.toEqual(
      new BrowserAuthError('unauthorized'),
    )
    expect(auth.isAuthenticated()).toBe(false)
    expect(fetch).toHaveBeenCalledTimes(4)
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

  it('clears all session material on logout', async () => {
    const storage = new MemoryStorage()
    const auth = createBrowserAuth(config, { storage, fetch: vi.fn() })
    await auth.startAuthorization()

    auth.logout()
    expect(storage.length).toBe(0)
    expect(auth.isAuthenticated()).toBe(false)
  })
})
