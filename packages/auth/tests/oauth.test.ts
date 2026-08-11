import { createHash } from 'node:crypto'
import { describe, expect, it, vi } from 'vitest'
import { AuthError, createAuthClient, type AuthSession, type SessionStore } from '../src/index'

const session: AuthSession = {
  accessToken: 'access-secret',
  refreshToken: 'refresh-secret',
  userId: 'user-1',
  email: 'person@example.com',
  expiresAt: 20_000,
}

function memoryStore(): SessionStore {
  let value: AuthSession | null = null
  return {
    load: vi.fn(async () => value),
    save: vi.fn(async (next) => void (value = next)),
    clear: vi.fn(async () => void (value = null)),
  }
}

function fixture(now = 1_000) {
  const fetch = vi.fn(
    async (_input?: RequestInfo | URL, _init?: RequestInit) =>
      new Response(
        JSON.stringify({
          token: session.accessToken,
          refresh_token: session.refreshToken,
          user_id: session.userId,
          email: session.email,
          expires_in: 19,
        }),
        { status: 200, headers: { 'content-type': 'application/json' } },
      ),
  )
  const client = createAuthClient({
    fetch,
    store: memoryStore(),
    now: () => now,
    randomBytes: (size) => new Uint8Array(size).fill(7),
  })
  return { client, fetch }
}

describe('OAuth authorization and callback', () => {
  it('bounds pending transactions and retains a bounded reused-state classification', async () => {
    let sequence = 0
    const fetch = vi.fn(
      async () =>
        new Response(
          JSON.stringify({ token: 'a', refresh_token: 'r', user_id: 'u', expires_in: 60 }),
          { status: 200, headers: { 'content-type': 'application/json' } },
        ),
    )
    const client = createAuthClient({
      fetch,
      store: memoryStore(),
      randomBytes: (size) => new Uint8Array(size).fill(++sequence),
    })
    const requests = Array.from({ length: 40 }, () => client.createAuthorizationRequest())

    await expect(
      client.consumeCallback(`wiswork://oauth/callback?code=old&state=${requests[0]!.state}`),
    ).rejects.toMatchObject({ code: 'invalid_state' })

    const current = requests.at(-1)!
    const callback = `wiswork://oauth/callback?code=ok&state=${current.state}`
    await client.consumeCallback(callback)
    await expect(client.consumeCallback(callback)).rejects.toMatchObject({
      code: 'callback_reused',
    })

    for (let index = 0; index < 80; index += 1) {
      const request = client.createAuthorizationRequest()
      await client.consumeCallback(`wiswork://oauth/callback?code=${index}&state=${request.state}`)
    }
    await expect(client.consumeCallback(callback)).rejects.toMatchObject({ code: 'invalid_state' })
  })

  it('removes expired transactions during normal authorization activity', async () => {
    let now = 1_000
    let sequence = 0
    const client = createAuthClient({
      fetch: vi.fn(),
      store: memoryStore(),
      now: () => now,
      randomBytes: (size) => new Uint8Array(size).fill(++sequence),
    })
    const expired = client.createAuthorizationRequest()
    now += 10 * 60_000
    client.createAuthorizationRequest()
    await expect(
      client.consumeCallback(`wiswork://oauth/callback?code=old&state=${expired.state}`),
    ).rejects.toMatchObject({ code: 'callback_expired' })
  })

  it('uses the configured Logto endpoint and PKCE S256', () => {
    const { client } = fixture()
    const request = client.createAuthorizationRequest()
    const url = new URL(request.url)
    expect(url.origin + url.pathname).toBe('https://auth.dev.wispaper.ai/oidc/auth')
    expect(url.searchParams.get('client_id')).toBe('y3xpwx3ytskxf66p0wztm')
    expect(url.searchParams.get('redirect_uri')).toBe('wiswork://oauth/callback')
    expect(url.searchParams.get('scope')).toBe('openid profile email offline_access')
    expect(url.searchParams.get('code_challenge_method')).toBe('S256')
    const expected = createHash('sha256')
      .update(Buffer.alloc(32, 7).toString('base64url'))
      .digest('base64url')
    expect(url.searchParams.get('code_challenge')).toBe(expected)
    expect(Buffer.from(request.state, 'base64url')).toHaveLength(32)
  })

  it('rejects state mismatch without consuming the valid transaction', async () => {
    const { client } = fixture()
    const { state } = client.createAuthorizationRequest()
    await expect(
      client.consumeCallback('wiswork://oauth/callback?code=ok&state=wrong'),
    ).rejects.toMatchObject({ code: 'invalid_state' })
    await expect(
      client.consumeCallback(`wiswork://oauth/callback?code=ok&state=${state}`),
    ).resolves.toMatchObject({ userId: 'user-1' })
  })

  it('allows a callback transaction to be consumed only once', async () => {
    const { client } = fixture()
    const { state } = client.createAuthorizationRequest()
    const url = `wiswork://oauth/callback?code=ok&state=${state}`
    await client.consumeCallback(url)
    await expect(client.consumeCallback(url)).rejects.toMatchObject({ code: 'callback_reused' })
  })

  it('rejects transactions at the exact expiry timestamp', async () => {
    let now = 1_000
    const { client } = fixture(now)
    const { state } = client.createAuthorizationRequest()
    now += 10 * 60_000
    client.setNowForTesting(() => now)
    await expect(
      client.consumeCallback(`wiswork://oauth/callback?code=ok&state=${state}`),
    ).rejects.toMatchObject({ code: 'callback_expired' })
  })

  it('exchanges the code with the gateway using GET without leaking secrets in errors', async () => {
    const { client, fetch } = fixture()
    const { state } = client.createAuthorizationRequest()
    await client.consumeCallback(`wiswork://oauth/callback?code=one-time-code&state=${state}`)
    const [input, init] = fetch.mock.calls[0]!
    const requestUrl = new URL(String(input))
    expect(requestUrl.origin + requestUrl.pathname).toBe(
      'https://gateway.dev.wispaper.ai/api/v1/auth/user/callback',
    )
    expect(requestUrl.searchParams.get('code')).toBe('one-time-code')
    expect(requestUrl.searchParams.get('code_verifier')).toBe(
      Buffer.alloc(32, 7).toString('base64url'),
    )
    expect(requestUrl.searchParams.get('redirect_uri')).toBe('wiswork://oauth/callback')
    expect(requestUrl.searchParams.get('client_id')).toBe('y3xpwx3ytskxf66p0wztm')
    expect([...requestUrl.searchParams.keys()].sort()).toEqual([
      'client_id',
      'code',
      'code_verifier',
      'redirect_uri',
    ])
    expect(init).toBeUndefined()

    const rejected = createAuthClient({
      store: memoryStore(),
      fetch: vi.fn(async () => new Response('one-time-code access-secret', { status: 500 })),
    })
    const rejectedRequest = rejected.createAuthorizationRequest()
    const error = await rejected
      .consumeCallback(`wiswork://oauth/callback?code=one-time-code&state=${rejectedRequest.state}`)
      .catch((reason: unknown) => reason)
    expect(String(error)).toBe('AuthError: network_error')
    expect(error).toMatchObject({
      diagnostic: { stage: 'callback_exchange', httpStatus: 500 },
    })
    expect(String(error)).not.toContain('one-time-code')
    expect(String(error)).not.toContain('access-secret')
  })

  it.each([
    'https://oauth/callback?code=ok&state=x',
    'wiswork://other/callback?code=ok&state=x',
    'wiswork://oauth/not-callback?code=ok&state=x',
  ])('rejects an invalid callback URL: %s', async (url) => {
    const { client, fetch } = fixture()
    client.createAuthorizationRequest()
    await expect(client.consumeCallback(url)).rejects.toBeInstanceOf(AuthError)
    expect(fetch).not.toHaveBeenCalled()
  })
})
