import { expect, it, vi } from 'vitest'
import { createBrowserAuthClient, createMemorySessionStore } from '../src/browser.js'
const config = {
  clientId: 'office-registered-client',
  redirectUri: 'https://office.example/auth/callback.html',
}
const payload = {
  token: 'access-secret',
  refresh_token: 'refresh-secret',
  user_id: 'user',
  expires_in: 60,
}
const callback = (state: string, extra = '') =>
  `${config.redirectUri}?code=code&state=${state}${extra}`
it('uses browser state and unchanged confidential gateway exchange with exact callback URL', async () => {
  const fetch = vi.fn(async (_input?: RequestInfo | URL, _init?: RequestInit) =>
      Response.json(payload),
    ),
    store = createMemorySessionStore()
  const client = createBrowserAuthClient({ config, store, fetch })
  const request = client.createAuthorizationRequest()
  expect(request.state).toMatch(/^[A-Za-z0-9_-]{43}$/)
  expect(new URL(request.url).searchParams.get('client_id')).toBe(config.clientId)
  await expect(
    client.consumeCallback(callback(request.state).replace('office.example', 'other.example')),
  ).rejects.toMatchObject({ code: 'invalid_callback' })
  await client.consumeCallback(callback(request.state))
  expect(fetch).toHaveBeenCalledTimes(1)
  const url = new URL(String(fetch.mock.calls[0]![0]))
  expect(url.searchParams.get('redirect_uri')).toBe(config.redirectUri)
  expect(url.searchParams.get('code')).toBe('code')
  expect(url.searchParams.has('code_verifier')).toBe(false)
  expect(await client.getAccountStatus()).toEqual({ loggedIn: true, userId: 'user' })
  await expect(client.consumeCallback(callback(request.state))).rejects.toMatchObject({
    code: 'callback_reused',
  })
})
it('keeps memory sessions cloned and clears without persistent browser storage', async () => {
  const store = createMemorySessionStore(),
    session = { accessToken: 'a', refreshToken: 'r', userId: 'u' }
  await store.save(session)
  session.accessToken = 'changed'
  const first = await store.load()
  expect(first?.accessToken).toBe('a')
  first!.userId = 'changed'
  expect((await store.load())?.userId).toBe('u')
  await store.clear()
  expect(await store.load()).toBeNull()
})
it.each([
  ['wrong path', 'https://office.example/other?'],
  ['wrong port', 'https://office.example:444/auth/callback.html?'],
  ['insecure scheme', 'http://office.example/auth/callback.html?'],
  ['credentials', 'https://user@office.example/auth/callback.html?'],
])('rejects %s before gateway exchange', async (_label, prefix) => {
  const fetch = vi.fn(async () => Response.json(payload)),
    client = createBrowserAuthClient({ config, store: createMemorySessionStore(), fetch })
  const { state } = client.createAuthorizationRequest()
  await expect(client.consumeCallback(`${prefix}code=code&state=${state}`)).rejects.toMatchObject({
    code: 'invalid_callback',
  })
  expect(fetch).not.toHaveBeenCalled()
})
it.each([
  '&iss=https://wrong.example',
  '&state=another',
  '&code=another',
  '&unknown=private',
  '#fragment',
])('rejects callback metadata %s without consuming the valid transaction', async (extra) => {
  const fetch = vi.fn(async () => Response.json(payload)),
    client = createBrowserAuthClient({ config, store: createMemorySessionStore(), fetch })
  const { state } = client.createAuthorizationRequest()
  await expect(client.consumeCallback(callback(state, extra))).rejects.toMatchObject({
    code: 'invalid_callback',
  })
  expect(fetch).not.toHaveBeenCalled()
  await client.consumeCallback(callback(state))
})
it('rejects wrong state and accepts the exact configured issuer', async () => {
  const client = createBrowserAuthClient({
      config,
      store: createMemorySessionStore(),
      fetch: async () => Response.json(payload),
    }),
    { state } = client.createAuthorizationRequest()
  await expect(client.consumeCallback(callback('different'))).rejects.toMatchObject({
    code: 'invalid_state',
  })
  await client.consumeCallback(callback(state, '&iss=https://auth.wispaper.ai/oidc'))
})
it('rejects a late previous login attempt and preserves the newer login', async () => {
  let sequence = 0,
    release!: (v: Response) => void
  const fetch = vi.fn(async () => new Promise<Response>((r) => (release = r))),
    store = createMemorySessionStore(),
    client = createBrowserAuthClient({
      config,
      store,
      fetch,
      randomBytes: (size) => new Uint8Array(size).fill(++sequence),
    })
  const first = client.createAuthorizationRequest(),
    pending = client.consumeCallback(callback(first.state))
  const rejected = expect(pending).rejects.toMatchObject({ code: 'auth_required' })
  await vi.waitFor(() => expect(release).toBeDefined())
  const next = client.createAuthorizationRequest()
  release(Response.json(payload))
  await rejected
  expect(await store.load()).toBeNull()
  fetch.mockImplementation(async () => Response.json({ ...payload, user_id: 'newer' }))
  await client.consumeCallback(callback(next.state))
  expect(await client.getAccountStatus()).toEqual({ loggedIn: true, userId: 'newer' })
})
it('logout prevents a pending callback exchange from reviving the session', async () => {
  let release!: (v: Response) => void
  const store = createMemorySessionStore(),
    client = createBrowserAuthClient({
      config,
      store,
      fetch: async () => new Promise<Response>((r) => (release = r)),
    })
  const pending = client.consumeCallback(callback(client.createAuthorizationRequest().state)),
    rejected = expect(pending).rejects.toMatchObject({ code: 'auth_required' })
  await vi.waitFor(() => expect(release).toBeDefined())
  await client.logout()
  release(Response.json(payload))
  await rejected
  expect(await client.getAccountStatus()).toEqual({ loggedIn: false })
  expect(await client.getAccessToken()).toBeNull()
})
it.each([200, 401])(
  'logout keeps a late refresh %i from restoring or changing session state',
  async (status) => {
    let release!: (v: Response) => void
    const store = createMemorySessionStore()
    await store.save({ accessToken: 'expired', refreshToken: 'r', userId: 'u', expiresAt: 0 })
    const fetch = vi.fn(
        async (_input?: RequestInfo | URL, _init?: RequestInit) =>
          new Promise<Response>((r) => (release = r)),
      ),
      client = createBrowserAuthClient({ config, store, fetch })
    const pending = client.refresh(),
      rejected = expect(pending).rejects.toMatchObject({ code: 'auth_required' })
    await vi.waitFor(() => expect(release).toBeDefined())
    const [url, init] = fetch.mock.calls[0]!
    expect(init?.method).toBe('POST')
    expect(JSON.parse(String(init?.body))).toEqual({ refresh_token: 'r' })
    expect(new URL(String(url)).searchParams.get('redirect_uri')).toBe(config.redirectUri)
    await client.logout()
    release(Response.json(payload, { status }))
    await rejected
    expect(await store.load()).toBeNull()
  },
)
it('requires an explicit HTTPS browser redirect rather than inventing a web client', () => {
  for (const redirectUri of [
    undefined,
    'wiswork://oauth/callback',
    'http://office.example/callback',
    'https://office.example/callback#fragment',
    'https://office.example/callback?code=preexisting',
  ])
    expect(() =>
      createBrowserAuthClient({
        store: createMemorySessionStore(),
        config: { ...config, redirectUri },
      }),
    ).toThrow('invalid_callback')
})
it('matches an explicitly configured HTTPS port without weakening Node callbacks', async () => {
  const redirectUri = 'https://office.example:444/auth/callback.html',
    client = createBrowserAuthClient({
      config: { ...config, redirectUri },
      store: createMemorySessionStore(),
      fetch: async () => Response.json(payload),
    }),
    { state } = client.createAuthorizationRequest()
  await client.consumeCallback(`${redirectUri}?code=code&state=${state}`)
  expect((await client.getAccountStatus()).loggedIn).toBe(true)
})

import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
it('browser dependency graph contains no Node entry, crypto, Buffer, or persistent browser cache', () => {
  const seen = new Set<string>()
  const visit = (file: string) => {
    if (seen.has(file)) return
    seen.add(file)
    const text = readFileSync(file, 'utf8')
    expect(text).not.toMatch(/node:|\bBuffer\b|localStorage|sessionStorage/)
    for (const match of text.matchAll(/(?:from\s+|import\s*)['"]([^'"]+)['"]/g)) {
      const path = match[1]!
      expect(path).not.toBe('./index')
      if (path.startsWith('.'))
        visit(resolve(dirname(file), path.endsWith('.ts') ? path : `${path}.ts`))
    }
  }
  visit(fileURLToPath(new URL('../src/browser.ts', import.meta.url)))
  expect(seen.size).toBe(4)
})
it.each([undefined, '', '   '])(
  'requires an explicitly configured nonempty browser clientId (%s)',
  (clientId) => {
    expect(() =>
      createBrowserAuthClient({
        store: createMemorySessionStore(),
        config: { redirectUri: config.redirectUri, clientId },
      }),
    ).toThrow('invalid_callback')
  },
)
