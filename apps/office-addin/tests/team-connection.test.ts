import { expect, it, vi } from 'vitest'
import {
  createBrowserAuthClient,
  createMemorySessionStore,
  type AuthClient,
} from '@wiswork/auth/browser'
import type { OfficeRelaySession, OfficeRelaySnapshot } from '../src/relay/session.js'
import { createOfficeTeamConnection } from '../src/agent/team-connection.js'
function fixture() {
  let loggedIn = false,
    account = 'account-a',
    relay: OfficeRelaySnapshot = { status: 'offline' }
  const listeners = new Set<() => void>()
  const auth = {
    createAuthorizationRequest: vi.fn(() => ({
      url: 'https://auth.example/authorize',
      state: 'state',
    })),
    consumeCallback: vi.fn(async () => {
      loggedIn = true
      return { accessToken: 'secret-access', refreshToken: 'secret-refresh', userId: account }
    }),
    getValidAccountStatus: vi.fn(async () => ({
      loggedIn,
      ...(loggedIn ? { userId: account, email: 'test@example.com' } : {}),
    })),
    getAccessToken: vi.fn(async () => (loggedIn ? 'secret-access' : null)),
    logout: vi.fn(async () => {
      loggedIn = false
    }),
  } as unknown as AuthClient
  const session: OfficeRelaySession = {
    snapshot: () => relay,
    subscribe: (l) => {
      listeners.add(l)
      return () => listeners.delete(l)
    },
    connect: vi.fn(async () => {
      relay = { status: 'connected', capabilities: ['presentation-team.v1'] }
      listeners.forEach((l) => l())
    }),
    disconnect: vi.fn(() => {
      relay = { status: 'offline' }
      listeners.forEach((l) => l())
    }),
    capabilityFetch: vi.fn(async () => Response.json({ identity: {} })),
    authenticatedFetch: vi.fn(),
    sendDiagnostic: vi.fn(),
    diagnosticSessionId: () => undefined,
  }
  const loginDialog = vi.fn(async () => 'https://addin.example/callback?code=code&state=state'),
    onUnavailable = vi.fn()
  let provider!: () => Promise<string | null>
  const createSession = vi.fn((options: any) => {
    expect(options.capabilities).toEqual(['presentation-team.v1'])
    provider = options.getTeamAccessToken
    return session
  })
  const connection = createOfficeTeamConnection({ auth, loginDialog, createSession, onUnavailable })
  return {
    auth,
    session,
    connection,
    loginDialog,
    createSession,
    onUnavailable,
    provider: () => provider(),
    setAccount() {
      account = 'account-b'
    },
    emit(state: OfficeRelaySnapshot) {
      relay = state
      listeners.forEach((l) => l())
    },
  }
}
it('signs in and connects only the independent team capability, publishing no credentials', async () => {
  const f = fixture()
  expect(f.connection.snapshot().phase).toBe('signed_out')
  await f.connection.signIn()
  expect(f.loginDialog).toHaveBeenCalledWith(
    'https://auth.example/authorize',
    expect.any(AbortSignal),
  )
  expect(f.connection.available()).toBe(true)
  expect(f.session.connect).toHaveBeenCalledWith('powerpoint')
  expect(JSON.stringify(f.connection.snapshot())).not.toContain('secret')
  expect(f.connection.snapshot().account?.userId).toBe('account-a')
  await f.connection.request({ operation: 'team_identity' })
  expect(f.session.capabilityFetch).toHaveBeenCalledWith(
    'presentation-team.v1',
    { operation: 'team_identity' },
    undefined,
  )
})
it('signout clears credentials, disconnects and invalidates shared content', async () => {
  const f = fixture()
  await f.connection.signIn()
  await f.connection.signOut()
  expect(f.connection.snapshot().phase).toBe('signed_out')
  expect(f.connection.available()).toBe(false)
  expect(f.onUnavailable).toHaveBeenCalled()
  expect(await f.provider()).toBeNull()
})
it('relay expiry clears team availability and notifies the workbench', async () => {
  const f = fixture()
  await f.connection.signIn()
  f.emit({ status: 'expired' })
  expect(f.connection.available()).toBe(false)
  expect(f.connection.snapshot().phase).toBe('error')
  expect(f.onUnavailable).toHaveBeenCalled()
})
it('account changes during token acquisition refuse forwarding and clear the old account', async () => {
  const f = fixture()
  await f.connection.signIn()
  f.setAccount()
  expect(await f.provider()).toBeNull()
  expect(f.connection.available()).toBe(false)
  expect(f.onUnavailable).toHaveBeenCalled()
})
it('rejects a mixed negotiated capability set', async () => {
  const f = fixture()
  await f.connection.signIn()
  f.emit({ status: 'connected', capabilities: ['presentation-team.v1', 'presentation.v1'] })
  expect(f.connection.available()).toBe(false)
  expect(f.onUnavailable).toHaveBeenCalled()
})
it('signout during a delayed login dialog prevents late callback consumption or pairing', async () => {
  const f = fixture()
  let finish!: (value: string) => void
  f.loginDialog.mockImplementationOnce(
    () =>
      new Promise((resolve) => {
        finish = resolve
      }),
  )
  const login = f.connection.signIn()
  await vi.waitFor(() => expect(f.loginDialog).toHaveBeenCalled())
  await f.connection.signOut()
  finish('https://addin.example/callback?code=late')
  await login
  expect(f.auth.consumeCallback).not.toHaveBeenCalled()
  expect(f.session.connect).not.toHaveBeenCalled()
  expect(f.connection.snapshot().phase).toBe('signed_out')
})
it('safe dialog failure never exposes raw error strings or endpoints', async () => {
  const f = fixture()
  f.loginDialog.mockRejectedValueOnce(Error('secret-access https://private.example'))
  await f.connection.signIn()
  expect(f.connection.snapshot().phase).toBe('error')
  expect(JSON.stringify(f.connection.snapshot())).not.toContain('secret-access')
  expect(JSON.stringify(f.connection.snapshot())).not.toContain('private.example')
})
it('dispose prevents late session events from resurrecting availability', async () => {
  const f = fixture()
  await f.connection.signIn()
  f.connection.dispose()
  f.emit({ status: 'connected', capabilities: ['presentation-team.v1'] })
  expect(f.connection.available()).toBe(false)
  expect(f.connection.snapshot().account).toBeUndefined()
})
function browserAuth(fetch: typeof globalThis.fetch, now?: () => number) {
  const store = createMemorySessionStore()
  return {
    store,
    auth: createBrowserAuthClient({
      store,
      fetch,
      now,
      config: {
        clientId: 'registered-test-client',
        authorizationEndpoint: 'https://auth.example/authorize',
        callbackEndpoint: 'https://gateway.example/callback',
        refreshEndpoint: 'https://gateway.example/refresh',
        redirectUri: 'https://addin.example/callback',
      },
    }),
  }
}
const browserDialog = async (url: string) => {
  const state = new URL(url).searchParams.get('state')!
  return `https://addin.example/callback?code=exact-code&state=${encodeURIComponent(state)}`
}
it('connects with the real browser OAuth client and switches accounts without exposing tokens', async () => {
  const f = fixture()
  let attempt = 0
  const fetch = vi.fn(async () =>
    Response.json({
      token: `access-${++attempt}`,
      refresh_token: `refresh-${attempt}`,
      user_id: `gateway-user-${attempt}`,
      expires_in: 3600,
    }),
  )
  const a = browserAuth(fetch as typeof globalThis.fetch)
  const connection = createOfficeTeamConnection({
    auth: a.auth,
    loginDialog: browserDialog,
    createSession: f.createSession,
    onUnavailable: f.onUnavailable,
  })
  await connection.signIn()
  expect(connection.snapshot().account?.userId).toBe('gateway-user-1')
  expect(await f.provider()).toBe('access-1')
  await connection.signOut()
  expect(await a.store.load()).toBeNull()
  await connection.signIn()
  expect(connection.snapshot().account?.userId).toBe('gateway-user-2')
  expect(await f.provider()).toBe('access-2')
  expect(JSON.stringify(connection.snapshot())).not.toContain('access-')
  expect(f.onUnavailable).toHaveBeenCalled()
  connection.dispose()
})
it('signout while a real callback exchange is delayed cannot resurrect the memory session', async () => {
  const f = fixture()
  let resolve!: (response: Response) => void
  const fetch = vi.fn(
    () =>
      new Promise<Response>((r) => {
        resolve = r
      }),
  )
  const a = browserAuth(fetch as typeof globalThis.fetch)
  const connection = createOfficeTeamConnection({
    auth: a.auth,
    loginDialog: browserDialog,
    createSession: f.createSession,
    onUnavailable: f.onUnavailable,
  })
  const login = connection.signIn()
  await vi.waitFor(() => expect(fetch).toHaveBeenCalledOnce())
  await connection.signOut()
  resolve(
    Response.json({
      token: 'late-access',
      refresh_token: 'late-refresh',
      user_id: 'late-user',
      expires_in: 3600,
    }),
  )
  await login
  expect(await a.store.load()).toBeNull()
  expect(connection.snapshot().phase).toBe('signed_out')
  expect(f.session.connect).not.toHaveBeenCalled()
  connection.dispose()
})
it('signout while real refresh is delayed prevents late forwarding and stored tokens', async () => {
  const f = fixture()
  let now = 1000,
    resolve!: (response: Response) => void
  const fetch = vi.fn(async (url: RequestInfo | URL) =>
    String(url).includes('/refresh')
      ? await new Promise<Response>((r) => {
          resolve = r
        })
      : Response.json({
          token: 'initial',
          refresh_token: 'refresh',
          user_id: 'user',
          expires_in: 3600,
        }),
  )
  const a = browserAuth(fetch as typeof globalThis.fetch, () => now)
  const connection = createOfficeTeamConnection({
    auth: a.auth,
    loginDialog: browserDialog,
    createSession: f.createSession,
    onUnavailable: f.onUnavailable,
  })
  await connection.signIn()
  now += 3600_000
  const request = connection.request({ operation: 'team_identity' })
  const rejected = expect(request).rejects.toThrow()
  await vi.waitFor(() => expect(resolve).toBeTypeOf('function'))
  await connection.signOut()
  resolve(
    Response.json({
      token: 'late-access',
      refresh_token: 'late-refresh',
      user_id: 'user',
      expires_in: 3600,
    }),
  )
  await rejected
  expect(await a.store.load()).toBeNull()
  expect(f.session.capabilityFetch).not.toHaveBeenCalled()
  expect(connection.available()).toBe(false)
  connection.dispose()
})
it('rejects tokens larger than the team transport budget and clears authentication', async () => {
  const f = fixture()
  await f.connection.signIn()
  vi.spyOn(f.auth, 'getAccessToken').mockResolvedValue('x'.repeat(4097))
  expect(await f.provider()).toBeNull()
  expect(f.connection.available()).toBe(false)
  expect(f.connection.snapshot().account).toBeUndefined()
  expect(f.session.capabilityFetch).not.toHaveBeenCalled()
})
it('publishes actual pairing code and cancellation prevents a delayed connection becoming ready', async () => {
  const f = fixture()
  let finish!: () => void
  vi.mocked(f.session.connect).mockImplementationOnce(
    () =>
      new Promise<void>((resolve) => {
        finish = resolve
      }),
  )
  const login = f.connection.signIn()
  await vi.waitFor(() => expect(f.session.connect).toHaveBeenCalled())
  f.emit({ status: 'pending', verificationCode: '654321', capabilities: ['presentation-team.v1'] })
  expect(f.connection.snapshot()).toMatchObject({ phase: 'pairing', verificationCode: '654321' })
  await f.connection.signOut()
  finish()
  await login
  expect(f.connection.snapshot().phase).toBe('signed_out')
  expect(f.connection.available()).toBe(false)
})
