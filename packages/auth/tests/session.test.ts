import { describe, expect, it, vi } from 'vitest'
import { createAuthClient, type AuthSession, type SessionStore } from '../src/index'

function deferred() {
  let resolve!: () => void
  const promise = new Promise<void>((done) => {
    resolve = done
  })
  return { promise, resolve }
}

function deferredSaveStore(initial: AuthSession | null) {
  let value = initial
  const saveEntered = deferred()
  const allowSave = deferred()
  const store: SessionStore = {
    load: vi.fn(async () => value),
    save: vi.fn(async (next) => {
      saveEntered.resolve()
      await allowSave.promise
      value = next
    }),
    clear: vi.fn(async () => {
      value = null
    }),
  }
  return { store, saveEntered: saveEntered.promise, allowSave: allowSave.resolve }
}

function twoCallbackHarness() {
  let value: AuthSession | null = null
  const store: SessionStore = {
    load: vi.fn(async () => value),
    save: vi.fn(async (next) => {
      value = next
    }),
    clear: vi.fn(async () => {
      value = null
    }),
  }
  const firstGate = deferred()
  const secondGate = deferred()
  const fetch = vi.fn(async (input: RequestInfo | URL) => {
    const code = new URL(String(input)).searchParams.get('code')
    const gate = code === 'first' ? firstGate : secondGate
    await gate.promise
    return new Response(
      JSON.stringify({
        token: code + '-access',
        refresh_token: code + '-refresh',
        user_id: code + '-user',
      }),
      { status: 200 },
    )
  })
  const client = createAuthClient({ store, fetch })
  const firstRequest = client.createAuthorizationRequest()
  const first = client.consumeCallback(
    'wiswork://oauth/callback?code=first&state=' + firstRequest.state,
  )
  const startSecond = async () => {
    await vi.waitFor(() => expect(fetch).toHaveBeenCalledTimes(1))
    const secondRequest = client.createAuthorizationRequest()
    const second = client.consumeCallback(
      'wiswork://oauth/callback?code=second&state=' + secondRequest.state,
    )
    await vi.waitFor(() => expect(fetch).toHaveBeenCalledTimes(2))
    return { second }
  }
  return { client, first, startSecond, firstGate, secondGate }
}

function setup() {
  const original: AuthSession = {
    accessToken: 'old-secret',
    refreshToken: 'refresh-secret',
    userId: 'u',
    expiresAt: 1,
  }
  let value: AuthSession | null = original
  const store: SessionStore = {
    load: vi.fn(async () => value),
    save: vi.fn(async (next) => void (value = next)),
    clear: vi.fn(async () => void (value = null)),
  }
  let releases!: () => void
  const gate = new Promise<void>((resolve) => (releases = resolve))
  const fetch = vi.fn(async () => {
    await gate
    return new Response(
      JSON.stringify({
        token: 'new-secret',
        refresh_token: 'new-refresh',
        user_id: 'u',
        expires_in: 3600,
      }),
      { status: 200 },
    )
  })
  const client = createAuthClient({ store, fetch, now: () => 10_000 })
  return { client, fetch, releases }
}

describe('session lifecycle', () => {
  it('reports logged out when no session exists without calling refresh', async () => {
    const store: SessionStore = {
      load: vi.fn(async () => null),
      save: vi.fn(),
      clear: vi.fn(),
    }
    const fetch = vi.fn()
    const client = createAuthClient({ store, fetch })

    await expect(client.getValidAccountStatus()).resolves.toEqual({ loggedIn: false })
    expect(fetch).not.toHaveBeenCalled()
  })

  it('returns a safe valid status without refreshing an unexpired session', async () => {
    const store: SessionStore = {
      load: vi.fn(async () => ({
        accessToken: 'unexposed-access',
        refreshToken: 'unexposed-refresh',
        userId: 'valid-user',
        email: 'valid@example.com',
        expiresAt: 200_000,
      })),
      save: vi.fn(),
      clear: vi.fn(),
    }
    const fetch = vi.fn()
    const client = createAuthClient({ store, fetch, now: () => 10_000 })

    const status = await client.getValidAccountStatus()

    expect(status).toEqual({ loggedIn: true, userId: 'valid-user', email: 'valid@example.com' })
    expect(JSON.stringify(status)).not.toContain('unexposed')
    expect(fetch).not.toHaveBeenCalled()
  })

  it('refreshes an expired session before reporting it as logged in', async () => {
    const { client, fetch, releases } = setup()
    const first = client.getValidAccountStatus()
    const second = client.getValidAccountStatus()
    await vi.waitFor(() => expect(fetch).toHaveBeenCalledTimes(1))
    releases()

    await expect(first).resolves.toEqual({ loggedIn: true, userId: 'u' })
    await expect(second).resolves.toEqual({ loggedIn: true, userId: 'u' })
    expect(fetch).toHaveBeenCalledTimes(1)
  })

  it.each([400, 401])(
    'clears invalid expired credentials and reports logged out after refresh %i',
    async (status) => {
      let value: AuthSession | null = {
        accessToken: 'expired-access',
        refreshToken: 'invalid-refresh',
        userId: 'expired-user',
        expiresAt: 1,
      }
      const store: SessionStore = {
        load: vi.fn(async () => value),
        save: vi.fn(async (next) => void (value = next)),
        clear: vi.fn(async () => void (value = null)),
      }
      const fetch = vi.fn(async () => new Response('', { status }))
      const client = createAuthClient({ store, fetch, now: () => 10_000 })

      await expect(client.getValidAccountStatus()).resolves.toEqual({ loggedIn: false })
      expect(await client.getAccountStatus()).toEqual({ loggedIn: false })
      expect(store.clear).toHaveBeenCalledTimes(1)
    },
  )

  it('does not report an expired session as logged in when refresh has a transient failure', async () => {
    const store: SessionStore = {
      load: vi.fn(async () => ({
        accessToken: 'expired-access',
        refreshToken: 'preserved-refresh',
        userId: 'expired-user',
        expiresAt: 1,
      })),
      save: vi.fn(),
      clear: vi.fn(),
    }
    const client = createAuthClient({
      store,
      fetch: vi.fn(async () => new Response('', { status: 503 })),
      now: () => 10_000,
    })

    await expect(client.getValidAccountStatus()).rejects.toMatchObject({ code: 'network_error' })
    expect(store.clear).not.toHaveBeenCalled()
  })

  it('coalesces concurrent refresh calls', async () => {
    const { client, fetch, releases } = setup()
    const first = client.refresh()
    const second = client.refresh()
    await vi.waitFor(() => expect(fetch).toHaveBeenCalledTimes(1))
    releases()
    expect(await first).toEqual(await second)
    expect(fetch).toHaveBeenCalledTimes(1)
  })

  it('does not restore a refreshed session after logout', async () => {
    const { client, fetch, releases } = setup()
    const refresh = client.refresh()
    await vi.waitFor(() => expect(fetch).toHaveBeenCalledTimes(1))
    await client.logout()
    releases()
    await expect(refresh).rejects.toMatchObject({ code: 'auth_required' })
    expect(await client.getAccountStatus()).toEqual({ loggedIn: false })
    expect(fetch).toHaveBeenCalledTimes(1)
  })

  it('serializes logout after a refresh already entered deferred save', async () => {
    const initial: AuthSession = {
      accessToken: 'old',
      refreshToken: 'old-refresh',
      userId: 'old-user',
    }
    const controlled = deferredSaveStore(initial)
    const client = createAuthClient({
      store: controlled.store,
      fetch: vi.fn(
        async () =>
          new Response(
            JSON.stringify({
              token: 'refreshed',
              refresh_token: 'refreshed-token',
              user_id: 'old-user',
            }),
            { status: 200 },
          ),
      ),
    })
    const refresh = client.refresh()
    await controlled.saveEntered
    const logout = client.logout()
    controlled.allowSave()
    await refresh
    await logout
    expect(await client.getAccountStatus()).toEqual({ loggedIn: false })
  })

  it('serializes logout after a callback already entered deferred save', async () => {
    const controlled = deferredSaveStore(null)
    const client = createAuthClient({
      store: controlled.store,
      fetch: vi.fn(
        async () =>
          new Response(
            JSON.stringify({
              token: 'new-access',
              refresh_token: 'new-refresh',
              user_id: 'new-user',
            }),
            { status: 200 },
          ),
      ),
    })
    const request = client.createAuthorizationRequest()
    const callback = client.consumeCallback(
      `wiswork://oauth/callback?code=ok&state=${request.state}`,
    )
    await controlled.saveEntered
    const logout = client.logout()
    controlled.allowSave()
    await callback
    await logout
    expect(await client.getAccountStatus()).toEqual({ loggedIn: false })
  })

  it('prevents an old-user refresh from overwriting a new-user callback', async () => {
    let value: AuthSession | null = {
      accessToken: 'old',
      refreshToken: 'old-refresh',
      userId: 'old-user',
    }
    const store: SessionStore = {
      load: vi.fn(async () => value),
      save: vi.fn(async (next) => {
        value = next
      }),
      clear: vi.fn(async () => {
        value = null
      }),
    }
    const allowRefresh = deferred()
    const fetch = vi.fn(async (input: RequestInfo | URL) => {
      if (String(input).includes('/refresh')) {
        await allowRefresh.promise
        return new Response(
          JSON.stringify({
            token: 'old-refreshed',
            refresh_token: 'old-refresh-2',
            user_id: 'old-user',
          }),
          { status: 200 },
        )
      }
      return new Response(
        JSON.stringify({ token: 'new-access', refresh_token: 'new-refresh', user_id: 'new-user' }),
        { status: 200 },
      )
    })
    const client = createAuthClient({ store, fetch })
    const refresh = client.refresh()
    await vi.waitFor(() => expect(fetch).toHaveBeenCalledTimes(1))
    const request = client.createAuthorizationRequest()
    await client.consumeCallback(`wiswork://oauth/callback?code=ok&state=${request.state}`)
    allowRefresh.resolve()
    await expect(refresh).rejects.toMatchObject({ code: 'auth_required' })
    expect(await client.getAccountStatus()).toEqual({ loggedIn: true, userId: 'new-user' })
  })

  it('keeps the latest login identity when the first callback exchange returns first', async () => {
    const { client, first, startSecond, firstGate, secondGate } = twoCallbackHarness()
    const { second } = await startSecond()
    firstGate.resolve()
    await expect(first).rejects.toMatchObject({ code: 'auth_required' })
    secondGate.resolve()
    await expect(second).resolves.toMatchObject({ userId: 'second-user' })
    expect(await client.getAccountStatus()).toEqual({ loggedIn: true, userId: 'second-user' })
  })

  it('keeps the latest login identity when the second callback exchange returns first', async () => {
    const { client, first, startSecond, firstGate, secondGate } = twoCallbackHarness()
    const { second } = await startSecond()
    secondGate.resolve()
    await expect(second).resolves.toMatchObject({ userId: 'second-user' })
    firstGate.resolve()
    await expect(first).rejects.toMatchObject({ code: 'auth_required' })
    expect(await client.getAccountStatus()).toEqual({ loggedIn: true, userId: 'second-user' })
  })

  it('invalidates both callback exchanges when logout happens during them', async () => {
    const { client, first, startSecond, firstGate, secondGate } = twoCallbackHarness()
    const { second } = await startSecond()
    await client.logout()
    firstGate.resolve()
    secondGate.resolve()
    await expect(first).rejects.toMatchObject({ code: 'auth_required' })
    await expect(second).rejects.toMatchObject({ code: 'auth_required' })
    expect(await client.getAccountStatus()).toEqual({ loggedIn: false })
  })

  it('recovers the mutation queue after a failed commit', async () => {
    const store: SessionStore = {
      load: vi.fn(async () => null),
      save: vi.fn(async () => {
        throw new Error('disk failure')
      }),
      clear: vi.fn(async () => undefined),
    }
    const client = createAuthClient({
      store,
      fetch: vi.fn(
        async () =>
          new Response(
            JSON.stringify({ token: 'access', refresh_token: 'refresh', user_id: 'user' }),
            { status: 200 },
          ),
      ),
    })
    const request = client.createAuthorizationRequest()
    await expect(
      client.consumeCallback(`wiswork://oauth/callback?code=ok&state=${request.state}`),
    ).rejects.toThrow('disk failure')
    await expect(client.logout()).resolves.toBeUndefined()
    expect(store.clear).toHaveBeenCalledTimes(1)
  })

  it('retries an authenticated request after 401 at most once', async () => {
    const { client, releases } = setup()
    releases()
    const request = vi.fn(async () => new Response('', { status: 401 }))
    const response = await client.fetchWithAuth(request)
    expect(response.status).toBe(401)
    expect(request).toHaveBeenCalledTimes(2)
  })

  it('never exposes tokens in account status', async () => {
    const { client, releases } = setup()
    releases()
    const status = await client.getAccountStatus()
    expect(status).toEqual({ loggedIn: true, userId: 'u' })
    expect(JSON.stringify(status)).not.toContain('secret')
  })
})
