import { describe, expect, it, vi } from 'vitest'

import { WISWORK_MESSAGES_URL } from '@wiswork/ai-provider'
import {
  createOfficeMessagesProxy,
  logoutAndRevokeOfficeBridge,
  officeBridgePortFromEnv,
  runOfficeBridgeLifecycle,
  validAccountStatusOrRevoke,
  officeBridgeEnabled,
} from '../src/main/office-bridge-runtime'

describe('Office bridge shell runtime', () => {
  it('is fail-closed unless explicitly enabled', () => {
    expect(officeBridgeEnabled({})).toBe(false)
    expect(officeBridgeEnabled({ WISWORK_OFFICE_BRIDGE_ENABLED: '0' })).toBe(false)
    expect(officeBridgeEnabled({ WISWORK_OFFICE_BRIDGE_ENABLED: '1' })).toBe(true)
  })
  it('uses a validated fixed configurable port', () => {
    expect(officeBridgePortFromEnv({})).toBe(43127)
    expect(officeBridgePortFromEnv({ WISWORK_OFFICE_BRIDGE_PORT: '44000' })).toBe(44000)
    for (const value of ['0', '65536', '1.2', 'nope']) {
      expect(() => officeBridgePortFromEnv({ WISWORK_OFFICE_BRIDGE_PORT: value })).toThrow(
        'invalid_office_bridge_port',
      )
    }
  })

  it('proxies only to the fixed messages destination through fetchWithAuth', async () => {
    const fetchWithAuth = vi.fn(async (request: (token: string) => Promise<Response>) =>
      request('private-wispaper-token'),
    )
    const upstreamFetch = vi.fn(async (_input: string | URL | Request, init?: RequestInit) => {
      expect(String(_input)).toBe(WISWORK_MESSAGES_URL)
      expect(new Headers(init?.headers).get('authorization')).toBe('Bearer private-wispaper-token')
      return new Response('data: ok\n\n', {
        status: 200,
        headers: { 'content-type': 'text/event-stream' },
      })
    })
    const proxy = createOfficeMessagesProxy({ fetchWithAuth, fetch: upstreamFetch })
    const response = await proxy({ body: { messages: [] }, signal: new AbortController().signal })
    expect(response.status).toBe(200)
    expect(fetchWithAuth).toHaveBeenCalledOnce()
    expect(JSON.stringify(response)).not.toContain('private-wispaper-token')
  })

  it('revokes all Office grants on terminal upstream authentication loss', async () => {
    const revoke = vi.fn()
    const proxy = createOfficeMessagesProxy({
      fetchWithAuth: vi.fn(async () => new Response('private upstream error', { status: 401 })),
      onTerminalAuthLoss: revoke,
    })
    await expect(proxy({ body: {}, signal: new AbortController().signal })).rejects.toThrow(
      'auth_required',
    )
    expect(revoke).toHaveBeenCalledOnce()
  })

  it('starts after ready and shuts down/revokes on quit', async () => {
    const events = new Map<string, () => void>()
    const app = {
      whenReady: vi.fn(async () => undefined),
      on: vi.fn((n: string, f: () => void) => events.set(n, f)),
    }
    const stop = vi.fn(async () => undefined)
    const bridge = { revokeAll: vi.fn(), shutdown: vi.fn() }
    const start = vi.fn(async () => ({ stop }))
    await runOfficeBridgeLifecycle({ app, bridge, start })
    expect(start).toHaveBeenCalledOnce()
    events.get('before-quit')?.()
    expect(bridge.revokeAll).toHaveBeenCalledOnce()
    expect(bridge.shutdown).toHaveBeenCalledOnce()
    expect(stop).toHaveBeenCalledOnce()
  })

  it('revokes every capability before clearing the PC session', async () => {
    const order: string[] = []
    await logoutAndRevokeOfficeBridge(
      { revokeAll: () => order.push('revoke') },
      async () => void order.push('logout'),
    )
    expect(order).toEqual(['revoke', 'logout'])
  })

  it('revokes capabilities when account validation reports terminal auth loss', async () => {
    const bridge = { revokeAll: vi.fn() }
    await expect(
      validAccountStatusOrRevoke(bridge, async () => ({ loggedIn: false })),
    ).resolves.toEqual({ loggedIn: false })
    expect(bridge.revokeAll).toHaveBeenCalledOnce()
  })
})
