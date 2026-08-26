import { describe, expect, it, vi } from 'vitest'

import { WISWORK_MESSAGES_URL } from '@wiswork/ai-provider'
import { createOfficeBridge } from '@wiswork/office-bridge'
import {
  createOfficeMessagesProxy,
  logoutAndRevokeOfficeBridge,
  officeBridgePortFromEnv,
  runOfficeBridgeLifecycle,
  validAccountStatusOrRevoke,
  syncOfficeBridgeAvailability,
  officeBridgeEnabled,
  officeBridgePortsFromEnv,
  bindOfficeBridgePortPool,
  officeBridgeDiagnosticForError,
} from '../src/main/office-bridge-runtime'

describe('Office bridge shell runtime', () => {
  it('parses the strict port pool and falls back only after address-in-use', async () => {
    const defaults = officeBridgePortsFromEnv({})
    expect(defaults).toHaveLength(64)
    expect(defaults[0]).toBe(43127)
    expect(() => officeBridgePortsFromEnv({ WISWORK_OFFICE_BRIDGE_PORTS: '43127,43127' })).toThrow()
    expect(() => officeBridgePortsFromEnv({ WISWORK_OFFICE_BRIDGE_PORTS: '43127,' })).toThrow()
    expect(officeBridgePortsFromEnv({ WISWORK_OFFICE_BRIDGE_PORT: '44000' })).toEqual([44000])
    expect(
      officeBridgePortsFromEnv({
        WISWORK_OFFICE_BRIDGE_PORT: '44000',
        WISWORK_OFFICE_BRIDGE_PORTS: '45000,45001',
      }),
    ).toEqual([45000, 45001])
    expect(() => officeBridgePortsFromEnv({ WISWORK_OFFICE_BRIDGE_PORT: 'invalid' })).toThrow(
      'invalid_office_bridge_port',
    )
    const start = vi.fn(async (port: number) => {
      if (port === 43127) throw Object.assign(new Error('busy'), { code: 'EADDRINUSE' })
      return 'server'
    })
    await expect(bindOfficeBridgePortPool([43127, 43128], start)).resolves.toEqual({
      port: 43128,
      server: 'server',
    })
    await expect(
      bindOfficeBridgePortPool([1, 2], async () => {
        throw Object.assign(new Error('denied'), { code: 'EACCES' })
      }),
    ).rejects.toThrow('denied')
  })
  it('distinguishes exhausted pools from configuration and bind failures', () => {
    expect(
      officeBridgeDiagnosticForError(
        Object.assign(new Error('office_bridge_pool_exhausted'), { code: 'EADDRINUSE' }),
      ),
    ).toBe('error:pool_exhausted')
    expect(officeBridgeDiagnosticForError(new Error('invalid_office_bridge_ports'))).toBe(
      'error:invalid_config',
    )
    let legacyConfigError: unknown
    try {
      officeBridgePortsFromEnv({ WISWORK_OFFICE_BRIDGE_PORT: 'invalid' })
    } catch (error) {
      legacyConfigError = error
    }
    expect(officeBridgeDiagnosticForError(legacyConfigError)).toBe('error:invalid_config')
    expect(
      officeBridgeDiagnosticForError(Object.assign(new Error('denied'), { code: 'EACCES' })),
    ).toBe('error:bind_failed')
  })
  it('keeps the loopback rollback bridge explicitly opt-in in every build', () => {
    expect(officeBridgeEnabled({}, false)).toBe(false)
    expect(officeBridgeEnabled({}, true)).toBe(false)
    expect(officeBridgeEnabled({ WISWORK_OFFICE_BRIDGE_ENABLED: '0' }, true)).toBe(false)
    expect(officeBridgeEnabled({ WISWORK_OFFICE_BRIDGE_ENABLED: '1' }, false)).toBe(true)
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

  it('reenables pairing after a real signed-in account status without restarting', async () => {
    const bridge = createOfficeBridge({
      allowedOrigin: 'https://office.example.test',
      proxy: vi.fn(),
      sessionAvailable: false,
    })
    await expect(
      syncOfficeBridgeAvailability(bridge, async () => ({ loggedIn: true })),
    ).resolves.toEqual({ loggedIn: true })
    const response = await bridge.handle(
      new Request('http://127.0.0.1:43127/v1/office/pairings', {
        method: 'POST',
        headers: {
          origin: 'https://office.example.test',
          'content-type': 'application/json',
        },
        body: JSON.stringify({ host_label: 'Word' }),
      }),
    )
    expect(response.status).toBe(202)
  })
})
