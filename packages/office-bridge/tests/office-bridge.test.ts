import { describe, expect, it, vi } from 'vitest'

import {
  assertLoopbackHost,
  createOfficeBridge,
  type MessagesProxy,
  type OfficeBridge,
} from '../src/index'

const origin = 'https://office.example.test'
const json = (value: unknown) => JSON.stringify(value)

function request(path: string, init: RequestInit = {}) {
  return new Request(`http://127.0.0.1:43127${path}`, {
    ...init,
    headers: { origin, ...init.headers },
  })
}

async function createPairing(bridge: ReturnType<typeof createOfficeBridge>) {
  const response = await bridge.handle(
    request('/v1/office/pairings', {
      method: 'POST',
      headers: { origin, 'content-type': 'application/json' },
      body: json({ host_label: 'Word' }),
    }),
  )
  return (await response.json()) as {
    pairing_id: string
    polling_secret: string
    expires_in: number
  }
}

describe('loopback and browser boundary', () => {
  it('accepts only the numeric IPv4 loopback bind host', () => {
    expect(assertLoopbackHost('127.0.0.1')).toBe('127.0.0.1')
    expect(() => assertLoopbackHost('localhost')).toThrow('loopback_host_required')
    expect(() => assertLoopbackHost('0.0.0.0')).toThrow('loopback_host_required')
    expect(() =>
      createOfficeBridge({ allowedOrigin: 'http://office.test', proxy: vi.fn() }),
    ).toThrow('invalid_allowed_origin')
  })

  it('rejects non-exact origins and returns an exact PNA preflight', async () => {
    const bridge = createOfficeBridge({ allowedOrigin: origin, proxy: vi.fn() })
    const denied = await bridge.handle(
      new Request('http://127.0.0.1:43127/v1/office/pairings', {
        method: 'POST',
        headers: { origin: 'https://evil.example', 'content-type': 'application/json' },
        body: '{}',
      }),
    )
    expect(denied.status).toBe(403)
    expect(denied.headers.get('access-control-allow-origin')).toBeNull()

    const preflight = await bridge.handle(
      request('/v1/office/messages', {
        method: 'OPTIONS',
        headers: {
          origin,
          'access-control-request-method': 'POST',
          'access-control-request-headers': 'authorization, content-type',
          'access-control-request-private-network': 'true',
        },
      }),
    )
    expect(preflight.status).toBe(204)
    expect(preflight.headers.get('access-control-allow-origin')).toBe(origin)
    expect(preflight.headers.get('access-control-allow-private-network')).toBe('true')
    expect(preflight.headers.get('vary')).toContain('Origin')
    expect(preflight.headers.get('access-control-allow-origin')).not.toBe('*')
  })
})

describe('pairing lifecycle', () => {
  it('creates independent random identifiers and reports pending', async () => {
    const bridge = createOfficeBridge({ allowedOrigin: origin, proxy: vi.fn() })
    const first = await createPairing(bridge)
    const second = await createPairing(bridge)
    expect(first.pairing_id).not.toBe(second.pairing_id)
    expect(first.polling_secret).not.toBe(second.polling_secret)
    expect(first.pairing_id.length).toBeGreaterThanOrEqual(22)
    expect(bridge.listPending()).toEqual([
      { pairingId: first.pairing_id, hostLabel: 'Word', origin },
      { pairingId: second.pairing_id, hostLabel: 'Word', origin },
    ])
    const poll = await bridge.handle(
      request(`/v1/office/pairings/${first.pairing_id}`, {
        headers: { origin, authorization: `Pairing ${first.polling_secret}` },
      }),
    )
    expect(await poll.json()).toEqual({ status: 'pending' })
  })

  it('approves only with a signed-in session and redeems a capability once', async () => {
    const bridge = createOfficeBridge({ allowedOrigin: origin, proxy: vi.fn() })
    const pairing = await createPairing(bridge)
    expect(bridge.approve(pairing.pairing_id, false)).toBe(false)
    expect(bridge.approve(pairing.pairing_id, true)).toBe(true)
    const first = await bridge.handle(
      request(`/v1/office/pairings/${pairing.pairing_id}`, {
        headers: { origin, authorization: `Pairing ${pairing.polling_secret}` },
      }),
    )
    const body = (await first.json()) as Record<string, unknown>
    expect(body.status).toBe('approved')
    expect(typeof body.capability).toBe('string')
    expect(JSON.stringify(body)).not.toMatch(/access_token|refresh_token/i)
    const replay = await bridge.handle(
      request(`/v1/office/pairings/${pairing.pairing_id}`, {
        headers: { origin, authorization: `Pairing ${pairing.polling_secret}` },
      }),
    )
    expect(replay.status).toBe(401)
    expect(await replay.json()).toEqual({ error: 'invalid_pairing' })
  })

  it('reports rejected and expired pairings without issuing grants', async () => {
    let now = 1_000
    const bridge = createOfficeBridge({
      allowedOrigin: origin,
      proxy: vi.fn(),
      now: () => now,
      pairingTtlMs: 100,
    })
    const rejected = await createPairing(bridge)
    bridge.reject(rejected.pairing_id)
    const rejectedPoll = await bridge.handle(
      request(`/v1/office/pairings/${rejected.pairing_id}`, {
        headers: { origin, authorization: `Pairing ${rejected.polling_secret}` },
      }),
    )
    expect(await rejectedPoll.json()).toEqual({ status: 'rejected' })

    const expired = await createPairing(bridge)
    now += 101
    const expiredPoll = await bridge.handle(
      request(`/v1/office/pairings/${expired.pairing_id}`, {
        headers: { origin, authorization: `Pairing ${expired.polling_secret}` },
      }),
    )
    expect(await expiredPoll.json()).toEqual({ status: 'expired' })
  })

  it('fails closed for invalid secrets and bounded pairing capacity', async () => {
    const bridge = createOfficeBridge({
      allowedOrigin: origin,
      proxy: vi.fn(),
      maxPairings: 1,
    })
    const pairing = await createPairing(bridge)
    const invalid = await bridge.handle(
      request(`/v1/office/pairings/${pairing.pairing_id}`, {
        headers: { origin, authorization: 'Pairing wrong' },
      }),
    )
    expect(invalid.status).toBe(401)
    const full = await bridge.handle(
      request('/v1/office/pairings', {
        method: 'POST',
        headers: { origin, 'content-type': 'application/json' },
        body: json({ host_label: 'Excel' }),
      }),
    )
    expect(full.status).toBe(429)
    expect(await full.json()).toEqual({ error: 'pairing_capacity' })
  })

  it('rate-limits pairing creation and rejects unsupported methods', async () => {
    const bridge = createOfficeBridge({
      allowedOrigin: origin,
      proxy: vi.fn(),
      maxPairings: 5,
      maxPairingCreatesPerMinute: 1,
    })
    await createPairing(bridge)
    const limited = await bridge.handle(
      request('/v1/office/pairings', {
        method: 'POST',
        headers: { origin, 'content-type': 'application/json' },
        body: json({ host_label: 'Excel' }),
      }),
    )
    expect(limited.status).toBe(429)
    const method = await bridge.handle(request('/v1/office/messages'))
    expect(method.status).toBe(405)
    expect(await method.json()).toEqual({ error: 'method_not_allowed' })
  })

  it('bounds live capabilities and frees capacity after capability expiry', async () => {
    let now = 1_000
    const bridge = createOfficeBridge({
      allowedOrigin: origin,
      proxy: vi.fn(),
      now: () => now,
      maxCapabilities: 1,
      capabilityTtlMs: 100,
    })
    const approveAndPoll = async () => {
      const pairing = await createPairing(bridge)
      bridge.approve(pairing.pairing_id, true)
      return bridge.handle(
        request(`/v1/office/pairings/${pairing.pairing_id}`, {
          headers: { origin, authorization: `Pairing ${pairing.polling_secret}` },
        }),
      )
    }
    expect((await approveAndPoll()).status).toBe(200)
    const full = await approveAndPoll()
    expect(full.status).toBe(429)
    expect(await full.json()).toEqual({ error: 'capability_capacity' })
    now += 101
    expect((await approveAndPoll()).status).toBe(200)
  })
})

describe('messages capability', () => {
  async function approvedBridge(proxy: MessagesProxy, options = {}) {
    const bridge = createOfficeBridge({ allowedOrigin: origin, proxy, ...options })
    const pairing = await createPairing(bridge)
    bridge.approve(pairing.pairing_id, true)
    const poll = await bridge.handle(
      request(`/v1/office/pairings/${pairing.pairing_id}`, {
        headers: { origin, authorization: `Pairing ${pairing.polling_secret}` },
      }),
    )
    const { capability } = (await poll.json()) as { capability: string }
    return { bridge, capability }
  }

  it('proxies bounded JSON using an opaque capability and streams safe output', async () => {
    const proxy = vi.fn<MessagesProxy>().mockResolvedValue({
      status: 200,
      contentType: 'text/event-stream',
      body: (async function* () {
        yield new TextEncoder().encode('data: ok\n\n')
      })(),
    })
    const { bridge, capability } = await approvedBridge(proxy)
    const response = await bridge.handle(
      request('/v1/office/messages', {
        method: 'POST',
        headers: {
          origin,
          authorization: `Bridge ${capability}`,
          'content-type': 'application/json',
        },
        body: json({ messages: [{ role: 'user', content: 'hi' }] }),
      }),
    )
    expect(response.status).toBe(200)
    expect(await response.text()).toBe('data: ok\n\n')
    expect(proxy).toHaveBeenCalledWith(
      expect.objectContaining({ body: { messages: [{ role: 'user', content: 'hi' }] } }),
    )
    expect(JSON.stringify(proxy.mock.calls)).not.toMatch(/access_token|refresh_token/i)
  })

  it('enforces content type, body, concurrency, stream, and timeout limits', async () => {
    let release!: () => void
    let markStarted!: () => void
    const waiting = new Promise<void>((resolve) => (release = resolve))
    const started = new Promise<void>((resolve) => (markStarted = resolve))
    const proxy: MessagesProxy = async () => {
      markStarted()
      await waiting
      return { status: 200, body: new Uint8Array([1]) }
    }
    const { bridge, capability } = await approvedBridge(proxy, {
      maxBodyBytes: 8,
      maxConcurrentMessages: 1,
      messageTimeoutMs: 1_000,
      maxResponseBytes: 4,
    })
    const headers = {
      origin,
      authorization: `Bridge ${capability}`,
      'content-type': 'application/json',
    }
    const first = bridge.handle(
      request('/v1/office/messages', { method: 'POST', headers, body: '{}' }),
    )
    await started
    const concurrent = await bridge.handle(
      request('/v1/office/messages', { method: 'POST', headers, body: '{}' }),
    )
    expect(concurrent.status).toBe(429)
    release()
    await first

    const oversized = await bridge.handle(
      request('/v1/office/messages', { method: 'POST', headers, body: '{"long":1}' }),
    )
    expect(oversized.status).toBe(413)
    const wrongType = await bridge.handle(
      request('/v1/office/messages', {
        method: 'POST',
        headers: { origin, authorization: `Bridge ${capability}` },
        body: '{}',
      }),
    )
    expect(wrongType.status).toBe(415)
  })

  it('expires and revokes capabilities on logout or shutdown', async () => {
    let now = 1_000
    const proxy = vi.fn<MessagesProxy>().mockResolvedValue({ status: 200, body: new Uint8Array() })
    const { bridge, capability } = await approvedBridge(proxy, {
      now: () => now,
      capabilityTtlMs: 100,
    })
    const call = () =>
      bridge.handle(
        request('/v1/office/messages', {
          method: 'POST',
          headers: {
            origin,
            authorization: `Bridge ${capability}`,
            'content-type': 'application/json',
          },
          body: '{}',
        }),
      )
    now += 101
    expect((await call()).status).toBe(401)
    bridge.revokeAll()
    expect((await call()).status).toBe(401)
    bridge.shutdown()
    expect((await call()).status).toBe(503)
  })

  it('maps proxy failures to stable errors without leaking upstream bodies', async () => {
    const upstreamSecret = 'upstream says token=secret'
    const { bridge, capability } = await approvedBridge(async () => {
      throw new Error(upstreamSecret)
    })
    const response = await bridge.handle(
      request('/v1/office/messages', {
        method: 'POST',
        headers: {
          origin,
          authorization: `Bridge ${capability}`,
          'content-type': 'application/json',
        },
        body: '{}',
      }),
    )
    expect(response.status).toBe(502)
    const body = await response.text()
    expect(body).toBe('{"error":"upstream_failed"}')
    expect(body).not.toContain(upstreamSecret)
  })

  it('returns only a stable error when a response exceeds its byte budget or times out', async () => {
    const oversized = await approvedBridge(async () => ({ status: 200, body: new Uint8Array(5) }), {
      maxResponseBytes: 4,
    })
    const invoke = (bridge: OfficeBridge, capability: string) =>
      bridge.handle(
        request('/v1/office/messages', {
          method: 'POST',
          headers: {
            origin,
            authorization: `Bridge ${capability}`,
            'content-type': 'application/json',
          },
          body: '{}',
        }),
      )
    expect((await invoke(oversized.bridge, oversized.capability)).status).toBe(502)

    const timedOut = await approvedBridge(() => new Promise(() => undefined), {
      messageTimeoutMs: 5,
    })
    const response = await invoke(timedOut.bridge, timedOut.capability)
    expect(response.status).toBe(502)
    expect(await response.json()).toEqual({ error: 'upstream_failed' })
  })
})
