import { describe, expect, it, vi } from 'vitest'
import {
  createOfficeRelayClient,
  officeRelayEndpointFromEnv,
  type RelaySocket,
} from '../src/main/office-relay-client'

class FakeSocket implements RelaySocket {
  readyState = 0
  sent: string[] = []
  listeners = new Map<string, Array<(event: any) => void>>()
  addEventListener(name: string, listener: (event: any) => void): void {
    this.listeners.set(name, [...(this.listeners.get(name) ?? []), listener])
  }
  send(data: string): void {
    this.sent.push(data)
  }
  close(): void {
    this.readyState = 3
    this.emit('close', {})
  }
  open(): void {
    this.readyState = 1
    this.emit('open', {})
  }
  message(value: object): void {
    this.emit('message', { data: JSON.stringify(value) })
  }
  emit(name: string, event: any): void {
    for (const listener of this.listeners.get(name) ?? []) listener(event)
  }
}

function setup(loggedIn = true) {
  const socket = new FakeSocket()
  const pending = vi.fn()
  const proxy = vi.fn(async () => ({
    status: 200,
    contentType: 'text/event-stream',
    body: (async function* () {
      yield new TextEncoder().encode('hello')
    })(),
  }))
  const client = createOfficeRelayClient({
    endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
    connect: () => socket,
    getValidAccountStatus: async () => ({ loggedIn }),
    getAccessToken: async () => (loggedIn ? 'access-token' : null),
    proxy,
    onPending: pending,
  })
  return { client, socket, pending, proxy }
}

describe('Office relay PC client', () => {
  it('accepts only secure relay endpoint configuration without embedded credentials', () => {
    expect(officeRelayEndpointFromEnv({})).toBe('wss://office.8-216-134-194.sslip.io/office-relay')
    expect(() =>
      officeRelayEndpointFromEnv({ WISWORK_OFFICE_RELAY_URL: 'ws://localhost/relay' }),
    ).toThrow('invalid_office_relay_url')
    expect(() =>
      officeRelayEndpointFromEnv({ WISWORK_OFFICE_RELAY_URL: 'wss://dev.example/office-relay' }),
    ).toThrow('invalid_office_relay_url')
    expect(() =>
      officeRelayEndpointFromEnv({
        WISWORK_OFFICE_RELAY_URL: 'wss://user:secret@example.com/relay',
      }),
    ).toThrow('invalid_office_relay_url')
  })

  it('requires sign-in and an exact six-digit code before connecting', async () => {
    await expect(setup(false).client.claim('123456')).rejects.toThrow('auth_required')
    await expect(setup().client.claim('12345')).rejects.toThrow('invalid_verification_code')
  })

  it('passes a freshly loaded access token to the authenticated socket factory on every claim', async () => {
    const sockets = [new FakeSocket(), new FakeSocket()]
    const connect = vi.fn((_url: string, _token: string) => sockets.shift()!)
    const getAccessToken = vi
      .fn()
      .mockResolvedValueOnce('token-one')
      .mockResolvedValueOnce('token-two')
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect,
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken,
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      onPending() {},
    })
    const first = client.claim('123456')
    await vi.waitFor(() => expect(connect).toHaveBeenCalledTimes(1))
    connect.mock.results[0]!.value.open()
    await first
    const second = client.claim('654321')
    await vi.waitFor(() => expect(connect).toHaveBeenCalledTimes(2))
    connect.mock.results[1]!.value.open()
    await second
    expect(connect).toHaveBeenNthCalledWith(
      1,
      'wss://office.8-216-134-194.sslip.io/office-relay',
      'token-one',
    )
    expect(connect).toHaveBeenNthCalledWith(
      2,
      'wss://office.8-216-134-194.sslip.io/office-relay',
      'token-two',
    )
  })

  it('claims, shows the relay-asserted host and code, and requires explicit approval', async () => {
    const { client, socket, pending } = setup()
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    expect(JSON.parse(socket.sent[0]!)).toEqual({
      version: 1,
      type: 'pc.claim',
      verification_code: '123456',
    })
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    expect(pending).toHaveBeenCalledWith({
      pairingId: 'pairing_12345678',
      hostLabel: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verificationCode: '123456',
    })
    await client.approve('pairing_12345678')
    expect(JSON.parse(socket.sent.at(-1)!)).toEqual({
      version: 1,
      type: 'pc.approve',
      pairing_id: 'pairing_12345678',
    })
  })

  it('fails closed if the relay echoes a different human verification code', async () => {
    const { client, socket, pending } = setup()
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '654321',
      expires_in: 120,
    })
    expect(pending).not.toHaveBeenCalled()
    expect(client.status()).toBe('disconnected:protocol_violation')
  })

  it('rejects approval before the exact pending pairing was locally approved', async () => {
    const { client, socket } = setup()
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
    })
    expect(client.status()).toBe('disconnected:protocol_violation')
  })

  it('rejects non-object agent request bodies without invoking the proxy', async () => {
    const { client, socket, proxy } = setup()
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 1,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
    })
    socket.message({
      version: 1,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
      body: '{"messages":[]}',
    })
    expect(proxy).not.toHaveBeenCalled()
    expect(client.status()).toBe('disconnected:protocol_violation')
  })

  it('bounds remembered request identifiers and rejects unknown relay error codes', async () => {
    const socket = new FakeSocket()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => {
        queueMicrotask(() => socket.open())
        return socket
      },
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken: async () => 'token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      onPending() {},
      maxRequestIds: 0,
    })
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 1,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
    })
    socket.message({
      version: 1,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
      body: {},
    })
    expect(client.status()).toBe('disconnected:protocol_violation')

    const second = setup()
    const secondClaim = second.client.claim('123456')
    await vi.waitFor(() => expect(second.socket.listeners.has('open')).toBe(true))
    second.socket.open()
    await secondClaim
    second.socket.message({ version: 1, type: 'relay.error', code: 'attacker_supplied_status' })
    expect(second.client.status()).toBe('disconnected:protocol_violation')
  })

  it('proxies bounded requests, streams base64 chunks, and mirrors completion metadata', async () => {
    const { client, socket, proxy } = setup()
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Excel',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 1,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
    })
    socket.message({
      version: 1,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
      body: { model: 'x' },
    })
    await vi.waitFor(() => expect(proxy).toHaveBeenCalled())
    await vi.waitFor(() =>
      expect(socket.sent.some((raw) => JSON.parse(raw).type === 'pc.done')).toBe(true),
    )
    const frames = socket.sent.map((raw) => JSON.parse(raw))
    expect(frames).toContainEqual({
      version: 1,
      type: 'pc.chunk',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      request_id: 'request_12345678',
      sequence: 0,
      data: 'aGVsbG8=',
    })
    expect(frames).toContainEqual({
      version: 1,
      type: 'pc.start',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      request_id: 'request_12345678',
      status: 200,
      content_type: 'text/event-stream',
    })
    expect(frames).toContainEqual({
      version: 1,
      type: 'pc.done',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      request_id: 'request_12345678',
    })
  })

  it.each([100, 199, 204, 205, 304])(
    'returns request_failed without ending the session for non-streaming status %i',
    async (status) => {
      const socket = new FakeSocket()
      const client = createOfficeRelayClient({
        endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
        connect: () => socket,
        getValidAccountStatus: async () => ({ loggedIn: true }),
        getAccessToken: async () => 'token',
        proxy: async () => ({ status, body: new Uint8Array() }),
        onPending() {},
      })
      const claiming = client.claim('123456')
      await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
      socket.open()
      await claiming
      socket.message({
        version: 1,
        type: 'pc.claimed',
        pairing_id: 'pairing_12345678',
        host: 'Word',
        origin: 'https://office.8-216-134-194.sslip.io',
        verification_code: '123456',
        expires_in: 120,
      })
      await client.approve('pairing_12345678')
      socket.message({
        version: 1,
        type: 'pc.approved',
        session_id: 'session_12345678',
        capability: 'secret-capability',
        expires_in: 1800,
      })
      socket.message({
        version: 1,
        type: 'relay.request',
        session_id: 'session_12345678',
        request_id: 'request_12345678',
        body: {},
      })
      await vi.waitFor(() =>
        expect(socket.sent.some((raw) => JSON.parse(raw).type === 'pc.error')).toBe(true),
      )
      const frames = socket.sent.map((raw) => JSON.parse(raw))
      expect(frames).not.toContainEqual(expect.objectContaining({ type: 'pc.start' }))
      expect(frames).toContainEqual(
        expect.objectContaining({ type: 'pc.error', code: 'request_failed' }),
      )
      expect(client.status()).toBe('paired')
    },
  )

  it('aborts active upstream work on relay.cancel', async () => {
    const socket = new FakeSocket()
    let aborted = false
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken: async () => 'access-token',
      proxy: ({ signal }) =>
        new Promise((_resolve, reject) => {
          signal.addEventListener('abort', () => {
            aborted = true
            reject(new Error('cancelled'))
          })
        }),
      onPending() {},
    })
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 1,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
    })
    socket.message({
      version: 1,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
      body: {},
    })
    await vi.waitFor(() => expect(client.status()).toBe('paired'))
    socket.message({
      version: 1,
      type: 'relay.cancel',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
    })
    await vi.waitFor(() => expect(aborted).toBe(true))
    expect(socket.sent.map((raw) => JSON.parse(raw).type)).not.toContain('pc.error')
    expect(client.status()).toBe('paired')
  })

  it('lets Relay own the 120s deadline and tolerates its late cancel after the 125s watchdog', async () => {
    vi.useFakeTimers()
    const socket = new FakeSocket()
    let aborted = false
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => {
        queueMicrotask(() => socket.open())
        return socket
      },
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken: async () => 'token',
      proxy: ({ signal }) =>
        new Promise((_resolve, reject) =>
          signal.addEventListener('abort', () => {
            aborted = true
            reject(new Error('aborted'))
          }),
        ),
      onPending() {},
    })
    await client.claim('123456')
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 1,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
    })
    socket.message({
      version: 1,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
      body: {},
    })
    await vi.advanceTimersByTimeAsync(120_000)
    expect(aborted).toBe(false)
    await vi.advanceTimersByTimeAsync(5_000)
    expect(aborted).toBe(true)
    socket.message({
      version: 1,
      type: 'relay.cancel',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
    })
    expect(client.status()).toBe('paired')
    vi.useRealTimers()
  })

  it('expires a pending pairing and removes its approval prompt', async () => {
    vi.useFakeTimers()
    const socket = new FakeSocket()
    const expired = vi.fn()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken: async () => 'token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      onPending() {},
      onPendingExpired: expired,
    })
    const claiming = client.claim('123456')
    await vi.advanceTimersByTimeAsync(0)
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 1,
    })
    await vi.advanceTimersByTimeAsync(1_000)
    expect(expired).toHaveBeenCalledWith('pairing_12345678')
    expect(client.listPending()).toEqual([])
    expect(client.status()).toBe('disconnected:pairing_expired')
    vi.useRealTimers()
  })

  it('closes the capability on logout or authentication loss', async () => {
    const { client, socket } = setup()
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'PowerPoint',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 1,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
    })
    client.revoke('auth_required')
    expect(socket.readyState).toBe(3)
    expect(client.status()).toBe('disconnected:auth_required')
  })

  it('uses v2 only with a fixed retrieval proxy and dispatches negotiated web requests', async () => {
    const socket = new FakeSocket()
    const retrievalProxy = vi.fn(async () => new TextEncoder().encode('{"results":[]}'))
    const agentProxy = vi.fn()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken: async () => 'access-token',
      proxy: agentProxy,
      retrievalProxy,
      onPending() {},
    })
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    expect(JSON.parse(socket.sent[0]!)).toEqual({
      version: 2,
      type: 'pc.negotiate',
      verification_code: '123456',
      capabilities: ['agent.v1', 'web-search.v1', 'web-fetch.v1', 'image-search.v1'],
    })
    socket.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 2,
      capabilities: ['agent.v1', 'web-search.v1'],
    })
    expect(JSON.parse(socket.sent[1]!)).toEqual({
      version: 2,
      type: 'pc.claim',
      verification_code: '123456',
      capabilities: ['agent.v1', 'web-search.v1', 'web-fetch.v1', 'image-search.v1'],
    })
    socket.message({
      version: 2,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
      capabilities: ['agent.v1', 'web-search.v1'],
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 2,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
      capabilities: ['agent.v1', 'web-search.v1'],
    })
    socket.message({
      version: 2,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
      capability_name: 'web-search.v1',
      body: { query: 'office', max_results: 3 },
    })
    await vi.waitFor(() => expect(retrievalProxy).toHaveBeenCalled())
    expect(retrievalProxy).toHaveBeenCalledWith(
      'web-search.v1',
      { query: 'office', max_results: 3 },
      expect.any(AbortSignal),
    )
    expect(agentProxy).not.toHaveBeenCalled()
    await vi.waitFor(() =>
      expect(socket.sent.map((raw) => JSON.parse(raw))).toContainEqual(
        expect.objectContaining({ version: 2, type: 'pc.done' }),
      ),
    )
  })

  it('explicitly negotiates the Office protocol and follows a Relay-asserted v1 pairing', async () => {
    const socket = new FakeSocket()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken: async () => 'access-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      negotiateCapabilities: true,
      onPending() {},
    })
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    expect(JSON.parse(socket.sent[0]!)).toEqual({
      version: 2,
      type: 'pc.negotiate',
      verification_code: '123456',
      capabilities: ['agent.v1'],
    })
    socket.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 1,
      capabilities: ['agent.v1'],
    })
    expect(JSON.parse(socket.sent[1]!)).toEqual({
      version: 1,
      type: 'pc.claim',
      verification_code: '123456',
    })
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    expect(client.status()).toBe('awaiting_approval')
  })

  it('does not silently downgrade when Relay skips explicit negotiation', async () => {
    const socket = new FakeSocket()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken: async () => 'access-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      negotiateCapabilities: true,
      onPending() {},
    })
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    expect(client.status()).toBe('disconnected:protocol_violation')
  })
})
