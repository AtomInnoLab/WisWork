import { describe, expect, it, vi } from 'vitest'
import {
  createOfficeRelayClient,
  officeRelayEndpointFromEnv,
  type RelaySocket,
} from '../src/main/office-relay-client'
import type { OfficeRelayBinding } from '../src/main/office-relay-binding-store'

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
  it.each(['account', 'token'] as const)(
    'does not connect when revoked while awaiting %s validation',
    async (phase) => {
      let releaseAccount!: (value: { loggedIn: boolean }) => void
      let releaseToken!: (value: string | null) => void
      const account = new Promise<{ loggedIn: boolean }>((resolve) => {
        releaseAccount = resolve
      })
      const token = new Promise<string | null>((resolve) => {
        releaseToken = resolve
      })
      const connect = vi.fn(() => new FakeSocket())
      const getAccessToken = vi.fn(() => token)
      const client = createOfficeRelayClient({
        endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
        connect,
        getValidAccountStatus: () => account,
        getAccessToken,
        proxy: async () => ({ status: 200, body: new Uint8Array() }),
        onPending() {},
      })

      const claim = client.claim('123456')
      if (phase === 'token') {
        releaseAccount({ loggedIn: true })
        await vi.waitFor(() => expect(getAccessToken).toHaveBeenCalledOnce())
      }
      client.revoke('logout')
      releaseAccount({ loggedIn: true })
      releaseToken('access-token')

      await expect(claim).rejects.toThrow('relay_connection_failed')
      expect(connect).not.toHaveBeenCalled()
      expect(client.status()).toBe('disconnected:logout')
    },
  )

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

  it('keeps ordinary v2 code pairing without enhanced fields when persistence is disabled', async () => {
    const socket = new FakeSocket()
    const onBinding = vi.fn()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'access-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      negotiateCapabilities: true,
      persistentPairing: false,
      onBinding,
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
    expect(JSON.parse(socket.sent[0]!)).not.toHaveProperty('features')

    socket.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 2,
      capabilities: ['agent.v1'],
    })
    expect(JSON.parse(socket.sent[1]!)).toEqual({
      version: 2,
      type: 'pc.claim',
      verification_code: '123456',
      capabilities: ['agent.v1'],
    })
    expect(onBinding).not.toHaveBeenCalled()
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

  it('reports Relay session expiry explicitly', async () => {
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
    expect(client.status()).toBe('paired')
    socket.message({ version: 1, type: 'relay.error', code: 'session_expired' })
    expect(client.status()).toBe('disconnected:session_expired')
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

  it('lets Relay own the 300s deadline and tolerates its late cancel after the 305s watchdog', async () => {
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
    await vi.advanceTimersByTimeAsync(300_000)
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

  it('does not apply the renewable 30 minute Relay idle TTL as a local hard expiry', async () => {
    vi.useFakeTimers()
    const { client, socket } = setup()
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
    await vi.advanceTimersByTimeAsync(31 * 60 * 1_000)
    expect(client.status()).toBe('paired')
    client.revoke('test_complete')
    vi.useRealTimers()
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

  it('conditionally negotiates pairing-resume.v1 and captures binding metadata after approval', async () => {
    const socket = new FakeSocket()
    const onBinding = vi.fn(async () => undefined)
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'fresh-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      negotiateCapabilities: true,
      persistentPairing: true,
      onBinding,
      onPending() {},
    })
    const claim = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claim
    expect(JSON.parse(socket.sent[0]!)).toEqual({
      version: 2,
      type: 'pc.negotiate',
      verification_code: '123456',
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    socket.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 2,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    expect(JSON.parse(socket.sent[1]!)).toEqual({
      version: 2,
      type: 'pc.claim',
      verification_code: '123456',
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    socket.message({
      version: 2,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    await client.approve('pairing_12345678')
    expect(JSON.parse(socket.sent.at(-1)!)).toEqual({
      version: 2,
      type: 'pc.approve',
      pairing_id: 'pairing_12345678',
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    socket.message({
      version: 2,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
      binding_id: 'binding_word_12345678',
    })
    await vi.waitFor(() =>
      expect(onBinding).toHaveBeenCalledWith({
        bindingId: 'binding_word_12345678',
        accountId: 'local-account',
        host: 'Word',
        origin: 'https://office.8-216-134-194.sslip.io',
        capabilities: ['agent.v1'],
        createdAt: expect.any(Number),
      }),
    )
    expect(client.status()).toBe('paired')
  })

  it('accepts only the exact enhanced short-session fallback when durable enrollment aborts', async () => {
    const sockets: FakeSocket[] = []
    const onBinding = vi.fn(async () => undefined)
    let persistenceAvailable = true
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => {
        const socket = new FakeSocket()
        sockets.push(socket)
        return socket
      },
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'fresh-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      negotiateCapabilities: true,
      persistentPairing: () => persistenceAvailable,
      onBinding,
      onPending() {},
    })

    const claim = client.claim('123456')
    await vi.waitFor(() => expect(sockets).toHaveLength(1))
    sockets[0]!.open()
    await claim
    sockets[0]!.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 2,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    sockets[0]!.message({
      version: 2,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    await client.approve('pairing_12345678')
    sockets[0]!.message({
      version: 2,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
      capabilities: ['agent.v1'],
      features: [],
    })
    expect(client.status()).toBe('paired')
    expect(onBinding).not.toHaveBeenCalled()

    client.revoke('next_claim')
    persistenceAvailable = false
    const ordinaryClaim = client.claim('654321')
    await vi.waitFor(() => expect(sockets).toHaveLength(2))
    sockets[1]!.open()
    await ordinaryClaim
    expect(JSON.parse(sockets[1]!.sent[0]!)).toEqual({
      version: 2,
      type: 'pc.negotiate',
      verification_code: '654321',
      capabilities: ['agent.v1'],
    })
  })

  it('accepts an explicit features-empty approval when Relay disables durable pairing', async () => {
    const socket = new FakeSocket()
    const onBinding = vi.fn(async () => undefined)
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'fresh-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      negotiateCapabilities: true,
      persistentPairing: true,
      onBinding,
      onPending() {},
    })

    const claim = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claim
    socket.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 2,
      capabilities: ['agent.v1'],
      features: [],
    })
    socket.message({
      version: 2,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
      capabilities: ['agent.v1'],
      features: [],
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 2,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
      capabilities: ['agent.v1'],
      features: [],
    })

    expect(client.status()).toBe('paired')
    expect(onBinding).not.toHaveBeenCalled()
  })

  it.each([
    {
      label: 'omitted fallback feature marker',
      approved: {},
    },
    {
      label: 'fallback with a binding id',
      approved: { features: [], binding_id: 'binding_word_12345678' },
    },
    {
      label: 'unnegotiated fallback feature',
      approved: { features: ['future-feature.v1'] },
    },
  ])('rejects $label after enhanced enrollment', async ({ approved }) => {
    const socket = new FakeSocket()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'fresh-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onPending() {},
    })
    const claim = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claim
    socket.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 2,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    socket.message({
      version: 2,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 2,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
      capabilities: ['agent.v1'],
      ...approved,
    })
    expect(client.status()).toBe('disconnected:protocol_violation')
  })

  it('revokes the durable binding and reports not remembered when encrypted persistence fails', async () => {
    const sockets: FakeSocket[] = []
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => {
        const socket = new FakeSocket()
        sockets.push(socket)
        return socket
      },
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: vi
        .fn()
        .mockResolvedValueOnce('pair-token')
        .mockResolvedValueOnce('revoke-token'),
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onBinding: async () => {
        throw new Error('disk_failure')
      },
      onPending() {},
    })
    const claim = client.claim('123456')
    await vi.waitFor(() => expect(sockets).toHaveLength(1))
    sockets[0]!.open()
    await claim
    sockets[0]!.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 2,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    sockets[0]!.message({
      version: 2,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    await client.approve('pairing_12345678')
    sockets[0]!.message({
      version: 2,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
      binding_id: 'binding_word_12345678',
    })

    await vi.waitFor(() => expect(sockets).toHaveLength(2))
    expect(client.status()).not.toBe('paired')
    sockets[1]!.open()
    await vi.waitFor(() => expect(sockets[1]!.sent).toHaveLength(1))
    expect(JSON.parse(sockets[1]!.sent[0]!)).toEqual({
      version: 2,
      type: 'pc.revoke_binding',
      binding_id: 'binding_word_12345678',
    })
    sockets[1]!.message({
      version: 2,
      type: 'pc.binding_revoked',
      binding_id: 'binding_word_12345678',
    })
    await vi.waitFor(() => expect(client.status()).toBe('disconnected:binding_not_remembered'))
  })

  it('serializes approval persistence before requests and ignores an exact duplicate approval', async () => {
    const socket = new FakeSocket()
    let releaseBinding!: () => void
    const bindingSaved = new Promise<void>((resolve) => {
      releaseBinding = resolve
    })
    const onBinding = vi.fn(() => bindingSaved)
    const proxy = vi.fn(async () => ({ status: 200, body: new Uint8Array() }))
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'token',
      proxy,
      persistentPairing: true,
      onBinding,
      onPending() {},
    })
    const claim = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claim
    socket.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 2,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    socket.message({
      version: 2,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    await client.approve('pairing_12345678')
    const approved = {
      version: 2,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
      binding_id: 'binding_word_12345678',
    }
    socket.message(approved)
    await vi.waitFor(() => expect(onBinding).toHaveBeenCalledOnce())
    socket.message(approved)
    socket.message({
      version: 2,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
      capability_name: 'agent.v1',
      body: {},
    })
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(client.status()).not.toBe('disconnected:protocol_violation')
    expect(proxy).not.toHaveBeenCalled()

    releaseBinding()
    await vi.waitFor(() => expect(client.status()).toBe('paired'))
    await vi.waitFor(() => expect(proxy).toHaveBeenCalledOnce())
    expect(onBinding).toHaveBeenCalledOnce()
  })

  it('rejects a claimed feature upgrade beyond the negotiated intersection', async () => {
    const socket = new FakeSocket()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onPending() {},
    })
    const claim = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claim
    socket.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 2,
      capabilities: ['agent.v1'],
      features: [],
    })
    socket.message({
      version: 2,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    await vi.waitFor(() => expect(client.status()).toBe('disconnected:protocol_violation'))
  })

  it('uses exact legacy v1 frames after enhanced negotiation selects pairing_version 1', async () => {
    const socket = new FakeSocket()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onPending() {},
    })
    const claim = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claim
    socket.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 1,
      capabilities: ['agent.v1'],
      features: [],
    })
    expect(JSON.parse(socket.sent.at(-1)!)).toEqual({
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

  it.each(['invalid_frame', 'unknown_type'] as const)(
    'falls back once only for explicit old-schema Relay error %s',
    async (code) => {
      const sockets: FakeSocket[] = []
      const getAccessToken = vi
        .fn()
        .mockResolvedValueOnce('token-one')
        .mockResolvedValueOnce('token-two')
      const client = createOfficeRelayClient({
        endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
        connect: () => {
          const socket = new FakeSocket()
          sockets.push(socket)
          return socket
        },
        getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
        getAccessToken,
        proxy: async () => ({ status: 200, body: new Uint8Array() }),
        negotiateCapabilities: true,
        persistentPairing: true,
        onPending() {},
      })
      const claim = client.claim('123456')
      await vi.waitFor(() => expect(sockets).toHaveLength(1))
      sockets[0]!.open()
      await claim
      expect(JSON.parse(sockets[0]!.sent[0]!)).toHaveProperty('features', ['pairing-resume.v1'])
      sockets[0]!.message({ version: 2, type: 'relay.error', code })
      await vi.waitFor(() => expect(sockets).toHaveLength(2))
      sockets[1]!.open()
      await vi.waitFor(() => expect(sockets[1]!.sent).toHaveLength(1))
      expect(JSON.parse(sockets[1]!.sent[0]!)).toEqual({
        version: 2,
        type: 'pc.negotiate',
        verification_code: '123456',
        capabilities: ['agent.v1'],
      })
      expect(getAccessToken).toHaveBeenNthCalledWith(1)
      expect(getAccessToken).toHaveBeenNthCalledWith(2)
    },
  )

  it.each(['invalid_code', 'auth_required', 'resume_rate_limited'] as const)(
    'does not legacy-fallback for semantic Relay error %s',
    async (code) => {
      const sockets: FakeSocket[] = []
      const client = createOfficeRelayClient({
        endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
        connect: () => {
          const socket = new FakeSocket()
          sockets.push(socket)
          return socket
        },
        getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
        getAccessToken: async () => 'token',
        proxy: async () => ({ status: 200, body: new Uint8Array() }),
        persistentPairing: true,
        onPending() {},
      })
      const claim = client.claim('123456')
      await vi.waitFor(() => expect(sockets).toHaveLength(1))
      sockets[0]!.open()
      await claim
      sockets[0]!.message({ version: 2, type: 'relay.error', code })
      await new Promise((resolve) => setTimeout(resolve, 0))
      expect(sockets).toHaveLength(1)
    },
  )

  it('does not downgrade on a network close and keeps the next claim enhanced', async () => {
    const sockets: FakeSocket[] = []
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => {
        const socket = new FakeSocket()
        sockets.push(socket)
        return socket
      },
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onPending() {},
    })
    const first = client.claim('123456')
    await vi.waitFor(() => expect(sockets).toHaveLength(1))
    sockets[0]!.open()
    await first
    sockets[0]!.close()
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(sockets).toHaveLength(1)

    const second = client.claim('123456')
    await vi.waitFor(() => expect(sockets).toHaveLength(2))
    sockets[1]!.open()
    await second
    expect(JSON.parse(sockets[1]!.sent[0]!)).toHaveProperty('features', ['pairing-resume.v1'])
  })

  it('uses a fresh access token for pc.resume and accepts waiting_for_office then standard v2 approval', async () => {
    const socket = new FakeSocket()
    const getAccessToken = vi.fn(async () => 'resume-token')
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken,
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      negotiateCapabilities: true,
      persistentPairing: true,
      onPending() {},
    })
    const binding: OfficeRelayBinding = {
      bindingId: 'binding_word_12345678',
      accountId: 'local-account',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      capabilities: ['agent.v1'],
      createdAt: 1,
    }
    const resume = client.resume(binding)
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await resume
    expect(getAccessToken).toHaveBeenCalledOnce()
    expect(JSON.parse(socket.sent[0]!)).toEqual({
      version: 2,
      type: 'pc.resume',
      binding_id: 'binding_word_12345678',
      capabilities: ['agent.v1'],
    })
    socket.message({ version: 2, type: 'pc.waiting_for_office' })
    expect(client.status()).toBe('waiting_for_office')
    socket.message({
      version: 2,
      type: 'pc.approved',
      session_id: 'session_resume_12345678',
      capability: 'fresh-capability',
      expires_in: 1800,
      capabilities: ['agent.v1'],
    })
    expect(client.status()).toBe('paired')
  })

  it('sends an exact authenticated revocation frame and requires its exact acknowledgement', async () => {
    const socket = new FakeSocket()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'revoke-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onPending() {},
    })
    const revocation = client.revokeBinding('binding_word_12345678', 'local-account')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await vi.waitFor(() => expect(socket.sent).toHaveLength(1))
    expect(JSON.parse(socket.sent[0]!)).toEqual({
      version: 2,
      type: 'pc.revoke_binding',
      binding_id: 'binding_word_12345678',
    })
    socket.message({
      version: 2,
      type: 'pc.binding_revoked',
      binding_id: 'binding_word_12345678',
    })
    await expect(revocation).resolves.toBeUndefined()
  })

  it('treats binding_unavailable as idempotent success only for an authenticated revocation', async () => {
    const socket = new FakeSocket()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'revoke-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onPending() {},
    })
    const revocation = client.revokeBinding('binding_word_12345678', 'local-account')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await vi.waitFor(() => expect(socket.sent).toHaveLength(1))
    socket.message({ version: 2, type: 'relay.error', code: 'binding_unavailable' })
    await expect(revocation).resolves.toBeUndefined()
    expect(client.status()).toBe('disconnected:binding_revoked')

    const claimSocket = new FakeSocket()
    const claiming = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => claimSocket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'claim-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onPending() {},
    })
    const claim = claiming.claim('123456')
    await vi.waitFor(() => expect(claimSocket.listeners.has('open')).toBe(true))
    claimSocket.open()
    await claim
    claimSocket.message({ version: 2, type: 'relay.error', code: 'binding_unavailable' })
    expect(claiming.status()).toBe('disconnected:relay_error')
  })

  it('refuses to revoke an old-account binding with replacement-account credentials', async () => {
    const connect = vi.fn(() => new FakeSocket())
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'replacement-account' }),
      getAccessToken: async () => 'replacement-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onPending() {},
    })

    await expect(client.revokeBinding('binding_word_12345678', 'old-account')).rejects.toThrow(
      'auth_required',
    )
    expect(connect).not.toHaveBeenCalled()
  })

  it('rechecks the expected revocation account after loading a token', async () => {
    let accountId = 'old-account'
    let releaseToken!: (token: string) => void
    const token = new Promise<string>((resolve) => {
      releaseToken = resolve
    })
    const connect = vi.fn(() => new FakeSocket())
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: accountId }),
      getAccessToken: () => token,
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onPending() {},
    })

    const revoking = client.revokeBinding('binding_word_12345678', 'old-account')
    await vi.waitFor(() => expect(client.status()).toBe('disconnected:new_revocation'))
    accountId = 'replacement-account'
    releaseToken('replacement-token')
    await expect(revoking).rejects.toThrow('auth_required')
    expect(connect).not.toHaveBeenCalled()
  })
})
