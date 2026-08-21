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
      officeRelayEndpointFromEnv({
        WISWORK_OFFICE_RELAY_URL: 'wss://user:secret@example.com/relay',
      }),
    ).toThrow('invalid_office_relay_url')
  })

  it('requires sign-in and an exact six-digit code before connecting', async () => {
    await expect(setup(false).client.claim('123456')).rejects.toThrow('auth_required')
    await expect(setup().client.claim('12345')).rejects.toThrow('invalid_verification_code')
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

  it('aborts active upstream work on relay.cancel', async () => {
    const socket = new FakeSocket()
    let aborted = false
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true }),
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
})
