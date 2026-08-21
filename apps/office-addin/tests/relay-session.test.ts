import { describe, expect, it, vi } from 'vitest'
import {
  OFFICE_RELAY_URL,
  createOfficeRelaySession,
  officeTransportMode,
  type RelayWebSocket,
} from '../src/relay/session.js'

class FakeSocket implements RelayWebSocket {
  static readonly OPEN = 1
  readonly OPEN = 1
  readyState = 0
  sent: string[] = []
  onopen: (() => void) | null = null
  onmessage: ((event: { data: unknown }) => void) | null = null
  onclose: (() => void) | null = null
  onerror: (() => void) | null = null
  send(value: string) { this.sent.push(value) }
  close() { this.readyState = 3; this.onclose?.() }
  open() { this.readyState = 1; this.onopen?.() }
  receive(value: unknown) { this.onmessage?.({ data: value }) }
}

const frame = (socket: FakeSocket, index: number) => JSON.parse(socket.sent[index]!)

describe('Office cloud relay session', () => {
  it('uses the fixed secure relay by default and loopback only as explicit rollback', () => {
    expect(OFFICE_RELAY_URL).toBe('wss://office.8-216-134-194.sslip.io/office-relay')
    expect(officeTransportMode({})).toBe('relay')
    expect(officeTransportMode({ VITE_WISWORK_OFFICE_TRANSPORT: 'loopback' })).toBe('loopback')
    expect(() => officeTransportMode({ VITE_WISWORK_OFFICE_TRANSPORT: 'http' })).toThrow(
      'invalid_office_transport',
    )
  })

  it('pairs on one WSS socket and exposes only the six-digit verification code', async () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({ createSocket: (url) => {
      expect(url).toBe(OFFICE_RELAY_URL)
      return socket
    } })
    const connecting = session.connect('word')
    expect(session.snapshot()).toEqual({ status: 'connecting' })
    socket.open()
    expect(frame(socket, 0)).toEqual({ version: 1, type: 'office.create', host: 'Word' })
    socket.receive(JSON.stringify({ version: 1, type: 'office.created', pairing_id: 'pair_1', polling_secret: 'poll_1', verification_code: '123456', expires_in: 120 }))
    expect(session.snapshot()).toEqual({ status: 'pending', verificationCode: '123456' })
    socket.receive(JSON.stringify({ version: 1, type: 'office.approved', session_id: 'session_1', capability: 'cap_1', expires_in: 1800 }))
    await connecting
    expect(session.snapshot()).toEqual({ status: 'connected' })
  })

  it('streams bounded SSE events, sends cancel, and completes once', async () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({ createSocket: () => socket })
    const connecting = session.connect('excel'); socket.open()
    socket.receive(JSON.stringify({ version: 1, type: 'office.created', pairing_id: 'p', polling_secret: 's', verification_code: '654321', expires_in: 120 }))
    socket.receive(JSON.stringify({ version: 1, type: 'office.approved', session_id: 'session', capability: 'cap', expires_in: 1800 }))
    await connecting
    const controller = new AbortController()
    const responsePending = session.authenticatedFetch('/v1/office/messages', { method: 'POST', body: '{"model":"fixed"}', signal: controller.signal })
    const request = frame(socket, 1)
    expect(request).toMatchObject({ version: 1, type: 'office.request', session_id: 'session', capability: 'cap', body: { model: 'fixed' } })
    socket.receive(JSON.stringify({ version: 1, type: 'relay.start', session_id: 'session', request_id: request.request_id, status: 200, content_type: 'text/event-stream' }))
    const response = await responsePending
    socket.receive(JSON.stringify({ version: 1, type: 'relay.chunk', session_id: 'session', request_id: request.request_id, sequence: 0, data: btoa('data: [DONE]\n') }))
    socket.receive(JSON.stringify({ version: 1, type: 'relay.done', session_id: 'session', request_id: request.request_id }))
    await expect(response.text()).resolves.toBe('data: [DONE]\n')
    controller.abort(); controller.abort()
    expect(socket.sent.map((value) => JSON.parse(value)).filter((value) => value.type === 'office.cancel')).toHaveLength(0)
  })

  it('cancels an active request, revokes on disconnect, and can pair again', async () => {
    const made: FakeSocket[] = []
    const session = createOfficeRelaySession({ createSocket: () => {
      const socket = new FakeSocket(); made.push(socket); return socket
    } })
    const first = session.connect('powerpoint'); made[0]!.open()
    made[0]!.receive(JSON.stringify({ version: 1, type: 'office.created', pairing_id: 'p1', polling_secret: 's1', verification_code: '111111', expires_in: 120 }))
    made[0]!.receive(JSON.stringify({ version: 1, type: 'office.approved', session_id: 'session1', capability: 'c1', expires_in: 1800 }))
    await first
    const controller = new AbortController()
    const cancelled = session.authenticatedFetch('/v1/office/messages', { method: 'POST', body: '{}', signal: controller.signal })
    controller.abort()
    await expect(cancelled).rejects.toThrow('relay_cancelled')
    expect(made[0]!.sent.map((value) => JSON.parse(value)).filter((value) => value.type === 'office.cancel')).toHaveLength(1)
    made[0]!.close()
    expect(session.snapshot()).toEqual({ status: 'offline' })

    const second = session.connect('powerpoint'); made[1]!.open()
    made[1]!.receive(JSON.stringify({ version: 1, type: 'office.created', pairing_id: 'p2', polling_secret: 's2', verification_code: '222222', expires_in: 120 }))
    made[1]!.receive(JSON.stringify({ version: 1, type: 'office.approved', session_id: 'session2', capability: 'c2', expires_in: 1800 }))
    await second
    expect(session.snapshot()).toEqual({ status: 'connected' })
  })

  it('fails closed on binary, malformed, extra-key, oversized, and unknown frames', async () => {
    for (const hostile of [new Uint8Array([1]), '{', JSON.stringify({ version: 1, type: 'office.created', extra: true }), 'x'.repeat(16 * 1024 + 1)]) {
      const socket = new FakeSocket()
      const session = createOfficeRelaySession({ createSocket: () => socket })
      const connecting = session.connect('word'); socket.open(); socket.receive(hostile)
      await connecting
      expect(session.snapshot()).toEqual({ status: 'offline' })
      expect(socket.readyState).toBe(3)
    }
  })

  it('rejects oversized requests before send and chunks before relay.start', async () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({ createSocket: () => socket })
    const connecting = session.connect('word'); socket.open()
    socket.receive(JSON.stringify({ version: 1, type: 'office.created', pairing_id: 'p', polling_secret: 's', verification_code: '123456', expires_in: 120 }))
    socket.receive(JSON.stringify({ version: 1, type: 'office.approved', session_id: 'sid', capability: 'cap', expires_in: 1800 }))
    await connecting
    await expect(session.authenticatedFetch('/v1/office/messages', { method: 'POST', body: 'x'.repeat(256 * 1024 + 1) })).rejects.toThrow('relay_request_too_large')
    expect(socket.sent).toHaveLength(1)

    const response = session.authenticatedFetch('/v1/office/messages', { method: 'POST', body: '{}' })
    const request = frame(socket, 1)
    socket.receive(JSON.stringify({ version: 1, type: 'relay.chunk', session_id: 'sid', request_id: request.request_id, sequence: 0, data: btoa('data') }))
    await expect(response).rejects.toThrow('relay_disconnected')
    expect(session.snapshot()).toEqual({ status: 'offline' })
  })

  it.each(['{', '[]', 'null', '"text"'])('rejects non-object request JSON %s before send', async (body) => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({ createSocket: () => socket })
    const connecting = session.connect('word'); socket.open()
    socket.receive(JSON.stringify({ version: 1, type: 'office.created', pairing_id: 'p', polling_secret: 's', verification_code: '123456', expires_in: 120 }))
    socket.receive(JSON.stringify({ version: 1, type: 'office.approved', session_id: 'sid', capability: 'cap', expires_in: 1800 }))
    await connecting
    await expect(session.authenticatedFetch('/v1/office/messages', { method: 'POST', body })).rejects.toThrow('relay_invalid_request')
    expect(socket.sent).toHaveLength(1)
  })

  it.each([204, 205, 304])('fails closed when relay.start has non-streaming status %i', async (status) => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({ createSocket: () => socket })
    const connecting = session.connect('word'); socket.open()
    socket.receive(JSON.stringify({ version: 1, type: 'office.created', pairing_id: 'p', polling_secret: 's', verification_code: '123456', expires_in: 120 }))
    socket.receive(JSON.stringify({ version: 1, type: 'office.approved', session_id: 'sid', capability: 'cap', expires_in: 1800 }))
    await connecting
    const response = session.authenticatedFetch('/v1/office/messages', { method: 'POST', body: '{}' })
    const request = frame(socket, 1)
    socket.receive(JSON.stringify({ version: 1, type: 'relay.start', session_id: 'sid', request_id: request.request_id, status, content_type: 'text/event-stream' }))
    await expect(response).rejects.toThrow('relay_disconnected')
    expect(session.snapshot()).toEqual({ status: 'offline' })
  })

  it('cleans up when request send throws and enforces pairing frame order', async () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({ createSocket: () => socket })
    const connecting = session.connect('word'); socket.open()
    socket.receive(JSON.stringify({ version: 1, type: 'office.approved', session_id: 'sid', capability: 'cap', expires_in: 1800 }))
    await connecting
    expect(session.snapshot()).toEqual({ status: 'offline' })

    const throwing = new FakeSocket()
    const connected = createOfficeRelaySession({ createSocket: () => throwing })
    const pairing = connected.connect('word'); throwing.open()
    throwing.receive(JSON.stringify({ version: 1, type: 'office.created', pairing_id: 'p', polling_secret: 's', verification_code: '123456', expires_in: 120 }))
    throwing.receive(JSON.stringify({ version: 1, type: 'office.approved', session_id: 'sid', capability: 'cap', expires_in: 1800 }))
    await pairing
    throwing.send = () => { throw new Error('socket failed') }
    await expect(connected.authenticatedFetch('/v1/office/messages', { method: 'POST', body: '{}' })).rejects.toThrow('relay_disconnected')
    expect(connected.snapshot()).toEqual({ status: 'offline' })
  })
})
