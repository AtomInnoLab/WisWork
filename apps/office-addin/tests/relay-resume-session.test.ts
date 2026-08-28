import { describe, expect, it, vi } from 'vitest'
import {
  createBrowserOfficeBindingInvalidationChannel,
  createOfficeRelaySession,
  type OfficeBindingInvalidation,
  type OfficeBindingInvalidationChannel,
  type RelayWebSocket,
} from '../src/relay/session.js'
import {
  PAIRING_RESUME_FEATURE,
  type OfficeBindingEnrollment,
  type OfficeBindingStore,
  type OfficeStoredBinding,
} from '../src/relay/binding-store.js'

class FakeSocket implements RelayWebSocket {
  readyState = 0
  sent: string[] = []
  onopen: (() => void) | null = null
  onmessage: ((event: { data: unknown }) => void) | null = null
  onclose: (() => void) | null = null
  onerror: (() => void) | null = null
  send(value: string) {
    this.sent.push(value)
  }
  close() {
    if (this.readyState === 3) return
    this.readyState = 3
    this.onclose?.()
  }
  open() {
    this.readyState = 1
    this.onopen?.()
  }
  receive(value: Record<string, unknown>) {
    this.onmessage?.({ data: JSON.stringify(value) })
  }
}

const publicKey = Buffer.concat([Buffer.from([0x04]), Buffer.alloc(64, 7)]).toString('base64url')
const signature = Buffer.alloc(64, 9).toString('base64url')
const privateKey = {
  type: 'private',
  extractable: false,
  algorithm: { name: 'ECDSA', namedCurve: 'P-256' },
  usages: ['sign'],
} as unknown as CryptoKey

const enrollment: OfficeBindingEnrollment = {
  host: 'word',
  capabilities: ['agent.v1'],
  publicKey,
  privateKey,
}

const binding: OfficeStoredBinding = {
  schemaVersion: 1,
  bindingId: 'binding_12345678',
  host: 'word',
  origin: 'https://office.8-216-134-194.sslip.io',
  capabilities: ['agent.v1'],
  privateKey,
}

class FakeBindingStore implements OfficeBindingStore {
  current: OfficeStoredBinding | undefined
  saves: Array<{
    enrollment: OfficeBindingEnrollment
    bindingId: string
    approvedCapabilities: readonly string[]
  }> = []
  forgets = 0
  signatures: string[] = []

  constructor(current?: OfficeStoredBinding) {
    this.current = current
  }

  async load() {
    return this.current
  }
  async createEnrollment() {
    return enrollment
  }
  async save(
    value: OfficeBindingEnrollment,
    bindingId: string,
    approvedCapabilities: readonly string[],
  ) {
    this.saves.push({ enrollment: value, bindingId, approvedCapabilities })
    this.current = { ...binding, bindingId, capabilities: approvedCapabilities }
  }
  async sign(_binding: OfficeStoredBinding, challenge: string) {
    this.signatures.push(challenge)
    return signature
  }
  async forget() {
    this.forgets += 1
    this.current = undefined
  }
}

const flush = async () => {
  await Promise.resolve()
  await Promise.resolve()
}

const sent = (socket: FakeSocket, index = 0) => JSON.parse(socket.sent[index]!)

class FakeInvalidationChannel implements OfficeBindingInvalidationChannel {
  listeners = new Set<(message: OfficeBindingInvalidation) => void>()

  subscribe(listener: (message: OfficeBindingInvalidation) => void) {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  broadcast(message: OfficeBindingInvalidation) {
    for (const listener of this.listeners) listener(message)
  }
}

describe('Office binding invalidation channel', () => {
  it('accepts only the exact origin-scoped wire shape and emits the canonical shape', () => {
    const postMessage = vi.fn<(message: unknown) => void>()
    let rawChannel:
      | {
          onmessage: ((event: MessageEvent<unknown>) => void) | null
          postMessage(message: unknown): void
        }
      | undefined
    const channel = createBrowserOfficeBindingInvalidationChannel((name) => {
      expect(name).toBe('wiswork-office-pairing-v1')
      rawChannel = { onmessage: null, postMessage }
      return rawChannel
    })!
    const listener = vi.fn()
    channel.subscribe(listener)

    rawChannel!.onmessage?.({
      data: {
        version: 1,
        type: 'binding.forgotten',
        origin: binding.origin,
        host: 'word',
        binding_id: binding.bindingId,
        extra: true,
      },
    } as MessageEvent<unknown>)
    rawChannel!.onmessage?.({
      data: {
        version: 1,
        type: 'binding.forgotten',
        origin: 'https://attacker.example',
        host: 'word',
        binding_id: binding.bindingId,
      },
    } as MessageEvent<unknown>)
    expect(listener).not.toHaveBeenCalled()

    rawChannel!.onmessage?.({
      data: {
        version: 1,
        type: 'binding.forgotten',
        origin: binding.origin,
        host: 'word',
        binding_id: binding.bindingId,
      },
    } as MessageEvent<unknown>)
    expect(listener).toHaveBeenCalledWith({
      origin: binding.origin,
      host: 'word',
      bindingId: binding.bindingId,
    })

    channel.broadcast({
      origin: binding.origin,
      host: 'word',
      bindingId: binding.bindingId,
    })
    expect(postMessage).toHaveBeenCalledWith({
      version: 1,
      type: 'binding.forgotten',
      origin: binding.origin,
      host: 'word',
      binding_id: binding.bindingId,
    })
  })
})

describe('Office persistent relay session', () => {
  it('enrolls once with a control feature separate from callable capabilities', async () => {
    const store = new FakeBindingStore()
    const sockets: FakeSocket[] = []
    const session = createOfficeRelaySession({
      capabilities: ['agent.v1'],
      bindingStore: store,
      createSocket: () => {
        const socket = new FakeSocket()
        sockets.push(socket)
        return socket
      },
    })

    const connecting = session.connect('word')
    await flush()
    sockets[0]!.open()
    expect(sent(sockets[0]!)).toEqual({
      version: 2,
      type: 'office.create',
      host: 'Word',
      capabilities: ['agent.v1'],
      features: [PAIRING_RESUME_FEATURE],
      binding_public_key: publicKey,
    })
    sockets[0]!.receive({
      version: 2,
      type: 'office.created',
      pairing_id: 'pairing_12345678',
      verification_code: '123456',
      expires_in: 120,
      features: [PAIRING_RESUME_FEATURE],
    })
    sockets[0]!.receive({
      version: 2,
      type: 'office.approved',
      session_id: 'session_12345678',
      capability: 'session_capability_12345678',
      expires_in: 1800,
      capabilities: ['agent.v1'],
      features: [PAIRING_RESUME_FEATURE],
      binding_id: 'binding_12345678',
    })

    await connecting
    expect(store.saves).toEqual([
      {
        enrollment,
        bindingId: 'binding_12345678',
        approvedCapabilities: ['agent.v1'],
      },
    ])
    expect(session.snapshot()).toEqual({ status: 'connected', capabilities: ['agent.v1'] })
    await expect(session.capabilityFetch(PAIRING_RESUME_FEATURE as never, {})).rejects.toThrow(
      'relay_capability_unavailable',
    )
  })

  it('keeps the short session but marks it not remembered when binding persistence fails', async () => {
    const store = new FakeBindingStore()
    store.save = async () => {
      throw new Error('binding_storage_unavailable')
    }
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({
      capabilities: ['agent.v1'],
      bindingStore: store,
      createSocket: () => socket,
    })
    const connecting = session.connect('word')
    await flush()
    socket.open()
    socket.receive({
      version: 2,
      type: 'office.created',
      pairing_id: 'pairing_12345678',
      verification_code: '123456',
      expires_in: 120,
      features: [PAIRING_RESUME_FEATURE],
    })
    socket.receive({
      version: 2,
      type: 'office.approved',
      session_id: 'session_12345678',
      capability: 'session_capability_12345678',
      expires_in: 1800,
      capabilities: ['agent.v1'],
      features: [PAIRING_RESUME_FEATURE],
      binding_id: 'binding_12345678',
    })

    await connecting
    expect(session.snapshot()).toEqual({
      status: 'connected',
      capabilities: ['agent.v1'],
      remembered: false,
    })
  })

  it('resumes by signing the exact Relay challenge and preserves waiting_for_pc', async () => {
    const store = new FakeBindingStore(binding)
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({
      capabilities: ['agent.v1'],
      bindingStore: store,
      createSocket: () => socket,
    })

    const connecting = session.connect('word')
    await flush()
    socket.open()
    expect(sent(socket)).toEqual({
      version: 2,
      type: 'office.resume',
      binding_id: 'binding_12345678',
      host: 'Word',
      capabilities: ['agent.v1'],
    })
    socket.receive({
      version: 2,
      type: 'office.challenge',
      binding_id: 'binding_12345678',
      challenge: 'challenge_12345678',
      expires_in: 30,
    })
    await flush()
    expect(store.signatures).toEqual(['challenge_12345678'])
    expect(sent(socket, 1)).toEqual({
      version: 2,
      type: 'office.prove',
      binding_id: 'binding_12345678',
      challenge: 'challenge_12345678',
      signature,
    })
    socket.receive({ version: 2, type: 'office.waiting_for_pc' })
    expect(session.snapshot()).toEqual({ status: 'waiting_for_pc' })
    expect(store.forgets).toBe(0)
    socket.receive({
      version: 2,
      type: 'office.approved',
      session_id: 'fresh_session_12345678',
      capability: 'fresh_capability_12345678',
      expires_in: 1800,
      capabilities: ['agent.v1'],
    })
    await connecting
    expect(session.snapshot()).toEqual({ status: 'connected', capabilities: ['agent.v1'] })
  })

  it('falls back exactly once to ordinary v2 when an old Relay closes enhanced enrollment', async () => {
    const store = new FakeBindingStore()
    const sockets: FakeSocket[] = []
    const session = createOfficeRelaySession({
      capabilities: ['agent.v1'],
      bindingStore: store,
      createSocket: () => {
        const socket = new FakeSocket()
        sockets.push(socket)
        return socket
      },
    })

    const connecting = session.connect('word')
    await flush()
    sockets[0]!.open()
    sockets[0]!.close()
    await flush()
    expect(sockets).toHaveLength(2)
    sockets[1]!.open()
    expect(sent(sockets[1]!)).toEqual({
      version: 2,
      type: 'office.create',
      host: 'Word',
      capabilities: ['agent.v1'],
    })
    sockets[1]!.receive({
      version: 2,
      type: 'office.created',
      pairing_id: 'pairing_12345678',
      verification_code: '123456',
      expires_in: 120,
    })
    sockets[1]!.receive({
      version: 2,
      type: 'office.approved',
      session_id: 'session_12345678',
      capability: 'session_capability_12345678',
      expires_in: 1800,
      capabilities: ['agent.v1'],
    })
    await connecting
    expect(store.saves).toHaveLength(0)
  })

  it.each(['binding_unavailable', 'binding_revoked', 'invalid_proof', 'capability_not_negotiated'])(
    'clears terminal %s bindings and returns to enhanced first pairing',
    async (code) => {
      const store = new FakeBindingStore(binding)
      const sockets: FakeSocket[] = []
      const session = createOfficeRelaySession({
        capabilities: ['agent.v1'],
        bindingStore: store,
        createSocket: () => {
          const socket = new FakeSocket()
          sockets.push(socket)
          return socket
        },
      })
      void session.connect('word')
      await flush()
      sockets[0]!.open()
      sockets[0]!.receive({ version: 2, type: 'relay.error', code })
      await flush()
      expect(store.forgets).toBe(1)
      expect(sockets).toHaveLength(2)
      sockets[1]!.open()
      expect(sent(sockets[1]!)).toMatchObject({
        type: 'office.create',
        features: [PAIRING_RESUME_FEATURE],
      })
    },
  )

  it('treats session_revoked as transient PC loss and resumes the retained binding', async () => {
    const store = new FakeBindingStore(binding)
    const sockets: FakeSocket[] = []
    const scheduled: Array<() => void> = []
    const session = createOfficeRelaySession({
      capabilities: ['agent.v1'],
      bindingStore: store,
      createSocket: () => {
        const socket = new FakeSocket()
        sockets.push(socket)
        return socket
      },
      schedule: (callback) => {
        scheduled.push(callback)
        return scheduled.length
      },
      cancelSchedule: () => undefined,
      random: () => 0.5,
    })
    const connecting = session.connect('word')
    await flush()
    sockets[0]!.open()
    sockets[0]!.receive({
      version: 2,
      type: 'office.challenge',
      binding_id: binding.bindingId,
      challenge: 'challenge_12345678',
      expires_in: 30,
    })
    await flush()
    sockets[0]!.receive({
      version: 2,
      type: 'office.approved',
      session_id: 'session_12345678',
      capability: 'session_capability_12345678',
      expires_in: 1800,
      capabilities: ['agent.v1'],
    })
    await connecting

    sockets[0]!.receive({ version: 2, type: 'relay.error', code: 'session_revoked' })
    await flush()
    expect(store.forgets).toBe(0)
    expect(session.snapshot()).toEqual({ status: 'reconnecting' })
    expect(scheduled).toHaveLength(1)
    scheduled[0]!()
    await flush()
    sockets[1]!.open()
    sockets[1]!.receive({
      version: 2,
      type: 'office.challenge',
      binding_id: binding.bindingId,
      challenge: 'challenge_87654321',
      expires_in: 30,
    })
    await flush()
    sockets[1]!.receive({ version: 2, type: 'office.waiting_for_pc' })
    expect(session.snapshot()).toEqual({ status: 'waiting_for_pc' })
  })

  it.each(['invalid_frame', 'unknown_type'])(
    'falls back once from a stored binding when an old Relay returns %s',
    async (code) => {
      const store = new FakeBindingStore(binding)
      const sockets: FakeSocket[] = []
      const scheduled: Array<() => void> = []
      const session = createOfficeRelaySession({
        capabilities: ['agent.v1'],
        bindingStore: store,
        createSocket: () => {
          const socket = new FakeSocket()
          sockets.push(socket)
          return socket
        },
        schedule: (callback) => {
          scheduled.push(callback)
          return scheduled.length
        },
        cancelSchedule: () => undefined,
      })
      void session.connect('word')
      await flush()
      sockets[0]!.open()
      sockets[0]!.receive({ version: 2, type: 'relay.error', code })
      await flush()

      expect(store.forgets).toBe(0)
      expect(scheduled).toHaveLength(0)
      expect(sockets).toHaveLength(2)
      sockets[1]!.open()
      expect(sent(sockets[1]!)).toEqual({
        version: 2,
        type: 'office.create',
        host: 'Word',
        capabilities: ['agent.v1'],
      })
      sockets[1]!.close()
      await flush()
      expect(sockets).toHaveLength(2)
      expect(session.snapshot()).toEqual({ status: 'offline' })
    },
  )

  it('falls back once when an old Relay closes before recognizing stored-binding resume', async () => {
    const store = new FakeBindingStore(binding)
    const sockets: FakeSocket[] = []
    const scheduled: Array<() => void> = []
    const session = createOfficeRelaySession({
      capabilities: ['agent.v1'],
      bindingStore: store,
      createSocket: () => {
        const socket = new FakeSocket()
        sockets.push(socket)
        return socket
      },
      schedule: (callback) => {
        scheduled.push(callback)
        return scheduled.length
      },
      cancelSchedule: () => undefined,
    })
    void session.connect('word')
    await flush()
    sockets[0]!.open()
    sockets[0]!.close()
    await flush()

    expect(scheduled).toHaveLength(0)
    expect(sockets).toHaveLength(2)
    sockets[1]!.open()
    expect(sent(sockets[1]!).type).toBe('office.create')
    sockets[1]!.close()
    await flush()
    expect(sockets).toHaveLength(2)
  })

  it('invalidates synchronously before asynchronous terminal binding deletion', async () => {
    const store = new FakeBindingStore(binding)
    let releaseForget!: () => void
    store.forget = async () => new Promise<void>((resolve) => (releaseForget = resolve))
    const sockets: FakeSocket[] = []
    const scheduled: Array<() => void> = []
    const session = createOfficeRelaySession({
      capabilities: ['agent.v1'],
      bindingStore: store,
      createSocket: () => {
        const socket = new FakeSocket()
        sockets.push(socket)
        return socket
      },
      schedule: (callback) => {
        scheduled.push(callback)
        return scheduled.length
      },
      cancelSchedule: () => undefined,
    })
    void session.connect('word')
    await flush()
    sockets[0]!.open()
    sockets[0]!.receive({ version: 2, type: 'relay.error', code: 'invalid_proof' })
    await flush()

    expect(sockets[0]!.readyState).toBe(3)
    sockets[0]!.close()
    expect(scheduled).toHaveLength(0)
    releaseForget()
    await flush()
    await flush()
    expect(sockets).toHaveLength(2)
    sockets[1]!.open()
    expect(sent(sockets[1]!).type).toBe('office.create')
  })

  it('invalidates synchronously before deleting an unusable signing key', async () => {
    const store = new FakeBindingStore(binding)
    store.sign = async () => {
      throw new Error('binding_key_unusable')
    }
    let releaseForget!: () => void
    store.forget = async () => new Promise<void>((resolve) => (releaseForget = resolve))
    const socket = new FakeSocket()
    const scheduled: Array<() => void> = []
    const session = createOfficeRelaySession({
      capabilities: ['agent.v1'],
      bindingStore: store,
      createSocket: () => socket,
      schedule: (callback) => {
        scheduled.push(callback)
        return scheduled.length
      },
      cancelSchedule: () => undefined,
    })
    void session.connect('word')
    await flush()
    socket.open()
    socket.receive({
      version: 2,
      type: 'office.challenge',
      binding_id: binding.bindingId,
      challenge: 'challenge_12345678',
      expires_in: 30,
    })
    await flush()

    expect(socket.readyState).toBe(3)
    socket.close()
    expect(scheduled).toHaveLength(0)
    releaseForget()
  })

  it('broadcasts manual forget across same-origin host/binding taskpanes', async () => {
    const store = new FakeBindingStore(binding)
    const channel = new FakeInvalidationChannel()
    const sockets = [new FakeSocket(), new FakeSocket()]
    const scheduled: Array<() => void> = []
    const sessions = sockets.map((socket) =>
      createOfficeRelaySession({
        capabilities: ['agent.v1'],
        bindingStore: store,
        bindingInvalidationChannel: channel,
        createSocket: () => socket,
        schedule: (callback) => {
          scheduled.push(callback)
          return scheduled.length
        },
        cancelSchedule: () => undefined,
      }),
    )
    void sessions[0]!.connect('word')
    void sessions[1]!.connect('word')
    await flush()
    sockets.forEach((socket) => socket.open())

    await sessions[0]!.forget()

    expect(sessions[1]!.snapshot()).toEqual({ status: 'offline' })
    sockets[1]!.close()
    expect(scheduled).toHaveLength(0)
  })

  it('ignores forget broadcasts for a different host or binding', async () => {
    const store = new FakeBindingStore(binding)
    const channel = new FakeInvalidationChannel()
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({
      capabilities: ['agent.v1'],
      bindingStore: store,
      bindingInvalidationChannel: channel,
      createSocket: () => socket,
    })
    void session.connect('word')
    await flush()
    socket.open()

    channel.broadcast({ origin: binding.origin, host: 'excel', bindingId: binding.bindingId })
    channel.broadcast({ origin: binding.origin, host: 'word', bindingId: 'binding_other' })

    expect(session.snapshot()).toEqual({ status: 'connecting' })
    expect(socket.readyState).toBe(1)
  })

  it.each(['challenge_expired', 'resume_rate_limited', 'resume_limit', 'peer_unavailable'])(
    'retains the binding and backs off after %s',
    async (code) => {
      const store = new FakeBindingStore(binding)
      const scheduled: Array<{ callback: () => void; delay: number }> = []
      const active = new FakeSocket()
      const retrying = createOfficeRelaySession({
        capabilities: ['agent.v1'],
        bindingStore: store,
        createSocket: () => active,
        schedule: (callback, delay) => {
          scheduled.push({ callback, delay })
          return scheduled.length
        },
        cancelSchedule: vi.fn(),
        random: () => 0.5,
      })
      void retrying.connect('word')
      await flush()
      active.open()
      active.receive({ version: 2, type: 'relay.error', code })
      await flush()
      expect(store.forgets).toBe(0)
      expect(retrying.snapshot()).toEqual({ status: 'reconnecting' })
      expect(scheduled.at(-1)?.delay).toBe(500)
    },
  )

  it('cancels automatic retry on explicit disconnect and deletes only on forget', async () => {
    const store = new FakeBindingStore(binding)
    const socket = new FakeSocket()
    const scheduled: Array<() => void> = []
    const cancelSchedule = vi.fn()
    const session = createOfficeRelaySession({
      capabilities: ['agent.v1'],
      bindingStore: store,
      createSocket: () => socket,
      schedule: (callback) => {
        scheduled.push(callback)
        return scheduled.length
      },
      cancelSchedule,
      random: () => 0.5,
    })
    void session.connect('word')
    await flush()
    socket.open()
    socket.receive({
      version: 2,
      type: 'office.challenge',
      binding_id: binding.bindingId,
      challenge: 'challenge_12345678',
      expires_in: 30,
    })
    await flush()
    socket.receive({ version: 2, type: 'office.waiting_for_pc' })
    socket.close()
    await flush()
    expect(scheduled).toHaveLength(1)

    session.disconnect()
    expect(cancelSchedule).toHaveBeenCalled()
    expect(store.forgets).toBe(0)
    scheduled[0]!()
    await flush()
    expect(session.snapshot()).toEqual({ status: 'offline' })

    await session.forget()
    expect(store.forgets).toBe(1)
  })

  it('caps exponential retry delay at thirty seconds', async () => {
    const store = new FakeBindingStore(binding)
    const sockets: FakeSocket[] = []
    const scheduled: Array<{ callback: () => void; delay: number }> = []
    const session = createOfficeRelaySession({
      capabilities: ['agent.v1'],
      bindingStore: store,
      createSocket: () => {
        const socket = new FakeSocket()
        sockets.push(socket)
        return socket
      },
      schedule: (callback, delay) => {
        scheduled.push({ callback, delay })
        return scheduled.length
      },
      cancelSchedule: () => undefined,
      random: () => 1,
    })
    void session.connect('word')
    await flush()

    for (let attempt = 0; attempt < 9; attempt += 1) {
      sockets[attempt]!.open()
      sockets[attempt]!.receive({
        version: 2,
        type: 'relay.error',
        code: 'resume_rate_limited',
      })
      await flush()
      scheduled[attempt]!.callback()
      await flush()
    }

    expect(scheduled.map(({ delay }) => delay)).toEqual([
      625, 1250, 2500, 5000, 10000, 20000, 30000, 30000, 30000,
    ])
  })

  it('uses ordinary pairing when persistent WebCrypto or IndexedDB storage is unsupported', () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({
      capabilities: ['agent.v1'],
      createSocket: () => socket,
    })
    void session.connect('word')
    socket.open()
    expect(sent(socket)).toEqual({
      version: 2,
      type: 'office.create',
      host: 'Word',
      capabilities: ['agent.v1'],
    })
  })

  it('falls back to ordinary v2 when injected persistent storage is unavailable', async () => {
    const store = new FakeBindingStore()
    store.load = async () => {
      throw new Error('binding_storage_unavailable')
    }
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({
      capabilities: ['agent.v1'],
      bindingStore: store,
      createSocket: () => socket,
    })

    void session.connect('word')
    await flush()
    socket.open()

    expect(sent(socket)).toEqual({
      version: 2,
      type: 'office.create',
      host: 'Word',
      capabilities: ['agent.v1'],
    })
  })

  it('cancels interrupted work and never replays it when a bound session reconnects', async () => {
    const store = new FakeBindingStore(binding)
    const sockets: FakeSocket[] = []
    const scheduled: Array<() => void> = []
    const session = createOfficeRelaySession({
      capabilities: ['agent.v1'],
      bindingStore: store,
      randomUUID: () => 'request_12345678',
      createSocket: () => {
        const socket = new FakeSocket()
        sockets.push(socket)
        return socket
      },
      schedule: (callback) => {
        scheduled.push(callback)
        return scheduled.length
      },
      cancelSchedule: () => undefined,
      random: () => 0.5,
    })
    const connecting = session.connect('word')
    await flush()
    sockets[0]!.open()
    sockets[0]!.receive({
      version: 2,
      type: 'office.challenge',
      binding_id: binding.bindingId,
      challenge: 'challenge_12345678',
      expires_in: 30,
    })
    await flush()
    sockets[0]!.receive({
      version: 2,
      type: 'office.approved',
      session_id: 'session_12345678',
      capability: 'session_capability_12345678',
      expires_in: 1800,
      capabilities: ['agent.v1'],
    })
    await connecting
    const request = session.authenticatedFetch('/v1/office/messages', {
      method: 'POST',
      body: '{}',
    })
    expect(sent(sockets[0]!, 2).type).toBe('office.request')

    sockets[0]!.close()
    await expect(request).rejects.toThrow('relay_disconnected')
    scheduled[0]!()
    await flush()
    sockets[1]!.open()

    expect(sockets[1]!.sent).toHaveLength(1)
    expect(sent(sockets[1]!)).toEqual({
      version: 2,
      type: 'office.resume',
      binding_id: binding.bindingId,
      host: 'Word',
      capabilities: ['agent.v1'],
    })
  })
})
