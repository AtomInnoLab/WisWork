import { describe, expect, it, vi } from 'vitest'
import {
  createOfficeRelayClient,
  type OfficeRelayClient,
  type OfficeRelayStatus,
  type RelaySocket,
} from '../src/main/office-relay-client'
import { createOfficeRelayPool } from '../src/main/office-relay-pool'
import type { OfficeRelayBinding } from '../src/main/office-relay-binding-store'

class FakeSocket implements RelaySocket {
  readyState = 0
  sent: string[] = []
  private readonly listeners = new Map<string, Array<(event: any) => void>>()
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
  message(frame: object): void {
    this.emit('message', { data: JSON.stringify(frame) })
  }
  private emit(name: string, event: any): void {
    for (const listener of this.listeners.get(name) ?? []) listener(event)
  }
}

type Child = OfficeRelayClient & {
  emitPending(pairingId: string, hostLabel?: 'Word' | 'Excel' | 'PowerPoint'): void
  emitStatus(status: OfficeRelayStatus): void
  emitBinding(binding: OfficeRelayBinding): Promise<void>
}

const binding = (index: number): OfficeRelayBinding => ({
  bindingId: `binding_${index}_12345678`,
  accountId: 'account-one',
  host: ['Word', 'Excel', 'PowerPoint'][index % 3] as OfficeRelayBinding['host'],
  origin: 'https://office.8-216-134-194.sslip.io',
  capabilities: ['agent.v1'],
  createdAt: index + 1,
})

function harness(maxClients = 12) {
  const children: Child[] = []
  const pending = vi.fn()
  const expired = vi.fn()
  const remembered = vi.fn()
  const statuses: OfficeRelayStatus[] = []
  const pool = createOfficeRelayPool({
    maxClients,
    onPending: pending,
    onPendingExpired: expired,
    onStatus: (status) => statuses.push(status),
    onBinding: remembered,
    createClient(events) {
      let status: OfficeRelayStatus = 'disconnected'
      let entries: ReturnType<OfficeRelayClient['listPending']> = []
      const child: Child = {
        claim: vi.fn(async () => {
          events.onStatus('connecting')
        }),
        resume: vi.fn(async () => {
          events.onStatus('connecting')
        }),
        revokeBinding: vi.fn(async () => undefined),
        approve: vi.fn(async (id) => entries.some((entry) => entry.pairingId === id)),
        reject: vi.fn((id) => entries.some((entry) => entry.pairingId === id)),
        listPending: () => entries,
        status: () => status,
        revoke: vi.fn((reason = 'revoked') =>
          events.onStatus(`disconnected:${reason}` as OfficeRelayStatus),
        ),
        emitPending(pairingId, hostLabel = 'Word') {
          entries = [
            {
              pairingId,
              hostLabel,
              origin: 'https://office.8-216-134-194.sslip.io',
              verificationCode: '123456',
            },
          ]
          events.onPending(entries[0]!)
        },
        emitStatus(next) {
          status = next
          if (next === 'paired') entries = []
          if (next.startsWith('disconnected:')) {
            const expiredIds = entries.map((entry) => entry.pairingId)
            entries = []
            expiredIds.forEach(events.onPendingExpired)
          }
          events.onStatus(next)
        },
        emitBinding(value) {
          return events.onBinding(value)
        },
      }
      children.push(child)
      return child
    },
  })
  return { pool, children, pending, expired, statuses, remembered }
}

describe('Office relay pool', () => {
  it('keeps three independently claimed Office sessions alive', async () => {
    const { pool, children } = harness()
    await Promise.all([pool.claim('111111'), pool.claim('222222'), pool.claim('333333')])
    expect(children).toHaveLength(3)
    expect(children.map((child) => child.claim)).toEqual([
      expect.any(Function),
      expect.any(Function),
      expect.any(Function),
    ])
    expect(
      children.map((child) => (child.claim as ReturnType<typeof vi.fn>).mock.calls[0]?.[0]),
    ).toEqual(['111111', '222222', '333333'])
    expect(
      children.every((child) => (child.revoke as ReturnType<typeof vi.fn>).mock.calls.length === 0),
    ).toBe(true)
  })

  it('isolates identical request ids and cancellation across three real child clients', async () => {
    const sockets: FakeSocket[] = []
    const signals: AbortSignal[] = []
    const pool = createOfficeRelayPool({
      onPending() {},
      createClient: (events) =>
        createOfficeRelayClient({
          endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
          connect: () => {
            const socket = new FakeSocket()
            sockets.push(socket)
            return socket
          },
          getValidAccountStatus: async () => ({ loggedIn: true }),
          getAccessToken: async () => 'access-token',
          proxy: async ({ signal }) => {
            signals.push(signal)
            return {
              status: 200,
              body: (async function* () {
                await new Promise<void>((resolve) =>
                  signal.addEventListener('abort', () => resolve(), { once: true }),
                )
                yield new Uint8Array()
              })(),
            }
          },
          onPending: events.onPending,
          onPendingExpired: events.onPendingExpired,
          onStatus: events.onStatus,
        }),
    })

    for (let index = 0; index < 3; index += 1) {
      const code = `${index + 1}`.repeat(6)
      const claim = pool.claim(code)
      await vi.waitFor(() => expect(sockets).toHaveLength(index + 1))
      sockets[index]!.open()
      await claim
      const pairingId = `pairing_${index}_12345678`
      sockets[index]!.message({
        version: 1,
        type: 'pc.claimed',
        pairing_id: pairingId,
        host: ['Word', 'Excel', 'PowerPoint'][index],
        origin: 'https://office.8-216-134-194.sslip.io',
        verification_code: code,
        expires_in: 120,
      })
      await pool.approve(pairingId)
      sockets[index]!.message({
        version: 1,
        type: 'pc.approved',
        session_id: `session_${index}_12345678`,
        capability: `capability_${index}_12345678`,
        expires_in: 1800,
      })
      expect(pool.status()).toBe('paired')
    }

    sockets.forEach((socket, index) =>
      socket.message({
        version: 1,
        type: 'relay.request',
        session_id: `session_${index}_12345678`,
        request_id: 'same_request_12345678',
        body: { index },
      }),
    )
    await vi.waitFor(() => expect(signals).toHaveLength(3))
    sockets[0]!.message({
      version: 1,
      type: 'relay.cancel',
      session_id: 'session_0_12345678',
      request_id: 'same_request_12345678',
    })
    expect(signals.map((signal) => signal.aborted)).toEqual([true, false, false])
    expect(pool.status()).toBe('paired')
    pool.revoke('shutdown')
  })

  it('flattens pending requests and routes approve and reject by exact pairing id', async () => {
    const { pool, children } = harness()
    await Promise.all([pool.claim('111111'), pool.claim('222222'), pool.claim('333333')])
    children[0]!.emitPending('pairing_word_123', 'Word')
    children[1]!.emitPending('pairing_excel_123', 'Excel')
    children[2]!.emitPending('pairing_powerpoint_123', 'PowerPoint')

    expect(pool.listPending().map((entry) => entry.hostLabel)).toEqual([
      'Word',
      'Excel',
      'PowerPoint',
    ])
    await expect(pool.approve('pairing_excel_123')).resolves.toBe(true)
    expect(children[0]!.approve).not.toHaveBeenCalled()
    expect(children[1]!.approve).toHaveBeenCalledWith('pairing_excel_123')
    expect(pool.reject('pairing_word_123')).toBe(true)
    expect(children[0]!.reject).toHaveBeenCalledWith('pairing_word_123')
    expect(pool.reject('unknown_pairing')).toBe(false)
  })

  it('keeps approval ownership until the child confirms or expires without reviving the snapshot', async () => {
    const { pool, children, expired } = harness()
    await pool.claim('111111')
    children[0]!.emitPending('pairing_word_123', 'Word')

    await expect(pool.approve('pairing_word_123')).resolves.toBe(true)
    expect(pool.listPending()).toEqual([])
    await expect(pool.approve('pairing_word_123')).resolves.toBe(false)
    expect(pool.reject('pairing_word_123')).toBe(false)

    children[0]!.emitStatus('disconnected:pairing_expired')
    expect(expired).toHaveBeenCalledOnce()
    expect(expired).toHaveBeenCalledWith('pairing_word_123')
    expect(pool.listPending()).toEqual([])
  })

  it('clears approval ownership only after pc.approved transitions the child to paired', async () => {
    const { pool, children, expired } = harness()
    await pool.claim('111111')
    children[0]!.emitPending('pairing_word_123', 'Word')
    await pool.approve('pairing_word_123')

    children[0]!.emitStatus('paired')
    children[0]!.emitStatus('disconnected:relay_closed')
    expect(expired).not.toHaveBeenCalled()
    expect(pool.listPending()).toEqual([])
  })

  it('isolates a child disconnect and preserves aggregate paired status', async () => {
    const { pool, children } = harness()
    await Promise.all([pool.claim('111111'), pool.claim('222222'), pool.claim('333333')])
    children.forEach((child) => child.emitStatus('paired'))
    children[0]!.emitStatus('disconnected:relay_closed')
    expect(pool.status()).toBe('paired')
    expect(children[1]!.revoke).not.toHaveBeenCalled()
    expect(children[2]!.revoke).not.toHaveBeenCalled()
  })

  it('reserves capacity before concurrent claims and releases failed claims', async () => {
    const { pool, children } = harness(2)
    const first = pool.claim('111111')
    const second = pool.claim('222222')
    await expect(pool.claim('333333')).rejects.toThrow('relay_capacity_exceeded')
    await Promise.all([first, second])

    ;(children[0]!.claim as ReturnType<typeof vi.fn>).mockRejectedValueOnce(new Error('boom'))
    children[0]!.emitStatus('disconnected:network_error')
    await expect(pool.claim('333333')).resolves.toBeUndefined()
  })

  it('releases a reservation when child claim setup fails', async () => {
    let attempts = 0
    const pool = createOfficeRelayPool({
      maxClients: 1,
      onPending() {},
      createClient: () => ({
        claim: async () => {
          attempts += 1
          if (attempts === 1) throw new Error('relay_connection_failed')
        },
        resume: async () => undefined,
        revokeBinding: async () => undefined,
        approve: async () => false,
        reject: () => false,
        listPending: () => [],
        status: () => 'disconnected',
        revoke() {},
      }),
    })
    await expect(pool.claim('111111')).rejects.toThrow('relay_connection_failed')
    await expect(pool.claim('222222')).resolves.toBeUndefined()
  })

  it('does not let a failed new claim hide an existing paired session', async () => {
    const { pool, children } = harness()
    await pool.claim('111111')
    children[0]!.emitStatus('paired')
    await pool.claim('222222')
    children[1]!.emitStatus('disconnected:network_error')
    expect(pool.status()).toBe('paired')
    children[0]!.emitStatus('disconnected:relay_closed')
    expect(pool.status()).toBe('disconnected:relay_closed')
  })

  it('globally revokes every child on logout, shutdown, or terminal auth loss', async () => {
    for (const reason of ['logout', 'shutdown', 'auth_required']) {
      const { pool, children } = harness()
      await Promise.all([pool.claim('111111'), pool.claim('222222'), pool.claim('333333')])
      pool.revoke(reason)
      expect(
        children.every(
          (child) => (child.revoke as ReturnType<typeof vi.fn>).mock.calls[0]?.[0] === reason,
        ),
      ).toBe(true)
    }
  })

  it('turns a terminal auth failure in one child into a global revoke', async () => {
    const { pool, children } = harness()
    await Promise.all([pool.claim('111111'), pool.claim('222222'), pool.claim('333333')])
    children[0]!.emitStatus('disconnected:auth_required')
    expect(children[1]!.revoke).toHaveBeenCalledWith('auth_required')
    expect(children[2]!.revoke).toHaveBeenCalledWith('auth_required')
  })

  it('blocks late binding capture and new enrollment after lifecycle suspension', async () => {
    const { pool, children } = harness()
    ;(pool as OfficeRelayPool & { suspend(reason: string): void }).suspend('logout')
    await expect(pool.claim('111111')).rejects.toThrow('relay_suspended')
    expect(children).toHaveLength(0)

    const active = harness()
    await active.pool.claim('111111')
    active.pool.suspend('account_switch')
    await expect(active.children[0]!.emitBinding(binding(0))).rejects.toThrow(
      'binding_lifecycle_suspended',
    )
    expect(active.remembered).not.toHaveBeenCalled()
    expect(active.children[0]!.revoke).toHaveBeenCalledWith('account_switch')
  })

  it('publishes the persistent lifecycle diagnostic when startup suspension disables the pool', () => {
    const { pool } = harness()
    pool.suspend('binding_lifecycle')
    expect(pool.status()).toBe('error:binding_lifecycle')
  })

  it('keeps twelve durable bindings in independent resume slots', async () => {
    const { pool, children } = harness()
    await Promise.all(Array.from({ length: 12 }, (_, index) => pool.resume(binding(index))))
    expect(children).toHaveLength(12)
    expect(
      children.map((child) => (child.resume as ReturnType<typeof vi.fn>).mock.calls[0]?.[0]),
    ).toEqual(Array.from({ length: 12 }, (_, index) => binding(index)))
    await expect(pool.resume(binding(12))).rejects.toThrow('relay_capacity_exceeded')
  })

  it('replenishes one waiting resume so the same binding can create independent concurrent sessions', async () => {
    const { pool, children } = harness()
    await pool.resume(binding(0))
    expect(children).toHaveLength(1)

    children[0]!.emitStatus('paired')
    await vi.waitFor(() => expect(children).toHaveLength(2))
    expect(children[1]!.resume).toHaveBeenCalledWith(binding(0))

    children[1]!.emitStatus('paired')
    await vi.waitFor(() => expect(children).toHaveLength(3))
    expect(children[0]!.revoke).not.toHaveBeenCalled()
    expect(children[1]!.revoke).not.toHaveBeenCalled()
  })

  it('retries each disconnected binding with bounded exponential backoff and cancels on shutdown', async () => {
    vi.useFakeTimers()
    const children: Child[] = []
    const delays: number[] = []
    const pool = createOfficeRelayPool({
      onPending() {},
      random: () => 0.5,
      baseReconnectMs: 100,
      maxReconnectMs: 250,
      setTimer(callback, delay) {
        delays.push(delay)
        return setTimeout(callback, delay)
      },
      clearTimer: clearTimeout,
      createClient(events) {
        let status: OfficeRelayStatus = 'disconnected'
        const child: Child = {
          claim: vi.fn(async () => undefined),
          resume: vi.fn(async () => events.onStatus('connecting')),
          revokeBinding: vi.fn(async () => undefined),
          approve: vi.fn(async () => false),
          reject: vi.fn(() => false),
          listPending: () => [],
          status: () => status,
          revoke: vi.fn(),
          emitPending() {},
          emitStatus(next) {
            status = next
            events.onStatus(next)
          },
        }
        children.push(child)
        return child
      },
    })
    await pool.resume(binding(0))
    children[0]!.emitStatus('disconnected:network_error')
    expect(delays).toEqual([100])
    await vi.advanceTimersByTimeAsync(100)
    expect(children[0]!.resume).toHaveBeenCalledTimes(2)
    children[0]!.emitStatus('disconnected:relay_closed')
    expect(delays).toEqual([100, 200])
    await vi.advanceTimersByTimeAsync(200)
    children[0]!.emitStatus('disconnected:resume_rate_limited')
    expect(delays).toEqual([100, 200, 250])

    pool.revoke('shutdown')
    await vi.advanceTimersByTimeAsync(1_000)
    expect(children[0]!.resume).toHaveBeenCalledTimes(3)
    vi.useRealTimers()
  })

  it('removes a terminally invalid binding while keeping other resumed bindings', async () => {
    const invalidated = vi.fn()
    const children: Child[] = []
    const pool = createOfficeRelayPool({
      onPending() {},
      onBindingInvalidated: invalidated,
      createClient(events) {
        const child: Child = {
          claim: vi.fn(async () => undefined),
          resume: vi.fn(async () => undefined),
          revokeBinding: vi.fn(async () => undefined),
          approve: vi.fn(async () => false),
          reject: vi.fn(() => false),
          listPending: () => [],
          status: () => 'disconnected',
          revoke: vi.fn(),
          emitPending() {},
          emitStatus: events.onStatus,
        }
        children.push(child)
        return child
      },
    })
    await Promise.all([pool.resume(binding(0)), pool.resume(binding(1))])
    children[0]!.emitStatus('disconnected:binding_unavailable')
    expect(invalidated).toHaveBeenCalledWith(binding(0))
    children[1]!.emitStatus('paired')
    expect(pool.status()).toBe('paired')
  })

  it('keeps a revocation slot reserved across the child new_revocation transition', async () => {
    let release!: () => void
    const blocked = new Promise<void>((resolve) => {
      release = resolve
    })
    const pool = createOfficeRelayPool({
      maxClients: 1,
      onPending() {},
      createClient(events) {
        return {
          claim: vi.fn(async () => undefined),
          resume: vi.fn(async () => events.onStatus('connecting')),
          revokeBinding: vi.fn(async () => {
            events.onStatus('disconnected:new_revocation')
            await blocked
          }),
          approve: vi.fn(async () => false),
          reject: vi.fn(() => false),
          listPending: vi.fn(() => []),
          status: vi.fn(() => 'connecting' as const),
          revoke: vi.fn(),
        }
      },
    })
    await pool.resume(binding(0))
    const revoking = pool.revokeBinding(binding(0).bindingId)
    await expect(pool.claim('111111')).rejects.toThrow('relay_capacity_exceeded')
    release()
    await revoking
  })

  it('publishes binding_not_remembered after real-client persistence failure revocation', async () => {
    const sockets: FakeSocket[] = []
    const statuses: OfficeRelayStatus[] = []
    const pool = createOfficeRelayPool({
      onPending() {},
      onStatus: (status) => statuses.push(status),
      onBinding: async () => {
        throw new Error('disk_failure')
      },
      createClient: (events) =>
        createOfficeRelayClient({
          endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
          connect: () => {
            const socket = new FakeSocket()
            sockets.push(socket)
            return socket
          },
          getValidAccountStatus: async () => ({ loggedIn: true, userId: 'account-one' }),
          getAccessToken: async () => 'token',
          proxy: async () => ({ status: 200, body: new Uint8Array() }),
          persistentPairing: true,
          onPending: events.onPending,
          onPendingExpired: events.onPendingExpired,
          onStatus: events.onStatus,
          onBinding: events.onBinding,
          onBindingInvalidated: events.onBindingInvalidated,
        }),
    })
    const claim = pool.claim('123456')
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
    await vi.waitFor(() => expect(pool.listPending()).toHaveLength(1))
    await pool.approve('pairing_12345678')
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
    sockets[1]!.open()
    await vi.waitFor(() => expect(sockets[1]!.sent).toHaveLength(1))
    sockets[1]!.message({
      version: 2,
      type: 'pc.binding_revoked',
      binding_id: 'binding_word_12345678',
    })
    await vi.waitFor(() => expect(pool.status()).toBe('disconnected:binding_not_remembered'))
    expect(statuses).toContain('disconnected:binding_not_remembered')
  })
})
