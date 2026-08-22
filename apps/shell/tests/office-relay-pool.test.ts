import { describe, expect, it, vi } from 'vitest'
import {
  createOfficeRelayClient,
  type OfficeRelayClient,
  type OfficeRelayStatus,
  type RelaySocket,
} from '../src/main/office-relay-client'
import { createOfficeRelayPool } from '../src/main/office-relay-pool'

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
}

function harness(maxClients = 12) {
  const children: Child[] = []
  const pending = vi.fn()
  const expired = vi.fn()
  const statuses: OfficeRelayStatus[] = []
  const pool = createOfficeRelayPool({
    maxClients,
    onPending: pending,
    onPendingExpired: expired,
    onStatus: (status) => statuses.push(status),
    createClient(events) {
      let status: OfficeRelayStatus = 'disconnected'
      let entries: ReturnType<OfficeRelayClient['listPending']> = []
      const child: Child = {
        claim: vi.fn(async () => {
          events.onStatus('connecting')
        }),
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
          if (next.startsWith('disconnected:')) entries = []
          events.onStatus(next)
        },
      }
      children.push(child)
      return child
    },
  })
  return { pool, children, pending, expired, statuses }
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
})
