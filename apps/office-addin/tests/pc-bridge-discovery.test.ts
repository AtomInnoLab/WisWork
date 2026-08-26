import { describe, expect, it, vi } from 'vitest'
import { createPcBridgeSession, discoverPcBridgeEndpoint } from '../src/pc-bridge/session.js'

const endpoints = [43127, 43120, 43121, 43122, 43123, 43124, 43125, 43126, 43128].map(
  (port) => `http://127.0.0.1:${port}`,
)
const health = () =>
  new Response(JSON.stringify({ service: 'wiswork-office-bridge', version: 1 }), { status: 200 })

describe('PC bridge discovery', () => {
  it('selects the preferred endpoint deterministically and cancels slower losers', async () => {
    let loserAborted = false
    const fetch = vi.fn((url: string | URL | Request, init?: RequestInit) => {
      if (String(url).includes(':43127/')) return Promise.resolve(health())
      return new Promise<Response>((_resolve, reject) => {
        init?.signal?.addEventListener('abort', () => {
          loserAborted = true
          reject(new DOMException('Aborted', 'AbortError'))
        })
      })
    })
    await expect(
      discoverPcBridgeEndpoint(endpoints, { fetch, probeTimeoutMs: 50, batchSize: 8 }),
    ).resolves.toBe(endpoints[0])
    expect(loserAborted).toBe(true)
  })

  it('falls back in configured order, ignoring malformed and hostile health responses', async () => {
    const fetch = vi.fn((url: string | URL | Request) => {
      const value = String(url)
      if (value.includes(':43127/')) return Promise.resolve(new Response('{', { status: 200 }))
      if (value.includes(':43120/'))
        return Promise.resolve(
          new Response(JSON.stringify({ service: 'evil', version: 1 }), { status: 200 }),
        )
      if (value.includes(':43121/'))
        return Promise.resolve(
          new Response(
            JSON.stringify({ service: 'wiswork-office-bridge', version: 1, token: 'hostile' }),
            { status: 200 },
          ),
        )
      if (value.includes(':43122/')) return Promise.resolve(health())
      return Promise.reject(new Error('offline'))
    })
    await expect(discoverPcBridgeEndpoint(endpoints, { fetch, probeTimeoutMs: 20 })).resolves.toBe(
      endpoints[3],
    )
  })

  it('creates exactly one pairing on the first healthy fallback endpoint', async () => {
    const candidates = endpoints.slice(0, 3)
    const fetch = vi.fn((url: string | URL | Request, init?: RequestInit) => {
      const value = String(url)
      if (value.endsWith('/v1/office/health')) {
        if (value.includes(':43120/')) return Promise.resolve(health())
        return Promise.reject(new Error('occupied or offline'))
      }
      if (init?.method === 'POST') return Promise.resolve(new Response('', { status: 401 }))
      return Promise.reject(new Error('unexpected request'))
    })
    const session = createPcBridgeSession({ endpoints: candidates, fetch, probeTimeoutMs: 20 })
    await session.connect('word')
    const pairingRequests = fetch.mock.calls.filter(
      ([url, init]) => String(url).endsWith('/v1/office/pairings') && init?.method === 'POST',
    )
    expect(pairingRequests).toHaveLength(1)
    expect(String(pairingRequests[0]![0])).toBe(`${endpoints[1]}/v1/office/pairings`)
    expect(session.snapshot()).toEqual({ status: 'signed_out' })
  })

  it('bounds offline discovery latency across batches', async () => {
    vi.useFakeTimers()
    const fetch = vi.fn(() => new Promise<Response>(() => undefined))
    const discovering = discoverPcBridgeEndpoint(endpoints, {
      fetch,
      probeTimeoutMs: 100,
      batchSize: 8,
    })
    await vi.advanceTimersByTimeAsync(200)
    await expect(discovering).resolves.toBeUndefined()
    expect(fetch).toHaveBeenCalledTimes(endpoints.length)
    vi.useRealTimers()
  })

  it('does not create a pairing from stale discovery after disconnect', async () => {
    let resolveHealth!: (response: Response) => void
    const fetch = vi.fn((url: string | URL | Request) => {
      if (String(url).endsWith('/v1/office/health'))
        return new Promise<Response>((resolve) => (resolveHealth = resolve))
      throw new Error('pairing must not be created')
    })
    const session = createPcBridgeSession({
      endpoints: [endpoints[0]!],
      fetch,
      probeTimeoutMs: 500,
    })
    const connecting = session.connect('word')
    await vi.waitFor(() => expect(fetch).toHaveBeenCalledTimes(1))
    session.disconnect()
    resolveHealth(health())
    await connecting
    expect(fetch).toHaveBeenCalledTimes(1)
    expect(session.snapshot()).toEqual({ status: 'offline' })
  })

  it('publishes connecting immediately while bridge discovery is in flight', async () => {
    let resolveHealth!: (response: Response) => void
    const fetch = vi.fn(() => new Promise<Response>((resolve) => (resolveHealth = resolve)))
    const session = createPcBridgeSession({
      endpoints: [endpoints[0]!],
      fetch,
      probeTimeoutMs: 500,
    })

    const connecting = session.connect('word')
    expect(session.snapshot()).toEqual({ status: 'connecting' })
    resolveHealth(new Response('', { status: 503 }))
    await connecting
    expect(session.snapshot()).toEqual({ status: 'offline' })
  })
})
