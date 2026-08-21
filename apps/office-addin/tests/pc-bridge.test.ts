import { describe, expect, it, vi } from 'vitest'
import { createPcBridgeSession } from '../src/pc-bridge/session.js'

const endpoint = 'http://127.0.0.1:43127'

describe('PC bridge session', () => {
  it('does not resurrect an approved capability after disconnect during JSON parsing', async () => {
    let release!: (response: Response) => void
    const delayedPoll = new Promise<Response>((resolve) => (release = resolve))
    const fetch = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            pairing_id: 'pair',
            polling_secret: 'poll',
            verification_code: '123456',
            expires_in: 120,
          }),
          { status: 202 },
        ),
      )
      .mockReturnValueOnce(delayedPoll)
    const session = createPcBridgeSession({ endpoint, fetch })
    const connecting = session.connect('word')
    await vi.waitFor(() => expect(session.snapshot().status).toBe('pending'))
    session.disconnect()
    release(
      new Response(JSON.stringify({ status: 'approved', capability: 'stale', expires_in: 900 })),
    )
    await connecting
    expect(session.snapshot().status).toBe('offline')
    await expect(session.authenticatedFetch('/v1/office/messages', {})).rejects.toThrow(
      'bridge_disconnected',
    )
  })

  it('does not let an old capability 401 clear a reconnected session', async () => {
    let releaseOld!: (response: Response) => void
    const fetch = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            pairing_id: 'p1',
            polling_secret: 's1',
            verification_code: '123456',
            expires_in: 120,
          }),
          {
            status: 202,
          },
        ),
      )
      .mockResolvedValueOnce(
        new Response(JSON.stringify({ status: 'approved', capability: 'cap1', expires_in: 900 })),
      )
      .mockReturnValueOnce(new Promise<Response>((resolve) => (releaseOld = resolve)))
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            pairing_id: 'p2',
            polling_secret: 's2',
            verification_code: '234567',
            expires_in: 120,
          }),
          {
            status: 202,
          },
        ),
      )
      .mockResolvedValueOnce(
        new Response(JSON.stringify({ status: 'approved', capability: 'cap2', expires_in: 900 })),
      )
      .mockResolvedValueOnce(new Response('', { status: 200 }))
    const session = createPcBridgeSession({ endpoint, fetch })
    await session.connect('word')
    const oldRequest = session.authenticatedFetch('/v1/office/messages', {})
    session.disconnect()
    await session.connect('word')
    releaseOld(new Response('', { status: 401 }))
    await oldRequest
    expect(session.snapshot().status).toBe('connected')
    await session.authenticatedFetch('/v1/office/messages', {})
    expect(new Headers(fetch.mock.calls[5]![1].headers).get('authorization')).toBe('Bridge cap2')
  })

  it.each(['create', 'poll'] as const)(
    'bounds a hung %s request by the connection deadline',
    async (stage) => {
      vi.useFakeTimers()
      const fetch =
        stage === 'create'
          ? vi.fn(() => new Promise<Response>(() => undefined))
          : vi
              .fn()
              .mockResolvedValueOnce(
                new Response(
                  JSON.stringify({
                    pairing_id: 'p',
                    polling_secret: 's',
                    verification_code: '123456',
                    expires_in: 120,
                  }),
                  { status: 202 },
                ),
              )
              .mockReturnValueOnce(new Promise<Response>(() => undefined))
      const session = createPcBridgeSession({ endpoint, fetch })
      const connecting = session.connect('word')
      await vi.advanceTimersByTimeAsync(120_000)
      await connecting
      expect(session.snapshot().status).toBe(stage === 'poll' ? 'expired' : 'offline')
      vi.useRealTimers()
    },
  )

  it('bounds a response body that never finishes', async () => {
    vi.useFakeTimers()
    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new TextEncoder().encode('{'))
      },
    })
    const session = createPcBridgeSession({
      endpoint,
      fetch: vi.fn().mockResolvedValue(new Response(body, { status: 202 })),
    })
    const connecting = session.connect('word')
    await vi.advanceTimersByTimeAsync(120_000)
    await connecting
    expect(session.snapshot().status).toBe('offline')
    vi.useRealTimers()
  })

  it('rejects oversized and invalid pairing protocol values', async () => {
    const invalidBodies = [
      'x'.repeat(5000),
      JSON.stringify({ pairing_id: '', polling_secret: 's', expires_in: 120 }),
      JSON.stringify({ pairing_id: 'p!', polling_secret: 's', expires_in: 120 }),
      JSON.stringify({
        pairing_id: 'p',
        polling_secret: 's',
        expires_in: Number.POSITIVE_INFINITY,
      }),
      JSON.stringify({ pairing_id: 'p', polling_secret: 's', expires_in: -1 }),
    ]
    for (const body of invalidBodies) {
      const session = createPcBridgeSession({
        endpoint,
        fetch: vi.fn().mockResolvedValue(new Response(body, { status: 202 })),
      })
      await session.connect('word')
      expect(session.snapshot().status).toBe('offline')
    }
  })

  it('rejects invalid or oversized approval protocol values', async () => {
    const invalidBodies = [
      'x'.repeat(5000),
      JSON.stringify({ status: 'approved', capability: '', expires_in: 900 }),
      JSON.stringify({ status: 'approved', capability: 'bad!', expires_in: 900 }),
      JSON.stringify({ status: 'approved', capability: 'c', expires_in: 0 }),
      JSON.stringify({ status: 'approved', capability: 'c', expires_in: 3601 }),
    ]
    for (const body of invalidBodies) {
      const fetch = vi
        .fn()
        .mockResolvedValueOnce(
          new Response(
            JSON.stringify({
              pairing_id: 'p',
              polling_secret: 's',
              verification_code: '123456',
              expires_in: 120,
            }),
            {
              status: 202,
            },
          ),
        )
        .mockResolvedValueOnce(new Response(body))
      const session = createPcBridgeSession({ endpoint, fetch, delay: async () => undefined })
      await session.connect('word')
      expect(session.snapshot().status).toBe('offline')
    }
  })
  it('creates a host-bound pairing and polls until approved', async () => {
    const fetch = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            pairing_id: 'pair',
            polling_secret: 'poll',
            verification_code: '123456',
            expires_in: 120,
          }),
          { status: 202 },
        ),
      )
      .mockResolvedValueOnce(new Response(JSON.stringify({ status: 'pending' })))
      .mockResolvedValueOnce(
        new Response(JSON.stringify({ status: 'approved', capability: 'cap', expires_in: 900 })),
      )
    const session = createPcBridgeSession({ endpoint, fetch, delay: async () => undefined })
    const states: string[] = []
    const snapshots: unknown[] = []
    session.subscribe(() => {
      states.push(session.snapshot().status)
      snapshots.push(session.snapshot())
    })

    await session.connect('word')

    expect(JSON.parse(fetch.mock.calls[0]![1].body)).toEqual({ host_label: 'Word' })
    expect(new Headers(fetch.mock.calls[1]![1].headers).get('authorization')).toBe('Pairing poll')
    expect(states).toContain('pending')
    expect(snapshots).toContainEqual({ status: 'pending', verificationCode: '123456' })
    expect(session.snapshot().status).toBe('connected')
  })

  it.each(['rejected', 'expired'] as const)(
    'surfaces %s without retaining a capability',
    async (status) => {
      const fetch = vi
        .fn()
        .mockResolvedValueOnce(
          new Response(
            JSON.stringify({
              pairing_id: 'pair',
              polling_secret: 'poll',
              verification_code: '123456',
              expires_in: 120,
            }),
            { status: 202 },
          ),
        )
        .mockResolvedValueOnce(new Response(JSON.stringify({ status })))
      const session = createPcBridgeSession({ endpoint, fetch, delay: async () => undefined })
      await session.connect('excel')
      expect(session.snapshot()).toEqual({ status })
      await expect(session.authenticatedFetch('/v1/office/messages', {})).rejects.toThrow(
        'bridge_disconnected',
      )
    },
  )

  it('uses an in-memory Bridge capability and disconnect clears it', async () => {
    const fetch = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            pairing_id: 'pair',
            polling_secret: 'poll',
            verification_code: '123456',
            expires_in: 120,
          }),
          { status: 202 },
        ),
      )
      .mockResolvedValueOnce(
        new Response(JSON.stringify({ status: 'approved', capability: 'cap', expires_in: 900 })),
      )
      .mockResolvedValueOnce(new Response('', { status: 200 }))
    const session = createPcBridgeSession({ endpoint, fetch, delay: async () => undefined })
    await session.connect('powerpoint')
    await session.authenticatedFetch('/v1/office/messages', { method: 'POST' })
    expect(new Headers(fetch.mock.calls[2]![1].headers).get('authorization')).toBe('Bridge cap')
    session.disconnect()
    await expect(session.authenticatedFetch('/v1/office/messages', {})).rejects.toThrow(
      'bridge_disconnected',
    )
    expect(JSON.stringify(session.snapshot())).not.toContain('cap')
  })

  it('maps network failure to offline and 401 to signed_out', async () => {
    const offline = createPcBridgeSession({
      endpoint,
      fetch: vi.fn().mockRejectedValue(new Error('secret')),
    })
    await offline.connect('word')
    expect(offline.snapshot()).toEqual({ status: 'offline' })

    const fetch = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            pairing_id: 'p',
            polling_secret: 's',
            verification_code: '123456',
            expires_in: 1,
          }),
          {
            status: 202,
          },
        ),
      )
      .mockResolvedValueOnce(new Response('', { status: 401 }))
    const signedOut = createPcBridgeSession({ endpoint, fetch, delay: async () => undefined })
    await signedOut.connect('word')
    expect(signedOut.snapshot()).toEqual({ status: 'signed_out' })
  })
})
