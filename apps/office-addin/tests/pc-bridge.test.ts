import { describe, expect, it, vi } from 'vitest'
import { createPcBridgeSession } from '../src/pc-bridge/session.js'

const endpoint = 'http://127.0.0.1:43127'

describe('PC bridge session', () => {
  it('creates a host-bound pairing and polls until approved', async () => {
    const fetch = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({ pairing_id: 'pair', polling_secret: 'poll', expires_in: 120 }),
          { status: 202 },
        ),
      )
      .mockResolvedValueOnce(new Response(JSON.stringify({ status: 'pending' })))
      .mockResolvedValueOnce(
        new Response(JSON.stringify({ status: 'approved', capability: 'cap', expires_in: 900 })),
      )
    const session = createPcBridgeSession({ endpoint, fetch, delay: async () => undefined })
    const states: string[] = []
    session.subscribe(() => states.push(session.snapshot().status))

    await session.connect('word')

    expect(JSON.parse(fetch.mock.calls[0]![1].body)).toEqual({ host_label: 'Word' })
    expect(new Headers(fetch.mock.calls[1]![1].headers).get('authorization')).toBe('Pairing poll')
    expect(states).toContain('pending')
    expect(session.snapshot().status).toBe('connected')
  })

  it.each(['rejected', 'expired'] as const)(
    'surfaces %s without retaining a capability',
    async (status) => {
      const fetch = vi
        .fn()
        .mockResolvedValueOnce(
          new Response(
            JSON.stringify({ pairing_id: 'pair', polling_secret: 'poll', expires_in: 120 }),
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
          JSON.stringify({ pairing_id: 'pair', polling_secret: 'poll', expires_in: 120 }),
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
        new Response(JSON.stringify({ pairing_id: 'p', polling_secret: 's', expires_in: 1 }), {
          status: 202,
        }),
      )
      .mockResolvedValueOnce(new Response('', { status: 401 }))
    const signedOut = createPcBridgeSession({ endpoint, fetch, delay: async () => undefined })
    await signedOut.connect('word')
    expect(signedOut.snapshot()).toEqual({ status: 'signed_out' })
  })
})
