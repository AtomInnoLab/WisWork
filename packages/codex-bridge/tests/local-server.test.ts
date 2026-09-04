import { request } from 'node:http'
import { describe, expect, it, vi } from 'vitest'
import { startResponsesBridge } from '../src/local-server.js'
import {
  prepareResponsesTurn,
  replayProtocolRecording,
  type ProtocolRecording,
} from '../src/index.js'
import recording from './fixtures/protocol-redacted-max-tokens.json'

function post(url: URL, secret: string, body: string, headers: Record<string, string> = {}) {
  return new Promise<{ status: number; body: string }>((resolve, reject) => {
    const req = request(
      url,
      {
        method: 'POST',
        headers: {
          authorization: `Bearer ${secret}`,
          'content-type': 'application/json',
          'content-length': String(Buffer.byteLength(body)),
          ...headers,
        },
      },
      (res) => {
        const chunks: Buffer[] = []
        res.on('data', (chunk) => chunks.push(Buffer.from(chunk)))
        res.on('end', () =>
          resolve({ status: res.statusCode ?? 0, body: Buffer.concat(chunks).toString() }),
        )
      },
    )
    req.on('error', reject)
    req.end(body)
  })
}

const prepared = () => ({
  messagesRequest: { model: 'openai/gpt-5.6-sol', messages: [], stream: true } as any,
  async *messagesStreamToResponses() {
    yield 'event: response.created\ndata: {"type":"response.created","response":{"id":"r1"}}\n\n'
    yield 'data: [DONE]\n\n'
  },
})

describe('local responses bridge', () => {
  it('captures real translated upstream frames via the fail-open export callback', async () => {
    const captures: ProtocolRecording[] = []
    const outcomes: string[] = []
    const bridge = await startResponsesBridge({
      fetchWithAuth: async () =>
        new Response(
          recording.frames.map((frame) => `data: ${JSON.stringify(frame)}\n\n`).join(''),
          { headers: { 'content-type': 'text/event-stream' } },
        ),
      prepareTurn: prepareResponsesTurn,
      onProtocolRecording: (capture, outcome) => {
        captures.push(capture)
        outcomes.push(outcome)
        throw new Error('ignored observer failure')
      },
    })
    try {
      const response = await post(
        new URL(bridge.responsesUrl),
        bridge.secret,
        JSON.stringify({ model: 'gpt-5.6-sol', input: 'private request' }),
      )
      expect(response.status).toBe(200)
      expect(response.body).toContain('response.incomplete')
      expect(captures).toHaveLength(1)
      expect(outcomes).toEqual(['incomplete'])
      expect((await replayProtocolRecording(captures[0])).events).toContain('response.incomplete')
    } finally {
      await bridge.close()
    }
  })
  it('binds numeric loopback with a random per-process credential and fixed path', async () => {
    const bridge = await startResponsesBridge({
      fetchWithAuth: async () =>
        new Response('', { headers: { 'content-type': 'text/event-stream' } }),
      prepareTurn: prepared,
    })
    try {
      expect(new URL(bridge.baseUrl).hostname).toBe('127.0.0.1')
      expect(Buffer.from(bridge.secret, 'base64url')).toHaveLength(32)
      expect(bridge.responsesUrl).toBe(`${bridge.baseUrl}/v1/responses`)
      expect((await post(new URL('/other', bridge.baseUrl), bridge.secret, '{}')).status).toBe(404)
    } finally {
      await bridge.close()
    }
  })

  it('authenticates before body parsing and never calls upstream for invalid credentials', async () => {
    const fetchWithAuth = vi.fn(async () => new Response(''))
    const prepareTurn = vi.fn(prepared)
    const bridge = await startResponsesBridge({ fetchWithAuth, prepareTurn })
    try {
      const result = await post(new URL(bridge.responsesUrl), 'A'.repeat(43), '{bad')
      expect(result).toMatchObject({ status: 401 })
      expect(fetchWithAuth).not.toHaveBeenCalled()
      expect(prepareTurn).not.toHaveBeenCalled()
      expect(result.body).not.toContain('A'.repeat(10))
      expect((await post(new URL(bridge.responsesUrl), `${bridge.secret}=`, '{}')).status).toBe(401)
    } finally {
      await bridge.close()
    }
  })

  it('applies raw-body and active-turn bounds before upstream', async () => {
    const fetchWithAuth = vi.fn(async () => new Response(''))
    const bridge = await startResponsesBridge({
      fetchWithAuth,
      prepareTurn: prepared,
      maxBodyBytes: 8,
      maxActiveTurns: 1,
    })
    try {
      expect((await post(new URL(bridge.responsesUrl), bridge.secret, '123456789')).status).toBe(
        413,
      )
      expect(fetchWithAuth).not.toHaveBeenCalled()
    } finally {
      await bridge.close()
    }
  })

  it('passes only the prepared fixed request and abort signal to WisUsage', async () => {
    const fetchWithAuth = vi.fn(async (_request, signal) => {
      expect(signal).toBeInstanceOf(AbortSignal)
      return new Response('event: message_stop\ndata: {"type":"message_stop"}\n\n', {
        headers: { 'content-type': 'text/event-stream' },
      })
    })
    const bridge = await startResponsesBridge({ fetchWithAuth, prepareTurn: prepared })
    try {
      const result = await post(
        new URL(bridge.responsesUrl),
        bridge.secret,
        JSON.stringify({ model: 'evil', upstream_url: 'https://evil.test' }),
      )
      expect(result.status).toBe(200)
      expect(fetchWithAuth).toHaveBeenCalledWith(
        expect.objectContaining({ model: 'openai/gpt-5.6-sol' }),
        expect.any(AbortSignal),
      )
      expect(JSON.stringify(fetchWithAuth.mock.calls)).not.toContain('evil.test')
    } finally {
      await bridge.close()
    }
  })

  it('reports only bounded upstream failure stages without response bodies', async () => {
    const diagnostics: string[] = []
    const bridge = await startResponsesBridge({
      fetchWithAuth: async () =>
        new Response('private provider failure', {
          status: 429,
          headers: { 'content-type': 'application/json' },
        }),
      prepareTurn: prepared,
      diagnostics: (code) => diagnostics.push(code),
    })
    try {
      const result = await post(new URL(bridge.responsesUrl), bridge.secret, '{}')
      expect(result.status).toBe(502)
      expect(diagnostics).toEqual(['responses_upstream_started', 'responses_upstream_rate_limited'])
      expect(JSON.stringify(diagnostics)).not.toContain('private')
    } finally {
      await bridge.close()
    }
  })

  it('reports a closed protocol reason without upstream content', async () => {
    const diagnostics: string[] = []
    const onDeterministicFailure = vi.fn()
    const bridge = await startResponsesBridge({
      fetchWithAuth: async () =>
        new Response('data: private\n\n', {
          headers: { 'content-type': 'text/event-stream' },
        }),
      prepareTurn: () => ({
        ...prepared(),
        async *messagesStreamToResponses() {
          yield 'event: response.created\ndata: {"type":"response.created"}\n\n'
          const error = new Error('invalid_messages_sse')
          error.name = 'ProtocolCompatibilityError'
          throw error
        },
      }),
      diagnostics: (code) => diagnostics.push(code),
      onDeterministicFailure,
    })
    try {
      await post(new URL(bridge.responsesUrl), bridge.secret, '{}').catch(() => undefined)
      expect(diagnostics).toEqual([
        'responses_upstream_started',
        'responses_stream_invalid_messages_sse',
      ])
      expect(JSON.stringify(diagnostics)).not.toContain('private')
      expect(onDeterministicFailure).toHaveBeenCalledWith('invalid_messages_sse')
    } finally {
      await bridge.close()
    }
  })
})
