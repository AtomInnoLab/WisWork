import { request } from 'node:http'
import { describe, expect, it, vi } from 'vitest'
import { startResponsesBridge } from '../src/local-server.js'
import {
  prepareResponsesTurn,
  replayProtocolRecording,
  type ProtocolRecording,
} from '../src/index.js'
import recording from './fixtures/protocol-redacted-max-tokens.json'
import capturedRequest from './fixtures/codex-0147-request.json'
import { prepareCarrierTurn } from './fixtures/carrier-authorization.js'

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
        res.on('aborted', () => reject(new Error('response_aborted')))
        res.on('error', reject)
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
  it.each(['tool arguments', 'hidden reasoning'])(
    'keeps actively streamed %s alive beyond the downstream socket idle window',
    async (bufferedContent) => {
      const code = 'text(await tools.mcp__wiswork__wiswork_read_document({"title":"设计契约"}))'
      const serialized = JSON.stringify({ code })
      const chunks = Array.from({ length: 12 }, (_, index) =>
        serialized.slice(
          Math.floor((index * serialized.length) / 12),
          Math.floor(((index + 1) * serialized.length) / 12),
        ),
      )
      const frames = [
        {
          type: 'message_start',
          message: { id: 'r1', model: 'openai/gpt-5.6-sol', usage: { input_tokens: 1 } },
        },
        ...(bufferedContent === 'hidden reasoning'
          ? [
              {
                type: 'content_block_start',
                index: 0,
                content_block: { type: 'thinking', thinking: '', signature: '' },
              },
              ...chunks.map(() => ({
                type: 'content_block_delta',
                index: 0,
                delta: { type: 'thinking_delta', thinking: 'private hidden reasoning' },
              })),
              { type: 'content_block_stop', index: 0 },
            ]
          : []),
        {
          type: 'content_block_start',
          index: bufferedContent === 'hidden reasoning' ? 1 : 0,
          content_block: { type: 'tool_use', id: 'call-1', name: 'exec', input: {} },
        },
        ...(bufferedContent === 'hidden reasoning' ? [serialized] : chunks).map((partial_json) => ({
          type: 'content_block_delta',
          index: bufferedContent === 'hidden reasoning' ? 1 : 0,
          delta: { type: 'input_json_delta', partial_json },
        })),
        { type: 'content_block_stop', index: bufferedContent === 'hidden reasoning' ? 1 : 0 },
        { type: 'message_delta', delta: { stop_reason: 'tool_use' } },
        { type: 'message_stop' },
      ]
      let nextFrame = 0
      let cancelled = false
      const diagnostics: string[] = []
      const activity: Array<string | undefined> = []
      const bridge = await startResponsesBridge({
        fetchWithAuth: async () =>
          new Response(
            new ReadableStream({
              async pull(controller) {
                if (nextFrame === frames.length) return controller.close()
                const frame = frames[nextFrame++]
                if (frame.type === 'content_block_delta')
                  await new Promise((resolve) => setTimeout(resolve, 20))
                if (!cancelled)
                  controller.enqueue(new TextEncoder().encode(`data: ${JSON.stringify(frame)}\n\n`))
              },
              cancel() {
                cancelled = true
              },
            }),
            { headers: { 'content-type': 'text/event-stream' } },
          ),
        prepareTurn: (input) => ({ ...prepareCarrierTurn(input), turnId: 'turn_1' }),
        diagnostics: (code) => diagnostics.push(code),
        onStreamActivity: (turnId) => {
          activity.push(turnId)
          throw new Error('private observer failure')
        },
        maxStreamIdleMs: 100,
      })
      try {
        const result = await post(
          new URL(bridge.responsesUrl),
          bridge.secret,
          JSON.stringify(capturedRequest),
        )
        expect(result.status).toBe(200)
        expect(result.body).toContain('response.completed')
        expect(result.body).toContain(JSON.stringify(code))
        const completedCalls = result.body
          .split('\n')
          .filter((line) => line.startsWith('data: {'))
          .map((line) => JSON.parse(line.slice(6)))
          .filter(
            (event) =>
              event.type === 'response.output_item.done' && event.item?.type === 'custom_tool_call',
          )
        expect(completedCalls).toHaveLength(1)
        expect(result.body).not.toContain('private hidden reasoning')
        expect(activity.length).toBeGreaterThan(12)
        expect(new Set(activity)).toEqual(new Set(['turn_1']))
        expect(diagnostics).toEqual(['responses_upstream_started'])
      } finally {
        await bridge.close()
      }
    },
  )

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

  it.each(['invalid_messages_sse', 'unsafe_custom_tool_input', 'tool_input_retry_limit_exceeded'])(
    'reports the closed protocol reason %s without upstream content',
    async (protocolCode) => {
      const diagnostics: string[] = []
      const onDeterministicFailure = vi.fn()
      const bridge = await startResponsesBridge({
        fetchWithAuth: async () =>
          new Response('data: private\n\n', {
            headers: { 'content-type': 'text/event-stream' },
          }),
        prepareTurn: () => ({
          ...prepared(),
          turnId: 'turn-a',
          async *messagesStreamToResponses() {
            yield 'event: response.created\ndata: {"type":"response.created"}\n\n'
            const error = new Error(protocolCode)
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
          `responses_stream_${protocolCode}`,
        ])
        expect(JSON.stringify(diagnostics)).not.toContain('private')
        expect(onDeterministicFailure).toHaveBeenCalledWith(protocolCode, 'turn-a')
      } finally {
        await bridge.close()
      }
    },
  )

  it('normalizes a private parser rejection and fails the active turn deterministically', async () => {
    const diagnostics: string[] = []
    const onDeterministicFailure = vi.fn()
    const bridge = await startResponsesBridge({
      fetchWithAuth: async () =>
        new Response('data: private\n\n', {
          headers: { 'content-type': 'text/event-stream' },
        }),
      prepareTurn: () => ({
        ...prepared(),
        turnId: 'turn-private-parser',
        // eslint-disable-next-line require-yield -- Simulate rejection before the first frame.
        async *messagesStreamToResponses() {
          const error = new Error('private_new_parser_code')
          error.name = 'ProtocolCompatibilityError'
          throw error
        },
      }),
      diagnostics: (code) => diagnostics.push(code),
      onDeterministicFailure,
    })
    try {
      await post(new URL(bridge.responsesUrl), bridge.secret, '{}').catch(() => undefined)
      expect(diagnostics).toContain('responses_stream_invalid_messages_event')
      expect(onDeterministicFailure).toHaveBeenCalledWith(
        'invalid_messages_event',
        'turn-private-parser',
      )
      expect(JSON.stringify(diagnostics)).not.toContain('private_new_parser_code')
    } finally {
      await bridge.close()
    }
  })

  it('reports an upstream stream stall as a deterministic turn failure', async () => {
    const onDeterministicFailure = vi.fn()
    const bridge = await startResponsesBridge({
      fetchWithAuth: async () =>
        new Response(
          new ReadableStream({
            start(controller) {
              controller.enqueue(new TextEncoder().encode('data: partial\n\n'))
            },
          }),
          { headers: { 'content-type': 'text/event-stream' } },
        ),
      prepareTurn: () => ({
        ...prepared(),
        turnId: 'turn-stalled',
        async *messagesStreamToResponses(chunks) {
          for await (const chunk of chunks)
            yield typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk)
        },
      }),
      onDeterministicFailure,
      maxStreamIdleMs: 10,
    })
    try {
      await post(new URL(bridge.responsesUrl), bridge.secret, '{}').catch(() => undefined)
      expect(onDeterministicFailure).toHaveBeenCalledWith(
        'upstream_stream_interrupted',
        'turn-stalled',
      )
    } finally {
      await bridge.close()
    }
  })

  it.each(['upstream stall', 'absolute deadline'])(
    'still terminates buffered input on %s and cancels its source',
    async (limit) => {
      let cancelled = false
      let sentStart = false
      const bridge = await startResponsesBridge({
        fetchWithAuth: async () =>
          new Response(
            new ReadableStream({
              async pull(controller) {
                if (sentStart && limit === 'upstream stall') return
                sentStart = true
                await new Promise((resolve) => setTimeout(resolve, 10))
                if (!cancelled) controller.enqueue(new TextEncoder().encode('partial input'))
              },
              cancel() {
                cancelled = true
              },
            }),
            { headers: { 'content-type': 'text/event-stream' } },
          ),
        prepareTurn: () => ({
          ...prepared(),
          async *messagesStreamToResponses(chunks) {
            yield 'event: response.created\ndata: {"type":"response.created"}\n\n'
            for await (const _chunk of chunks) {
              // Deliberately buffer: no executable call exists until all input arrives.
            }
          },
        }),
        maxStreamIdleMs: 100,
        maxTurnDurationMs: limit === 'absolute deadline' ? 60 : 2_000,
      })
      try {
        await expect(post(new URL(bridge.responsesUrl), bridge.secret, '{}')).rejects.toThrow()
        await vi.waitFor(() => expect(cancelled).toBe(true))
      } finally {
        await bridge.close()
      }
    },
  )
})
