import { describe, expect, it, vi } from 'vitest'
import { WISWORK_DEFAULT_MODEL, WISWORK_MESSAGES_URL } from '@wiswork/ai-provider'
import {
  MAX_REQUEST_BODY_LENGTH,
  MAX_STREAM_TOOL_INPUT_LENGTH,
  createOfficeAgentTransport,
} from '../src/agent/transport.js'
import type { BrowserAuth } from '../src/auth/browser-auth.js'
import type { RuntimeConfig } from '../src/config.js'

const config = { messagesUrl: WISWORK_MESSAGES_URL } as RuntimeConfig

function sse(lines: string[]): Response {
  return new Response(`${lines.join('\n')}\n`, {
    status: 200,
    headers: { 'content-type': 'text/event-stream' },
  })
}

function callbacks() {
  return {
    onDelta: vi.fn(),
    onToolCall: vi.fn(),
    onStopReason: vi.fn(),
    onDone: vi.fn(),
    onError: vi.fn(),
  }
}

describe('Office Agent transport', () => {
  it('uses the trusted endpoint and fixed model/region', async () => {
    const authenticatedFetch = vi.fn().mockResolvedValue(sse([]))
    const auth = { authenticatedFetch } as unknown as BrowserAuth
    const cb = callbacks()
    createOfficeAgentTransport(config, auth).stream(
      { system: 'sys', messages: [{ role: 'user', text: 'hi' }], tools: [] },
      cb,
    )
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(authenticatedFetch).toHaveBeenCalledWith(
      WISWORK_MESSAGES_URL,
      expect.objectContaining({ method: 'POST', signal: expect.any(AbortSignal) }),
    )
    const init = authenticatedFetch.mock.calls[0]![1] as RequestInit
    expect(new Headers(init.headers).get('x-req-location')).toBe('sg')
    expect(JSON.parse(init.body as string)).toMatchObject({ model: WISWORK_DEFAULT_MODEL })
  })

  it('fails closed when runtime configuration does not normalize to the trusted endpoint', () => {
    const auth = { authenticatedFetch: vi.fn() } as unknown as BrowserAuth
    expect(() =>
      createOfficeAgentTransport({ ...config, messagesUrl: 'https://evil.test/v1/messages' }, auth),
    ).toThrow('transport_unavailable')
    expect(auth.authenticatedFetch).not.toHaveBeenCalled()
  })

  it('normalizes text, tool calls, and stop reasons', async () => {
    const auth = {
      authenticatedFetch: vi
        .fn()
        .mockResolvedValue(
          sse([
            'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"hello"}}',
            'data: {"type":"content_block_start","index":1,"content_block":{"type":"tool_use","id":"c1","name":"read_selection"}}',
            'data: {"type":"content_block_delta","index":1,"delta":{"type":"input_json_delta","partial_json":"{}"}}',
            'data: {"type":"content_block_stop","index":1}',
            'data: {"type":"message_delta","delta":{"stop_reason":"tool_use"}}',
          ]),
        ),
    } as unknown as BrowserAuth
    const cb = callbacks()
    createOfficeAgentTransport(config, auth).stream({ system: '', messages: [], tools: [] }, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onDelta).toHaveBeenCalledWith('hello')
    expect(cb.onToolCall).toHaveBeenCalledWith({ id: 'c1', name: 'read_selection', input: {} })
    expect(cb.onStopReason).toHaveBeenCalledWith('tool_use')
    expect(cb.onError).not.toHaveBeenCalled()
  })

  it('cancels and completes exactly once without surfacing an abort error', async () => {
    const auth = {
      authenticatedFetch: vi.fn(
        (_url, init: RequestInit) =>
          new Promise<Response>((_resolve, reject) => {
            init.signal?.addEventListener('abort', () =>
              reject(new DOMException('secret', 'AbortError')),
            )
          }),
      ),
    } as unknown as BrowserAuth
    const cb = callbacks()
    const handle = createOfficeAgentTransport(config, auth).stream(
      { system: '', messages: [], tools: [] },
      cb,
    )
    handle.cancel()
    handle.cancel()
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).not.toHaveBeenCalled()
  })

  it('reports stable HTTP/network errors without upstream bodies', async () => {
    const secret = 'upstream-secret-body'
    const auth = {
      authenticatedFetch: vi.fn().mockResolvedValue(new Response(secret, { status: 502 })),
    } as unknown as BrowserAuth
    const cb = callbacks()
    createOfficeAgentTransport(config, auth).stream({ system: '', messages: [], tools: [] }, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).toHaveBeenCalledWith('transport_http_502')
    expect(cb.onError.mock.calls.flat().join(' ')).not.toContain(secret)
  })

  it('does not forward arbitrary exception codes', async () => {
    const auth = {
      authenticatedFetch: vi.fn().mockRejectedValue({ code: 'private_secret_detail' }),
    } as unknown as BrowserAuth
    const cb = callbacks()
    createOfficeAgentTransport(config, auth).stream({ system: '', messages: [], tools: [] }, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).toHaveBeenCalledWith('transport_network')
  })

  it('does not trust transport-prefixed exception messages from dependencies', async () => {
    const auth = {
      authenticatedFetch: vi.fn().mockRejectedValue(new Error('transport_token_private-secret')),
    } as unknown as BrowserAuth
    const cb = callbacks()
    createOfficeAgentTransport(config, auth).stream({ system: '', messages: [], tools: [] }, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).toHaveBeenCalledWith('transport_network')
  })

  it('bounds accumulated streamed tool input across individually valid SSE lines', async () => {
    const fragment = 'x'.repeat(1024)
    const deltas = Array.from(
      { length: Math.ceil(MAX_STREAM_TOOL_INPUT_LENGTH / fragment.length) + 1 },
      () =>
        `data: ${JSON.stringify({ type: 'content_block_delta', index: 1, delta: { type: 'input_json_delta', partial_json: fragment } })}`,
    )
    const auth = {
      authenticatedFetch: vi
        .fn()
        .mockResolvedValue(
          sse([
            'data: {"type":"content_block_start","index":1,"content_block":{"type":"tool_use","id":"c1","name":"propose_append_text"}}',
            ...deltas,
          ]),
        ),
    } as unknown as BrowserAuth
    const cb = callbacks()
    createOfficeAgentTransport(config, auth).stream({ system: '', messages: [], tools: [] }, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).toHaveBeenCalledWith('transport_tool_input_too_large')
    expect(cb.onToolCall).not.toHaveBeenCalled()
  })

  it('bounds an unterminated SSE line before it can accumulate indefinitely', async () => {
    const auth = {
      authenticatedFetch: vi
        .fn()
        .mockResolvedValue(new Response(`data: ${'x'.repeat(70 * 1024)}`, { status: 200 })),
    } as unknown as BrowserAuth
    const cb = callbacks()
    createOfficeAgentTransport(config, auth).stream({ system: '', messages: [], tools: [] }, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).toHaveBeenCalledWith('transport_stream_too_large')
  })

  it('rejects an oversized outbound request before authenticated network I/O', async () => {
    const auth = { authenticatedFetch: vi.fn() } as unknown as BrowserAuth
    const cb = callbacks()
    createOfficeAgentTransport(config, auth).stream(
      { system: 'x'.repeat(MAX_REQUEST_BODY_LENGTH + 1), messages: [], tools: [] },
      cb,
    )
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).toHaveBeenCalledWith('transport_request_too_large')
    expect(auth.authenticatedFetch).not.toHaveBeenCalled()
  })
})
