import { describe, expect, it, vi } from 'vitest'
import { WISWORK_DEFAULT_MODEL } from '@wiswork/ai-provider'
import {
  MAX_COMPLETED_TOOL_CALLS,
  MAX_REQUEST_BODY_LENGTH,
  MAX_STREAM_TEXT_LENGTH,
  MAX_STREAM_TOOL_INPUT_LENGTH,
  STREAM_RESPONSE_TIMEOUT_MS,
  createPcBridgeAgentTransport,
} from '../src/agent/transport.js'

interface TestBridge {
  authenticatedFetch: (path: string, init: RequestInit) => Promise<Response>
}

const createTestTransport = (bridge: TestBridge) => createPcBridgeAgentTransport(bridge as never)

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
  it.each([
    ['relay_session_expired', 'session_expired'],
    ['relay_request_timeout', 'request_timeout'],
    ['relay_auth_required', 'auth_required'],
    ['relay_upstream_error', 'provider_unavailable'],
    ['relay_disconnected', 'network_error'],
  ])('preserves the safe category for %s', async (message, code) => {
    const cb = callbacks()
    createTestTransport({
      authenticatedFetch: async () => {
        throw new Error(message)
      },
    }).stream({ system: '', messages: [], tools: [] }, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).toHaveBeenCalledWith(code)
  })

  const activity = (state = 'running', extra: Record<string, unknown> = {}) => ({
    type: 'wiswork_tool_activity',
    generation: 3,
    call_id: 'call_search123',
    tool_name: 'image_search',
    state,
    started_at: 1000,
    query: 'volcano',
    ...extra,
  })
  const searchRequest = {
    system: '',
    messages: [],
    tools: [{ name: 'image_search', description: 'images', inputSchema: { type: 'object' } }],
  }

  it('observes host retrieval without asking the local agent to execute it', async () => {
    const cb = callbacks()
    const observe = vi.fn()
    const handleToolFrame = vi.fn()
    const transport = createPcBridgeAgentTransport({
      snapshot: () => ({ enhanced: { session_generation: 3 } }),
      handleToolFrame,
      authenticatedFetch: vi.fn(async () =>
        sse([
          `data: ${JSON.stringify(activity())}`,
          `data: ${JSON.stringify(activity('complete', { summary: '1 result', result_count: 1, display: { kind: 'images', items: [{ url: 'https://example.com/volcano', title: 'Volcano' }] } }))}`,
        ]),
      ),
    } as any)
    transport.setToolActivityHandler?.(observe)
    transport.stream(searchRequest, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(observe.mock.calls.map(([event]) => event.state)).toEqual(['running', 'complete'])
    expect(observe.mock.calls[1]?.[0].display.items[0].url).toBe('https://example.com/volcano')
    expect(cb.onToolCall).not.toHaveBeenCalled()
    expect(handleToolFrame).not.toHaveBeenCalled()
    expect(cb.onError).not.toHaveBeenCalled()
  })

  it.each([33, 1024])(
    'observes %i semantic attempts within the document-session limit',
    async (count) => {
      const cb = callbacks(),
        observe = vi.fn(),
        handleToolFrame = vi.fn()
      const transport = createPcBridgeAgentTransport({
        snapshot: () => ({ enhanced: { session_generation: 3 } }),
        handleToolFrame,
        authenticatedFetch: async () =>
          sse(
            Array.from({ length: count }, (_, index) =>
              ['running', 'complete'].map(
                (state) =>
                  `data: ${JSON.stringify(activity(state, { call_id: `call_search_${index}` }))}`,
              ),
            ).flat(),
          ),
      })
      transport.setToolActivityHandler?.(observe)
      transport.stream(searchRequest, cb)
      await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
      expect(cb.onError).not.toHaveBeenCalled()
      expect(observe).toHaveBeenCalledTimes(count * 2)
      expect(cb.onToolCall).not.toHaveBeenCalled()
      expect(handleToolFrame).not.toHaveBeenCalled()
    },
  )

  it('observes allowed non-retrieval router rejection without executing the tool', async () => {
    const cb = callbacks(),
      observe = vi.fn(),
      handleToolFrame = vi.fn()
    const transport = createPcBridgeAgentTransport({
      snapshot: () => ({ enhanced: { session_generation: 3 } }),
      handleToolFrame,
      authenticatedFetch: async () =>
        sse(
          ['running', 'error'].map(
            (state) =>
              `data: ${JSON.stringify(activity(state, { type: 'wiswork_tool_lifecycle', tool_name: 'get_document_text', query: undefined }))}`,
          ),
        ),
    })
    transport.setToolActivityHandler?.(observe)
    transport.stream(
      {
        ...searchRequest,
        tools: [
          { name: 'get_document_text', description: 'read', inputSchema: { type: 'object' } },
        ],
      },
      cb,
    )
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).not.toHaveBeenCalled()
    expect(observe.mock.calls.map(([event]) => event.state)).toEqual(['running', 'error'])
    expect(cb.onToolCall).not.toHaveBeenCalled()
    expect(handleToolFrame).not.toHaveBeenCalled()
  })

  it('retains the hard 1024-call bound for enhanced observation streams', async () => {
    const cb = callbacks(),
      observe = vi.fn()
    const transport = createPcBridgeAgentTransport({
      snapshot: () => ({ enhanced: { session_generation: 3 } }),
      authenticatedFetch: async () =>
        sse(
          Array.from({ length: 1025 }, (_, index) =>
            ['running', 'complete'].map(
              (state) =>
                `data: ${JSON.stringify(activity(state, { call_id: `call_search_${index}` }))}`,
            ),
          ).flat(),
        ),
    })
    transport.setToolActivityHandler?.(observe)
    transport.stream(searchRequest, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).toHaveBeenCalledWith('transport_invalid_stream')
    expect(observe).toHaveBeenCalledTimes(2048)
  })

  it('bounds activity metadata by UTF-8 bytes, not JavaScript character count', async () => {
    const frame = activity('running', {
      display: {
        kind: 'images',
        items: Array.from({ length: 8 }, () => ({
          url: `https://example.com/${'x'.repeat(1000)}`,
          title: '图'.repeat(160),
        })),
      },
    })
    expect(JSON.stringify(frame).length).toBeLessThan(12 * 1024)
    expect(new TextEncoder().encode(JSON.stringify(frame)).byteLength).toBeGreaterThan(12 * 1024)
    const cb = callbacks(),
      observe = vi.fn()
    const transport = createPcBridgeAgentTransport({
      snapshot: () => ({ enhanced: { session_generation: 3 } }),
      authenticatedFetch: async () => sse([`data: ${JSON.stringify(frame)}`]),
    })
    transport.setToolActivityHandler?.(observe)
    transport.stream(searchRequest, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).toHaveBeenCalledWith('transport_invalid_stream')
    expect(observe).not.toHaveBeenCalled()
  })

  it('binds the initial enhanced generation negotiated after the request starts', async () => {
    let generation: number | undefined
    const cb = callbacks(),
      observe = vi.fn()
    const transport = createPcBridgeAgentTransport({
      snapshot: () =>
        generation === undefined ? {} : { enhanced: { session_generation: generation } },
      authenticatedFetch: vi.fn(async () => {
        // PC promotes an initially standard session while accepting this first request.
        generation = 3
        return sse([
          `data: ${JSON.stringify(activity())}`,
          `data: ${JSON.stringify(activity('complete', { summary: 'Retrieval complete', result_count: 0 }))}`,
        ])
      }),
    })
    transport.setToolActivityHandler?.(observe)
    transport.stream(searchRequest, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(observe.mock.calls.map(([event]) => event.state)).toEqual(['running', 'complete'])
    expect(cb.onError).not.toHaveBeenCalled()
    expect(cb.onToolCall).not.toHaveBeenCalled()
  })

  it.each([undefined, 3])(
    'rejects a generation change after the first accepted activity (initial %s)',
    async (initial) => {
      let generation = initial
      let stream!: ReadableStreamDefaultController<Uint8Array>
      const cb = callbacks(),
        observe = vi.fn()
      const transport = createPcBridgeAgentTransport({
        snapshot: () =>
          generation === undefined ? {} : { enhanced: { session_generation: generation } },
        authenticatedFetch: vi.fn(async () => {
          generation = 3
          return new Response(
            new ReadableStream({
              start(controller) {
                stream = controller
                controller.enqueue(
                  new TextEncoder().encode(`data: ${JSON.stringify(activity())}\n\n`),
                )
              },
            }),
          )
        }),
      })
      transport.setToolActivityHandler?.(observe)
      transport.stream(searchRequest, cb)
      await vi.waitFor(() => expect(observe).toHaveBeenCalledOnce())
      generation = 4
      stream.enqueue(
        new TextEncoder().encode(
          `data: ${JSON.stringify(activity('complete', { generation: 4, summary: 'Retrieval complete' }))}\n\n`,
        ),
      )
      stream.close()
      await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
      expect(cb.onError).toHaveBeenCalledWith('transport_invalid_stream')
      expect(observe.mock.calls.map(([event]) => event.state)).toEqual(['running', 'error'])
      expect(cb.onToolCall).not.toHaveBeenCalled()
    },
  )

  it.each([undefined, 3])(
    'rejects unnegotiated or replaced generation before the first activity (initial %s)',
    async (initial) => {
      let generation = initial
      const cb = callbacks(),
        observe = vi.fn()
      const transport = createPcBridgeAgentTransport({
        snapshot: () =>
          generation === undefined ? {} : { enhanced: { session_generation: generation } },
        authenticatedFetch: vi.fn(async () => {
          if (generation === 3) generation = 4
          return sse([
            `data: ${JSON.stringify(activity('running', { generation: generation ?? 3 }))}`,
          ])
        }),
      })
      transport.setToolActivityHandler?.(observe)
      transport.stream(searchRequest, cb)
      await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
      expect(cb.onError).toHaveBeenCalledWith('transport_invalid_stream')
      expect(observe).not.toHaveBeenCalled()
    },
  )

  it.each([
    { generation: 2 },
    { tool_name: 'execute_office_js' },
    { query: 'x'.repeat(241) },
    { unexpected_secret: 'secret' },
    { display: { kind: 'images', items: [{ url: 'https://user:secret@example.com/' }] } },
    { display: { kind: 'images', items: [{ url: 'https://127.0.0.1/' }] } },
  ])('rejects invalid or stale host retrieval activity: %j', async (extra) => {
    const cb = callbacks(),
      observe = vi.fn()
    const transport = createPcBridgeAgentTransport({
      snapshot: () => ({ enhanced: { session_generation: 3 } }),
      authenticatedFetch: vi.fn(async () =>
        sse([`data: ${JSON.stringify(activity('running', extra))}`]),
      ),
    } as any)
    transport.setToolActivityHandler?.(observe)
    transport.stream(searchRequest, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(observe).not.toHaveBeenCalled()
    expect(cb.onError).toHaveBeenCalledWith('transport_invalid_stream')
  })

  it('does not publish host activity from a tool-free visual-review subturn', async () => {
    const cb = callbacks(),
      observe = vi.fn()
    const transport = createPcBridgeAgentTransport({
      snapshot: () => ({ enhanced: { session_generation: 3 } }),
      authenticatedFetch: vi.fn(async () => sse([`data: ${JSON.stringify(activity())}`])),
    } as any)
    transport.setToolActivityHandler?.(observe)
    transport.stream({ system: '', messages: [], tools: [] }, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(observe).not.toHaveBeenCalled()
  })

  it('keeps the main retrieval observer active across a tool-free review subturn', async () => {
    let main!: ReadableStreamDefaultController<Uint8Array>
    const observe = vi.fn(),
      cb = callbacks()
    const transport = createPcBridgeAgentTransport({
      snapshot: () => ({ enhanced: { session_generation: 3 } }),
      authenticatedFetch: vi
        .fn()
        .mockResolvedValueOnce(
          new Response(
            new ReadableStream({
              start(controller) {
                main = controller
              },
            }),
          ),
        )
        .mockResolvedValueOnce(sse([])),
    })
    transport.setToolActivityHandler?.(observe)
    transport.stream(searchRequest, cb)
    main.enqueue(new TextEncoder().encode(`data: ${JSON.stringify(activity())}\n\n`))
    await vi.waitFor(() => expect(observe).toHaveBeenCalledOnce())
    const review = callbacks()
    transport.stream({ system: '', messages: [], tools: [] }, review)
    await vi.waitFor(() => expect(review.onDone).toHaveBeenCalledOnce())
    main.enqueue(
      new TextEncoder().encode(
        `data: ${JSON.stringify(activity('complete', { summary: 'Retrieval complete', result_count: 0 }))}\n\n`,
      ),
    )
    main.close()
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(observe.mock.calls.map(([event]) => event.state)).toEqual(['running', 'complete'])
  })

  it.each(['duplicate-start', 'duplicate-result', 'orphan-result'])(
    'rejects %s without duplicate observations or execution',
    async (scenario) => {
      const start = activity(),
        complete = activity('complete', { summary: 'Retrieval complete' })
      const events =
        scenario === 'duplicate-start'
          ? [start, start]
          : scenario === 'duplicate-result'
            ? [start, complete, complete]
            : [complete]
      const cb = callbacks(),
        observe = vi.fn()
      const transport = createPcBridgeAgentTransport({
        snapshot: () => ({ enhanced: { session_generation: 3 } }),
        authenticatedFetch: vi.fn(async () =>
          sse(events.map((event) => `data: ${JSON.stringify(event)}`)),
        ),
      })
      transport.setToolActivityHandler?.(observe)
      transport.stream(searchRequest, cb)
      await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
      expect(cb.onError).toHaveBeenCalledWith('transport_invalid_stream')
      expect(observe.mock.calls.map(([event]) => event.state)).toEqual(
        scenario === 'duplicate-start'
          ? ['running', 'error']
          : scenario === 'duplicate-result'
            ? ['running', 'complete']
            : [],
      )
      expect(cb.onToolCall).not.toHaveBeenCalled()
    },
  )

  it('detaches the observer before late responses when its session is disposed', async () => {
    const cb = callbacks(),
      observe = vi.fn()
    let respond!: (response: Response) => void
    const transport = createPcBridgeAgentTransport({
      snapshot: () => ({ enhanced: { session_generation: 3 } }),
      authenticatedFetch: () =>
        new Promise<Response>((resolve) => {
          respond = resolve
        }),
    })
    transport.setToolActivityHandler?.(observe)
    transport.stream(searchRequest, cb)
    transport.setToolActivityHandler?.(undefined)
    respond(sse([`data: ${JSON.stringify(activity())}`]))
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(observe).not.toHaveBeenCalled()
  })

  it('closes running host activity on cancellation', async () => {
    const cb = callbacks(),
      observe = vi.fn()
    const transport = createPcBridgeAgentTransport({
      snapshot: () => ({ enhanced: { session_generation: 3 } }),
      authenticatedFetch: vi.fn(
        async () =>
          new Response(
            new ReadableStream({
              start(controller) {
                controller.enqueue(
                  new TextEncoder().encode(`data: ${JSON.stringify(activity())}\n\n`),
                )
              },
            }),
          ),
      ),
    } as any)
    transport.setToolActivityHandler?.(observe)
    const handle = transport.stream(searchRequest, cb)
    await vi.waitFor(() => expect(observe).toHaveBeenCalledOnce())
    handle.cancel()
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(observe.mock.calls.map(([event]) => event.state)).toEqual(['running', 'error'])
    expect(observe.mock.calls[1]?.[0].summary).toBe('Tool interrupted')
  })

  it('streams through the local PC bridge without provider credentials', async () => {
    const authenticatedFetch = vi.fn().mockResolvedValue(sse([]))
    const cb = callbacks()
    createPcBridgeAgentTransport({ authenticatedFetch } as never).stream(
      { system: 'sys', messages: [{ role: 'user', text: 'hi' }], tools: [] },
      cb,
    )
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(authenticatedFetch).toHaveBeenCalledWith(
      '/v1/office/messages',
      expect.objectContaining({ method: 'POST', signal: expect.any(AbortSignal) }),
    )
    const headers = new Headers(authenticatedFetch.mock.calls[0]![1].headers)
    expect(headers.has('authorization')).toBe(false)
  })
  it('uses the fixed provider model through the PC bridge', async () => {
    const authenticatedFetch = vi.fn().mockResolvedValue(sse([]))
    const auth = { authenticatedFetch }
    const cb = callbacks()
    createTestTransport(auth).stream(
      { system: 'sys', messages: [{ role: 'user', text: 'hi' }], tools: [] },
      cb,
    )
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(authenticatedFetch).toHaveBeenCalledWith(
      '/v1/office/messages',
      expect.objectContaining({ method: 'POST', signal: expect.any(AbortSignal) }),
    )
    const init = authenticatedFetch.mock.calls[0]![1] as RequestInit
    expect(new Headers(init.headers).has('authorization')).toBe(false)
    expect(JSON.parse(init.body as string)).toMatchObject({ model: WISWORK_DEFAULT_MODEL })
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
    } as TestBridge
    const cb = callbacks()
    createTestTransport(auth).stream({ system: '', messages: [], tools: [] }, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onDelta).toHaveBeenCalledWith('hello')
    expect(cb.onToolCall).toHaveBeenCalledWith({ id: 'c1', name: 'read_selection', input: {} })
    expect(cb.onStopReason).toHaveBeenCalledWith('tool_use')
    expect(cb.onError).not.toHaveBeenCalled()
  })

  it('cancels and completes exactly once without surfacing an abort error', async () => {
    const auth = {
      authenticatedFetch: vi.fn(() => new Promise<Response>(() => undefined)),
    } as TestBridge
    const cb = callbacks()
    const handle = createTestTransport(auth).stream({ system: '', messages: [], tools: [] }, cb)
    handle.cancel()
    handle.cancel()
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).not.toHaveBeenCalled()
  })

  it('reports stable HTTP/network errors without upstream bodies', async () => {
    const secret = 'upstream-secret-body'
    const auth = {
      authenticatedFetch: vi.fn().mockResolvedValue(new Response(secret, { status: 502 })),
    } as TestBridge
    const cb = callbacks()
    createTestTransport(auth).stream({ system: '', messages: [], tools: [] }, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).toHaveBeenCalledWith('transport_http_502')
    expect(cb.onError.mock.calls.flat().join(' ')).not.toContain(secret)
  })

  it('does not forward arbitrary exception codes', async () => {
    const auth = {
      authenticatedFetch: vi.fn().mockRejectedValue({ code: 'private_secret_detail' }),
    } as TestBridge
    const cb = callbacks()
    createTestTransport(auth).stream({ system: '', messages: [], tools: [] }, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).toHaveBeenCalledWith('transport_network')
  })

  it('does not trust transport-prefixed exception messages from dependencies', async () => {
    const auth = {
      authenticatedFetch: vi.fn().mockRejectedValue(new Error('transport_token_private-secret')),
    } as TestBridge
    const cb = callbacks()
    createTestTransport(auth).stream({ system: '', messages: [], tools: [] }, cb)
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
    } as TestBridge
    const cb = callbacks()
    createTestTransport(auth).stream({ system: '', messages: [], tools: [] }, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).toHaveBeenCalledWith('transport_tool_input_too_large')
    expect(cb.onToolCall).not.toHaveBeenCalled()
  })

  it('bounds an unterminated SSE line before it can accumulate indefinitely', async () => {
    const auth = {
      authenticatedFetch: vi
        .fn()
        .mockResolvedValue(new Response(`data: ${'x'.repeat(70 * 1024)}`, { status: 200 })),
    } as TestBridge
    const cb = callbacks()
    createTestTransport(auth).stream({ system: '', messages: [], tools: [] }, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).toHaveBeenCalledWith('transport_stream_too_large')
  })

  it('rejects an oversized outbound request before authenticated network I/O', async () => {
    const auth = { authenticatedFetch: vi.fn() } as TestBridge
    const cb = callbacks()
    createTestTransport(auth).stream(
      { system: 'x'.repeat(MAX_REQUEST_BODY_LENGTH + 1), messages: [], tools: [] },
      cb,
    )
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).toHaveBeenCalledWith('transport_request_too_large')
    expect(auth.authenticatedFetch).not.toHaveBeenCalled()
  })

  it('bounds cumulative text output across many individually small deltas', async () => {
    const fragment = 'x'.repeat(1024)
    const auth = {
      authenticatedFetch: vi
        .fn()
        .mockResolvedValue(
          sse(
            Array.from(
              { length: Math.ceil(MAX_STREAM_TEXT_LENGTH / fragment.length) + 1 },
              () =>
                `data: ${JSON.stringify({ type: 'content_block_delta', index: 0, delta: { type: 'text_delta', text: fragment } })}`,
            ),
          ),
        ),
    } as TestBridge
    const cb = callbacks()
    createTestTransport(auth).stream({ system: '', messages: [], tools: [] }, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).toHaveBeenCalledWith('transport_stream_budget_exceeded')
    expect(cb.onDelta).toHaveBeenCalledTimes(MAX_STREAM_TEXT_LENGTH / fragment.length)
  })

  it('bounds sequential completed tool calls across the whole response', async () => {
    const lines = Array.from({ length: MAX_COMPLETED_TOOL_CALLS + 1 }, (_, index) => [
      `data: ${JSON.stringify({ type: 'content_block_start', index, content_block: { type: 'tool_use', id: `c${index}`, name: 'read_selection' } })}`,
      `data: ${JSON.stringify({ type: 'content_block_stop', index })}`,
    ]).flat()
    const auth = {
      authenticatedFetch: vi.fn().mockResolvedValue(sse(lines)),
    } as TestBridge
    const cb = callbacks()
    createTestTransport(auth).stream({ system: '', messages: [], tools: [] }, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).toHaveBeenCalledWith('transport_stream_budget_exceeded')
    expect(cb.onToolCall).toHaveBeenCalledTimes(MAX_COMPLETED_TOOL_CALLS)
  })

  it('times out response consumption, cancels the reader, and completes once', async () => {
    vi.useFakeTimers()
    const cancel = vi.fn()
    const hanging = new ReadableStream<Uint8Array>({ cancel })
    const auth = {
      authenticatedFetch: vi.fn().mockResolvedValue(new Response(hanging, { status: 200 })),
    } as TestBridge
    const cb = callbacks()
    createTestTransport(auth).stream({ system: '', messages: [], tools: [] }, cb)
    await vi.advanceTimersByTimeAsync(STREAM_RESPONSE_TIMEOUT_MS)
    expect(cb.onError).toHaveBeenCalledWith('transport_timeout')
    expect(cb.onDone).toHaveBeenCalledOnce()
    expect(cancel).toHaveBeenCalledOnce()
    vi.useRealTimers()
  })

  it('does not extend the standard response budget for text progress', async () => {
    vi.useFakeTimers()
    try {
      let source!: ReadableStreamDefaultController<Uint8Array>
      const cb = callbacks()
      createPcBridgeAgentTransport({
        authenticatedFetch: async () =>
          new Response(
            new ReadableStream({
              start(controller) {
                source = controller
              },
            }),
          ),
      }).stream(searchRequest, cb)
      await vi.advanceTimersByTimeAsync(200_000)
      source.enqueue(
        new TextEncoder().encode(
          'data: {"type":"content_block_delta","delta":{"type":"text_delta","text":"progress"}}\n\n',
        ),
      )
      await vi.advanceTimersByTimeAsync(80_000)
      expect(cb.onDelta).toHaveBeenCalledWith('progress')
      expect(cb.onError).toHaveBeenCalledWith('transport_timeout')
      expect(cb.onDone).toHaveBeenCalledOnce()
    } finally {
      vi.useRealTimers()
    }
  })

  it.each(['complete', 'idle', 'absolute', 'heartbeat'] as const)(
    'bounds an Enhanced multi-step turn by progress and total duration: %s',
    async (ending) => {
      vi.useFakeTimers()
      try {
        let source!: ReadableStreamDefaultController<Uint8Array>
        const cancel = vi.fn()
        const cb = callbacks()
        const observe = vi.fn()
        // The first request can promote the session after stream() has started.
        let enhanced: { session_generation: number } | undefined = undefined
        const transport = createPcBridgeAgentTransport({
          snapshot: () => ({ enhanced }),
          authenticatedFetch: async () =>
            new Response(
              new ReadableStream({
                start(controller) {
                  source = controller
                },
                cancel,
              }),
            ),
        })
        transport.setToolActivityHandler?.(observe)
        transport.stream(searchRequest, cb)
        enhanced = { session_generation: 3 }
        const send = (value: unknown) =>
          source.enqueue(new TextEncoder().encode(`data: ${JSON.stringify(value)}\n\n`))
        if (ending === 'heartbeat') {
          await vi.advanceTimersByTimeAsync(200_000)
          source.enqueue(new TextEncoder().encode(': heartbeat\n\n'))
          send({ type: 'ping' })
          await vi.advanceTimersByTimeAsync(80_000)
        } else {
          const steps = ending === 'absolute' ? 8 : 2
          for (let index = 0; index < steps; index += 1) {
            await vi.advanceTimersByTimeAsync(200_000)
            const extra = { call_id: `call_progress_${index}` }
            send(activity('running', extra))
            send(activity('complete', extra))
            await vi.advanceTimersByTimeAsync(0)
            expect(cb.onError).not.toHaveBeenCalled()
          }
          expect(observe).toHaveBeenCalledTimes(steps * 2)
          if (ending === 'complete') source.close()
          else await vi.advanceTimersByTimeAsync(ending === 'idle' ? 280_000 : 200_000)
        }
        await vi.advanceTimersByTimeAsync(0)
        expect(cb.onDone).toHaveBeenCalledOnce()
        if (ending === 'complete') expect(cb.onError).not.toHaveBeenCalled()
        else {
          expect(cb.onError).toHaveBeenCalledWith('transport_timeout')
          expect(cancel).toHaveBeenCalledOnce()
        }
      } finally {
        vi.useRealTimers()
      }
    },
  )

  it('contains a response reader cancellation rejection during timeout cleanup', async () => {
    vi.useFakeTimers()
    const hanging = new ReadableStream<Uint8Array>({
      cancel: () => Promise.reject(new Error('private cancel failure')),
    })
    const auth = {
      authenticatedFetch: vi.fn().mockResolvedValue(new Response(hanging, { status: 200 })),
    } as TestBridge
    const cb = callbacks()
    createTestTransport(auth).stream({ system: '', messages: [], tools: [] }, cb)
    await vi.advanceTimersByTimeAsync(STREAM_RESPONSE_TIMEOUT_MS)
    expect(cb.onError).toHaveBeenCalledWith('transport_timeout')
    expect(cb.onDone).toHaveBeenCalledOnce()
    vi.useRealTimers()
  })
})
