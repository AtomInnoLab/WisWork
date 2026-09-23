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
  it('keeps the stream deadline below Relay while allowing long model turns', () => {
    expect(STREAM_RESPONSE_TIMEOUT_MS).toBe(280_000)
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
it('sends the actual QA capture image through the PC provider request', async () => {
  const { createPresentationQaSkill } = await import('../src/skills/powerpoint/presentation-qa.js')
  const { InMemoryVfs } = await import('../src/skills/shared/vfs.js')
  const { createHash } = await import('node:crypto')
  const png =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII='
  const artifact = {
    documentId: 'doc',
    projectId: 'project',
    requestId: 'request',
    pptxBase64: 'UEsDBAAAAAA=',
    slideCount: 1,
    pages: [{ id: 'page1', title: 'First', sourceSlideId: '256#' }],
  }
  let record: import('../src/skills/powerpoint/presentation-qa.js').PresentationQaRecord | undefined
  const skill = createPresentationQaSkill({
    available: () => true,
    artifact: () => artifact,
    documentId: async () => 'doc',
    readReceipt: () => ({
      state: 'complete',
      documentId: 'doc',
      slideIds: ['host'],
      checkpoint: {
        version: 1,
        artifactDigest: createHash('sha256').update(artifact.pptxBase64).digest('hex'),
        sourceSlideIds: ['256#'],
        baselineSlideIds: [],
        completed: [{ sourceSlideId: '256#', slideId: 'host' }],
      },
    }),
    inspectPage: async () => ({
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes: [],
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: png },
    }),
    readQa: () => record,
    writeQa: async (_key, value) => {
      record = value
    },
    vfs: new InMemoryVfs(),
  })
  const execution = await skill.executeTool({
    id: 'capture',
    name: 'capture_presentation_page_qa',
    input: { page_id: 'page1' },
  })
  expect(execution.isError).not.toBe(true)
  const authenticatedFetch = vi.fn().mockResolvedValue(sse([])),
    cb = callbacks()
  createTestTransport({ authenticatedFetch }).stream(
    {
      system: 'Review actual image',
      messages: [
        {
          role: 'assistant',
          text: '',
          toolCalls: [
            { id: 'capture', name: 'capture_presentation_page_qa', input: { page_id: 'page1' } },
          ],
        },
        {
          role: 'tool',
          results: [
            {
              id: 'capture',
              name: 'capture_presentation_page_qa',
              output: execution.output,
              content: execution.modelContent,
            },
          ],
        },
      ],
      tools: [],
    },
    cb,
  )
  await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
  expect(cb.onError).not.toHaveBeenCalled()
  const body = JSON.parse(authenticatedFetch.mock.calls[0]![1].body as string)
  expect(body.messages[1].content[0].content).toEqual([
    { type: 'text', text: execution.output },
    { type: 'image', source: { type: 'base64', media_type: 'image/png', data: png } },
  ])
})
it('discloses removed historical images while preserving every newest tool image', async () => {
  const authenticatedFetch = vi.fn().mockResolvedValue(sse([])),
    cb = callbacks()
  const result = (id: string) => ({
    id,
    name: 'screenshot_slide',
    output: 'screenshot',
    content: [
      {
        type: 'image' as const,
        image: { mime: 'image/png' as const, base64: id === 'old' ? 'b2xk' : 'bmV3' },
      },
    ],
  })
  createTestTransport({ authenticatedFetch }).stream(
    {
      system: '',
      messages: [
        { role: 'tool', results: [result('old')] },
        { role: 'assistant', text: 'Inspect next slide' },
        { role: 'tool', results: [result('new')] },
      ],
      tools: [],
    },
    cb,
  )
  await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
  const body = JSON.parse(authenticatedFetch.mock.calls[0]![1].body as string)
  expect(body.messages[0].content[0].content).toContain('not included in this request')
  expect(body.messages[0].content[0].content).toContain('capture')
  expect(body.messages[2].content[0].content[1]).toEqual({
    type: 'image',
    source: { type: 'base64', media_type: 'image/png', data: 'bmV3' },
  })
})
it('rejects oversized current images and multibyte requests without silently dropping media', async () => {
  for (const request of [
    {
      system: '',
      messages: [
        {
          role: 'tool' as const,
          results: [
            {
              id: 'latest',
              name: 'screenshot_slide',
              output: 'screenshot',
              content: [
                {
                  type: 'image' as const,
                  image: {
                    mime: 'image/png' as const,
                    base64: 'A'.repeat(MAX_REQUEST_BODY_LENGTH),
                  },
                },
              ],
            },
          ],
        },
      ],
      tools: [],
    },
    { system: '中'.repeat(Math.ceil(MAX_REQUEST_BODY_LENGTH / 2)), messages: [], tools: [] },
  ]) {
    const authenticatedFetch = vi.fn(),
      cb = callbacks()
    createTestTransport({ authenticatedFetch }).stream(request, cb)
    await vi.waitFor(() => expect(cb.onDone).toHaveBeenCalledOnce())
    expect(cb.onError).toHaveBeenCalledWith('transport_request_too_large')
    expect(authenticatedFetch).not.toHaveBeenCalled()
  }
})
