import { describe, expect, it } from 'vitest'
import { prepareResponsesTurn, ProtocolCompatibilityError } from '../src/index.js'
import captured from './fixtures/codex-0147-request.json'
import { prepareCarrierTurn } from './fixtures/carrier-authorization.js'

const noToolTurn = () => prepareResponsesTurn({ model: 'gpt-5.6-sol', input: 'Hello' })

const start =
  'event: message_start\ndata: {"type":"message_start","message":{"id":"msg_1","model":"openai/gpt-5.6-sol","usage":{"input_tokens":10,"cache_read_input_tokens":4,"cache_creation_input_tokens":2}}}\n\n'
const delta =
  'event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"end_turn"},"usage":{"output_tokens":3}}\n\n'
const stop = 'event: message_stop\ndata: {"type":"message_stop"}\n\n'

async function* chunks(...values: Array<string | Uint8Array>): AsyncGenerator<string | Uint8Array> {
  yield* values
}

async function collect(
  source: AsyncIterable<string>,
): Promise<Array<{ event: string; data: Record<string, any> }>> {
  const events: Array<{ event: string; data: Record<string, any> }> = []
  for await (const frame of source) {
    const event = /^event: ([^\n]+)$/m.exec(frame)?.[1]
    const data = /^data: (.+)$/m.exec(frame)?.[1]
    if (event && data) events.push({ event, data: JSON.parse(data) })
  }
  return events
}

async function collectUntilFailure(
  source: AsyncIterable<string>,
): Promise<{ frames: string[]; error: unknown }> {
  const frames: string[] = []
  let error: unknown
  try {
    for await (const frame of source) frames.push(frame)
  } catch (caught) {
    error = caught
  }
  return { frames, error }
}

async function expectStreamCode(
  values: Array<string | Uint8Array>,
  code: string,
  custom = false,
  limits: Record<string, number> = {},
): Promise<void> {
  let caught: unknown
  try {
    const turn = custom
      ? prepareCarrierTurn(structuredClone(captured), limits)
      : prepareResponsesTurn({ model: 'gpt-5.6-sol', input: 'Hello' }, limits)
    await collect(turn.messagesStreamToResponses(chunks(...values)))
  } catch (error) {
    caught = error
  }
  expect(caught).toEqual(
    expect.objectContaining<Partial<ProtocolCompatibilityError>>({ message: code }),
  )
  expect(String(caught)).not.toContain('secret prompt')
}

describe('bounded Anthropic SSE state machine', () => {
  it.each([
    ['empty stream', [], 'premature_messages_eof'],
    ['start only', [start], 'premature_messages_eof'],
    ['missing delta', [start, stop], 'invalid_messages_event_order'],
    ['duplicate start', [start, start], 'invalid_messages_event_order'],
    ['duplicate delta', [start, delta, delta], 'invalid_messages_event_order'],
    ['post-terminal event', [start, delta, stop, start], 'post_terminal_messages_event'],
    ['upstream done marker', [start, 'data: [DONE]\n\n'], 'unsupported_messages_sse'],
  ])('rejects %s', async (_label, values, code) => expectStreamCode(values, code))

  it.each([
    ['negative', -1, 'invalid_messages_block_index'],
    ['noncontiguous', 1, 'invalid_messages_block_index'],
  ])('rejects %s block index', async (_label, index, code) => {
    await expectStreamCode(
      [
        start,
        `event: content_block_start\ndata: {"type":"content_block_start","index":${index},"content_block":{"type":"text","text":""}}\n\n`,
      ],
      code,
    )
  })

  it('rejects a stopped block index reused later', async () => {
    await expectStreamCode(
      [
        start,
        'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"text","text":""}}\n\n',
        'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n',
        'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"text","text":""}}\n\n',
      ],
      'invalid_messages_block_index',
    )
  })

  it.each([
    ['ping', `${start}${delta}${stop}event: ping\ndata: {"type":"ping"}\n\n`],
    ['garbage', `${start}${delta}${stop}secret trailing garbage`],
  ])('withholds completion when %s follows message_stop', async (_label, source) => {
    const result = await collectUntilFailure(noToolTurn().messagesStreamToResponses(chunks(source)))
    expect(result.error).toBeInstanceOf(ProtocolCompatibilityError)
    expect(result.frames.join('')).not.toContain('event: response.completed')
    expect(result.frames.join('')).not.toContain('data: [DONE]')
    expect(String(result.error)).not.toContain('secret')
  })

  it('ignores comment keepalives and accepts CR-only SSE line endings', async () => {
    const crStream = `: keepalive\n\n${start}${delta}${stop}`.replaceAll('\n', '\r')
    const events = await collect(noToolTurn().messagesStreamToResponses(chunks(crStream)))
    expect(events.at(-1)?.event).toBe('response.completed')
  })

  it('rejects invalid UTF-8 rather than replacement-decoding', async () => {
    await expectStreamCode([new Uint8Array([0xc3, 0x28])], 'invalid_messages_utf8')
  })

  it.each([
    ['negative input', { input_tokens: -7 }, 'invalid_messages_usage'],
    ['fractional input', { input_tokens: 1.5 }, 'invalid_messages_usage'],
    [
      'negative cache read',
      { input_tokens: 1, cache_read_input_tokens: -1 },
      'invalid_messages_usage',
    ],
    [
      'negative cache creation',
      { input_tokens: 1, cache_creation_input_tokens: -1 },
      'invalid_messages_usage',
    ],
  ])('rejects %s usage', async (_label, usage, code) => {
    await expectStreamCode(
      [
        `event: message_start\ndata: ${JSON.stringify({ type: 'message_start', message: { id: 'x', model: 'openai/gpt-5.6-sol', usage } })}\n\n`,
      ],
      code,
    )
  })

  it('preserves cache creation and read accounting', async () => {
    const events = await collect(noToolTurn().messagesStreamToResponses(chunks(start, delta, stop)))
    expect(events.at(-1)?.data.response.usage).toMatchObject({
      input_tokens: 16,
      input_tokens_details: { cached_tokens: 4, cache_write_tokens: 2 },
    })
  })

  it.each([
    ['frame bytes', { maxSseFrameBytes: 16 }, [start], 'sse_frame_limit_exceeded'],
    ['buffer bytes', { maxSseBufferBytes: 16 }, [start.slice(0, 20)], 'sse_buffer_limit_exceeded'],
    ['frame count', { maxSseFrames: 1 }, [start, delta], 'sse_frame_count_limit_exceeded'],
    [
      'text accumulation',
      { maxAccumulatedText: 3 },
      [
        start,
        'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"text","text":""}}\n\n',
        'event: content_block_delta\ndata: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"hello"}}\n\n',
      ],
      'output_text_limit_exceeded',
    ],
    [
      'tool arguments',
      { maxToolArguments: 3 },
      [
        start,
        'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"c","name":"exec","input":{}}}\n\n',
        'event: content_block_delta\ndata: {"type":"content_block_delta","index":0,"delta":{"type":"input_json_delta","partial_json":"{\\"x\\":1}"}}\n\n',
      ],
      'tool_arguments_limit_exceeded',
    ],
    [
      'block count',
      { maxBlocks: 1 },
      [
        start,
        'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"text","text":""}}\n\n',
        'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n',
        'event: content_block_start\ndata: {"type":"content_block_start","index":1,"content_block":{"type":"text","text":""}}\n\n',
      ],
      'output_block_limit_exceeded',
    ],
    ['total output', { maxTotalOutput: 1 }, [start], 'total_output_limit_exceeded'],
  ])('enforces %s limit', async (_label, limits, values, code) => {
    await expectStreamCode(
      values as string[],
      code as string,
      Boolean((limits as any).maxToolArguments),
      limits as Record<string, number>,
    )
  })

  it('preserves bounded production redacted thinking as encrypted Responses reasoning', async () => {
    const events = await collect(
      noToolTurn().messagesStreamToResponses(
        chunks(
          start,
          'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"redacted_thinking","data":"opaque-reasoning"}}\n\n',
          'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n',
          delta,
          stop,
        ),
      ),
    )

    const reasoning = events.find((event) => event.event === 'response.output_item.done')
    expect(reasoning?.data.item).toEqual({
      id: 'item_0',
      type: 'reasoning',
      status: 'completed',
      summary: [],
      encrypted_content: 'opaque-reasoning',
    })
    expect(events.at(-1)?.data.response.output).toContainEqual(reasoning?.data.item)
  })

  it('accepts the live WisUsage nested plaintext and encrypted reasoning envelope', async () => {
    const events = await collect(
      noToolTurn().messagesStreamToResponses(
        chunks(
          start,
          'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"thinking","thinking":"private prefix","signature":null}}\n\n',
          'event: content_block_delta\ndata: {"type":"content_block_delta","index":0,"delta":{"type":"thinking_delta","thinking":" private suffix"}}\n\n',
          'event: content_block_start\ndata: {"type":"content_block_start","index":1,"content_block":{"type":"redacted_thinking","data":"opaque-live-reasoning"}}\n\n',
          'event: content_block_stop\ndata: {"type":"content_block_stop","index":1}\n\n',
          'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n',
          'event: content_block_start\ndata: {"type":"content_block_start","index":2,"content_block":{"type":"text","text":"","citations":[]}}\n\n',
          'event: content_block_delta\ndata: {"type":"content_block_delta","index":2,"delta":{"type":"text_delta","text":"done"}}\n\n',
          'event: content_block_stop\ndata: {"type":"content_block_stop","index":2}\n\n',
          delta,
          stop,
        ),
      ),
    )

    const serialized = JSON.stringify(events)
    expect(serialized).not.toContain('private prefix')
    expect(serialized).not.toContain('private suffix')
    expect(events).toContainEqual(
      expect.objectContaining({
        event: 'response.output_item.done',
        data: expect.objectContaining({
          item: expect.objectContaining({
            id: 'item_0',
            type: 'reasoning',
            encrypted_content: 'opaque-live-reasoning',
          }),
        }),
      }),
    )
    expect(events.at(-1)?.event).toBe('response.completed')
  })

  it.each([
    [
      'inside a text block',
      [
        'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"text","text":""}}\n\n',
        'event: content_block_start\ndata: {"type":"content_block_start","index":1,"content_block":{"type":"redacted_thinking","data":"opaque"}}\n\n',
      ],
      'invalid_messages_event_order',
    ],
    [
      'with a noncontiguous index',
      [
        'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"thinking","thinking":""}}\n\n',
        'event: content_block_start\ndata: {"type":"content_block_start","index":2,"content_block":{"type":"redacted_thinking","data":"opaque"}}\n\n',
      ],
      'invalid_messages_block_index',
    ],
    [
      'without closing the encrypted child first',
      [
        'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"thinking","thinking":""}}\n\n',
        'event: content_block_start\ndata: {"type":"content_block_start","index":1,"content_block":{"type":"redacted_thinking","data":"opaque"}}\n\n',
        'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n',
      ],
      'invalid_messages_event_order',
    ],
  ])('rejects a nested encrypted reasoning block %s', async (_label, events, code) => {
    await expectStreamCode([start, ...events], code)
  })

  it('bounds the combined plaintext and encrypted nested reasoning envelope', async () => {
    await expectStreamCode(
      [
        start,
        'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"thinking","thinking":"private-text"}}\n\n',
        'event: content_block_start\ndata: {"type":"content_block_start","index":1,"content_block":{"type":"redacted_thinking","data":"opaque-value"}}\n\n',
      ],
      'reasoning_content_limit_exceeded',
      false,
      { maxStringLength: 20 },
    )
  })

  it('bounds encrypted reasoning without exposing it in the error', async () => {
    await expectStreamCode(
      [
        start,
        'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"redacted_thinking","data":"secret prompt that is too long"}}\n\n',
      ],
      'reasoning_content_limit_exceeded',
      false,
      { maxStringLength: 20 },
    )
  })

  it('validates and discards bounded plaintext thinking without exposing it or blocking the turn', async () => {
    const events = await collect(
      noToolTurn().messagesStreamToResponses(
        chunks(
          start,
          'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"thinking","thinking":""}}\n\n',
          'event: content_block_delta\ndata: {"type":"content_block_delta","index":0,"delta":{"type":"thinking_delta","thinking":"private chain of thought"}}\n\n',
          'event: content_block_delta\ndata: {"type":"content_block_delta","index":0,"delta":{"type":"signature_delta","signature":"opaque-signature"}}\n\n',
          'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n',
          delta,
          stop,
        ),
      ),
    )

    expect(JSON.stringify(events)).not.toContain('private chain of thought')
    expect(JSON.stringify(events)).not.toContain('opaque-signature')
    expect(events).toContainEqual(
      expect.objectContaining({
        event: 'response.output_item.done',
        data: expect.objectContaining({
          item: {
            id: 'item_0',
            type: 'reasoning',
            status: 'completed',
            summary: [],
          },
        }),
      }),
    )
    expect(events.at(-1)?.event).toBe('response.completed')
  })

  it.each([
    [
      'inline thinking',
      { type: 'thinking', thinking: 'private inline reasoning' },
      [] as Array<Record<string, unknown>>,
    ],
    [
      'inline thinking and signature',
      { type: 'thinking', thinking: 'private inline reasoning', signature: 'inline-signature' },
      [] as Array<Record<string, unknown>>,
    ],
    [
      'normalized null signature',
      { type: 'thinking', thinking: '', signature: null },
      [{ type: 'thinking_delta', thinking: 'private normalized reasoning' }],
    ],
    [
      'mixed inline and delta thinking',
      { type: 'thinking', thinking: 'private prefix', signature: '' },
      [
        { type: 'thinking_delta', thinking: ' private suffix' },
        { type: 'signature_delta', signature: 'delta-signature' },
      ],
    ],
  ])('accepts and fully redacts %s', async (_label, contentBlock, deltas) => {
    const stream = [
      start,
      `event: content_block_start\ndata: ${JSON.stringify({ type: 'content_block_start', index: 0, content_block: contentBlock })}\n\n`,
      ...deltas.map(
        (reasoningDelta) =>
          `event: content_block_delta\ndata: ${JSON.stringify({ type: 'content_block_delta', index: 0, delta: reasoningDelta })}\n\n`,
      ),
      'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n',
      delta,
      stop,
    ]
    const events = await collect(noToolTurn().messagesStreamToResponses(chunks(...stream)))

    const serialized = JSON.stringify(events)
    expect(serialized).not.toContain('private')
    expect(serialized).not.toContain('signature')
    expect(events.at(-1)?.event).toBe('response.completed')
  })

  it('bounds discarded plaintext thinking and its signature', async () => {
    await expectStreamCode(
      [
        start,
        'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"thinking","thinking":""}}\n\n',
        `event: content_block_delta\ndata: ${JSON.stringify({ type: 'content_block_delta', index: 0, delta: { type: 'thinking_delta', thinking: '1'.repeat(60) } })}\n\n`,
        `event: content_block_delta\ndata: ${JSON.stringify({ type: 'content_block_delta', index: 0, delta: { type: 'signature_delta', signature: '2'.repeat(41) } })}\n\n`,
      ],
      'reasoning_content_limit_exceeded',
      false,
      { maxStringLength: 100 },
    )
  })

  it('accepts bounded OpenAI reasoning token usage from the normalized WisUsage stream', async () => {
    const events = await collect(
      noToolTurn().messagesStreamToResponses(
        chunks(
          start,
          'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"thinking","thinking":"private"}}\n\n',
          'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n',
          'event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"end_turn"},"usage":{"output_tokens":3,"output_tokens_details":{"reasoning_tokens":2}}}\n\n',
          stop,
        ),
      ),
    )

    expect(JSON.stringify(events)).not.toContain('private')
    expect(events.at(-1)?.data.response.usage.output_tokens_details).toEqual({
      reasoning_tokens: 2,
    })
    expect(events.at(-1)?.event).toBe('response.completed')
  })

  it('rejects ambiguous normalized reasoning token usage', async () => {
    await expectStreamCode(
      [
        start,
        'event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"end_turn"},"usage":{"output_tokens":3,"output_tokens_details":{"thinking_tokens":1,"reasoning_tokens":2}}}\n\n',
      ],
      'invalid_messages_usage',
      false,
    )
  })

  it.each([
    ['non-string thinking', { type: 'thinking', thinking: null }],
    ['structured signature', { type: 'thinking', thinking: '', signature: { value: 'secret' } }],
    ['unknown field', { type: 'thinking', thinking: '', private: 'secret' }],
  ])('rejects malformed plaintext reasoning: %s', async (_label, contentBlock) => {
    await expectStreamCode(
      [
        start,
        `event: content_block_start\ndata: ${JSON.stringify({ type: 'content_block_start', index: 0, content_block: contentBlock })}\n\n`,
      ],
      'unsupported_reasoning_block',
      false,
    )
  })

  it('maps max-token truncation inside tool JSON to an incomplete response without a tool call', async () => {
    const events = await collect(
      prepareCarrierTurn(structuredClone(captured)).messagesStreamToResponses(
        chunks(
          start,
          'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"redacted_thinking","data":"opaque-reasoning"}}\n\n',
          'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n',
          'event: content_block_start\ndata: {"type":"content_block_start","index":1,"content_block":{"type":"tool_use","id":"c","name":"exec","input":{}}}\n\n',
          'event: content_block_delta\ndata: {"type":"content_block_delta","index":1,"delta":{"type":"input_json_delta","partial_json":"{\\"code\\":\\"await tools.mcp__wiswork__wiswork_read_document("}}\n\n',
          'event: content_block_stop\ndata: {"type":"content_block_stop","index":1}\n\n',
          'event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"max_tokens"},"usage":{"output_tokens":4096}}\n\n',
          stop,
        ),
      ),
    )

    expect(
      events.filter(
        (event) =>
          event.event === 'response.output_item.added' &&
          event.data.item?.type === 'custom_tool_call',
      ),
    ).toHaveLength(0)
    expect(events.map((event) => event.event)).not.toContain('response.custom_tool_call_input.done')
    expect(events.at(-1)?.event).toBe('response.incomplete')
    expect(events.at(-1)?.data.response.incomplete_details).toEqual({
      reason: 'max_output_tokens',
    })
  })

  it('does not execute even syntactically complete tool JSON when the provider reports max tokens', async () => {
    const safeCode = 'await tools.mcp__wiswork__wiswork_read_document({})'
    const events = await collect(
      prepareCarrierTurn(structuredClone(captured)).messagesStreamToResponses(
        chunks(
          start,
          'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"c","name":"exec","input":{}}}\n\n',
          `event: content_block_delta\ndata: ${JSON.stringify({ type: 'content_block_delta', index: 0, delta: { type: 'input_json_delta', partial_json: JSON.stringify({ code: safeCode }) } })}\n\n`,
          'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n',
          'event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"max_tokens"},"usage":{"output_tokens":4096}}\n\n',
          stop,
        ),
      ),
    )

    expect(events.some((event) => event.data.item?.type === 'custom_tool_call')).toBe(false)
    expect(events.at(-1)?.event).toBe('response.incomplete')
  })

  it('rejects an unadvertised tool from the bound turn', async () => {
    await expectStreamCode(
      [
        start,
        'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"c","name":"other","input":{}}}\n\n',
      ],
      'unadvertised_tool_call',
      true,
    )
  })

  it.each([
    'text(await tools.apply_patch({}))',
    'text(await tools["mcp__wiswork__wiswork_read_document"]({}))',
    'const x = "mcp__wiswork__wiswork_read_document"; text(await tools[x]({}))',
    'text(await tools.mcp__wiswork__wiswork_read_document({})); text(1)',
    'text(await tools.mcp__wiswork__wiswork_read_document({x: 1}))',
  ])('rejects unsafe custom code without echoing it', async (code) => {
    await expectStreamCode(
      [
        start,
        'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"c","name":"exec","input":{}}}\n\n',
        `event: content_block_delta\ndata: ${JSON.stringify({ type: 'content_block_delta', index: 0, delta: { type: 'input_json_delta', partial_json: JSON.stringify({ code }) } })}\n\n`,
        'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n',
        'event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"tool_use"},"usage":{"output_tokens":1}}\n\n',
        stop,
      ],
      'unsafe_custom_tool_input',
      true,
    )
  })

  it('rejects a giant single transport chunk before buffering it', async () => {
    await expectStreamCode(['x'.repeat(1_000_000)], 'sse_buffer_limit_exceeded', false, {
      maxSseFrameBytes: 64,
      maxSseBufferBytes: 64,
    })
  })
})
