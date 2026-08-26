import { describe, expect, it } from 'vitest'
import {
  messagesSseToResponses,
  ProtocolCompatibilityError,
  type StreamConversionContext,
} from '../src/index.js'

const noTools: StreamConversionContext = {
  advertisedTools: {},
  usedCallIds: [],
  allowedExecMethods: [],
}

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

async function expectStreamCode(
  values: Array<string | Uint8Array>,
  code: string,
  context: StreamConversionContext = noTools,
  limits: Record<string, number> = {},
): Promise<void> {
  let caught: unknown
  try {
    await collect(messagesSseToResponses(chunks(...values), context, limits))
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

  it('ignores comment keepalives and accepts CR-only SSE line endings', async () => {
    const crStream = `: keepalive\n\n${start}${delta}${stop}`.replaceAll('\n', '\r')
    const events = await collect(messagesSseToResponses(chunks(crStream), noTools))
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
    const events = await collect(messagesSseToResponses(chunks(start, delta, stop), noTools))
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
        'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"c","name":"read","input":{}}}\n\n',
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
    const context = (limits as any).maxToolArguments
      ? { advertisedTools: { read: 'function' as const }, usedCallIds: [], allowedExecMethods: [] }
      : noTools
    await expectStreamCode(
      values as string[],
      code as string,
      context,
      limits as Record<string, number>,
    )
  })

  it('rejects thinking blocks while opaque round-trip is disabled', async () => {
    await expectStreamCode(
      [
        start,
        'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"thinking","thinking":"secret prompt"}}\n\n',
      ],
      'unsupported_reasoning_block',
    )
  })

  it.each([
    [
      'unknown tool',
      'other',
      { advertisedTools: { read: 'function' as const }, usedCallIds: [], allowedExecMethods: [] },
      'unadvertised_tool_call',
    ],
    [
      'wrong exec kind',
      'exec',
      { advertisedTools: { exec: 'function' as const }, usedCallIds: [], allowedExecMethods: [] },
      'tool_kind_mismatch',
    ],
    [
      'used call ID',
      'read',
      {
        advertisedTools: { read: 'function' as const },
        usedCallIds: ['c'],
        allowedExecMethods: [],
      },
      'duplicate_call_id',
    ],
  ])('rejects %s from stream context', async (_label, name, context, code) => {
    await expectStreamCode(
      [
        start,
        `event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"c","name":"${name}","input":{}}}\n\n`,
      ],
      code,
      context,
    )
  })

  it.each([
    'text(await tools.apply_patch({}))',
    'text(await tools["mcp__wiswork__read_document"]({}))',
    'const x = "mcp__wiswork__read_document"; text(await tools[x]({}))',
    'text(await tools.mcp__wiswork__read_document({})); text(1)',
    'text(await tools.mcp__wiswork__read_document({x: 1}))',
  ])('rejects unsafe custom code without echoing it', async (code) => {
    const context: StreamConversionContext = {
      advertisedTools: { exec: 'custom' },
      usedCallIds: [],
      allowedExecMethods: ['mcp__wiswork__read_document'],
    }
    await expectStreamCode(
      [
        start,
        'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"c","name":"exec","input":{}}}\n\n',
        `event: content_block_delta\ndata: ${JSON.stringify({ type: 'content_block_delta', index: 0, delta: { type: 'input_json_delta', partial_json: JSON.stringify({ code }) } })}\n\n`,
        'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n',
      ],
      'unsafe_custom_tool_input',
      context,
    )
  })
})
