import { describe, expect, it } from 'vitest'
import * as bridge from '../src/index.js'
import captured from './fixtures/codex-0147-request.json'
import { prepareCarrierTurn } from './fixtures/carrier-authorization.js'

const clone = (): Record<string, any> => structuredClone(captured)
const prepareResponsesTurn = (input: unknown, limits: Record<string, number> = {}) =>
  prepareCarrierTurn(input, limits)

async function* chunks(...values: string[]): AsyncGenerator<string> {
  yield* values
}

const start =
  'event: message_start\ndata: {"type":"message_start","message":{"id":"msg_1","model":"openai/gpt-5.6-sol","usage":{"input_tokens":1}}}\n\n'
const delta =
  'event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"end_turn"},"usage":{"output_tokens":1}}\n\n'
const stop = 'event: message_stop\ndata: {"type":"message_stop"}\n\n'

async function consume(source: AsyncIterable<string>): Promise<void> {
  for await (const _frame of source) void _frame
}

describe('final pinned bridge contract', () => {
  it.each(['exec_command', 'apply_patch'])('rejects top-level function tool %s', (name) => {
    const body = clone()
    body.tools = [{ type: 'function', name, parameters: {} }]
    expect(() => prepareResponsesTurn(body)).toThrowError('unsupported_top_level_tools')
  })

  it('exposes only a frozen bound-turn API', () => {
    expect('responsesToMessages' in bridge).toBe(false)
    expect('responsesToMessagesWithContext' in bridge).toBe(false)
    expect('messagesSseToResponses' in bridge).toBe(false)
    const turn = prepareResponsesTurn(clone())
    expect(Object.isFrozen(turn)).toBe(true)
    expect(Object.keys(turn).sort()).toEqual(['messagesRequest', 'messagesStreamToResponses'])
    expect(turn.messagesStreamToResponses).toHaveLength(1)
  })

  it('binds request limits to the stream converter', async () => {
    const turn = prepareResponsesTurn(clone(), { maxSseFrames: 2 })
    await expect(
      consume(turn.messagesStreamToResponses(chunks(start, delta, stop))),
    ).rejects.toThrow('sse_frame_count_limit_exceeded')
  })

  it('accepts the faithful MCP turn metadata shape', () => {
    const turn = prepareResponsesTurn(clone())
    expect(turn.messagesRequest.tools).toHaveLength(1)
    expect(turn.messagesRequest.tools![0].name).toBe('exec')
  })

  it('rejects malformed workspace metadata', () => {
    const body = clone()
    const packed = JSON.parse(body.client_metadata['x-codex-turn-metadata'])
    packed.workspaces['/workspace'].associated_remote_urls.extra = 7
    body.client_metadata['x-codex-turn-metadata'] = JSON.stringify(packed)
    expect(() => prepareResponsesTurn(body)).toThrowError('unsupported_client_metadata')
  })

  it('fails a deeply nested request with a fixed compatibility error, not RangeError', () => {
    const body = clone()
    let cursor: Record<string, unknown> = body
    for (let index = 0; index < 20_000; index += 1) {
      const child: Record<string, unknown> = {}
      cursor.deep = child
      cursor = child
    }
    expect(() => prepareResponsesTurn(body)).toThrowError('request_nesting_limit_exceeded')
  })

  it('enforces aggregate request bytes and nodes', () => {
    expect(() => prepareResponsesTurn(clone(), { maxRequestBytes: 32 })).toThrowError(
      'request_bytes_limit_exceeded',
    )
    expect(() => prepareResponsesTurn(clone(), { maxRequestNodes: 3 })).toThrowError(
      'request_nodes_limit_exceeded',
    )
  })

  it.each([
    '// @exec: {"yield_time_ms":1000}\ntext(await tools.mcp__wiswork__wiswork_read_document({"id":"1"}))',
    '// @exec: {"yield_time_ms":1000,"max_output_tokens":100}\nawait tools.mcp__wiswork__wiswork_read_document({})',
  ])('accepts one bounded exec pragma', async (code) => {
    const turn = prepareResponsesTurn(clone())
    const toolStart =
      'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"c","name":"exec","input":{},"caller":{"type":"direct"}}}\n\n'
    const toolDelta = `event: content_block_delta\ndata: ${JSON.stringify({ type: 'content_block_delta', index: 0, delta: { type: 'input_json_delta', partial_json: JSON.stringify({ code }) } })}\n\n`
    const toolStop = 'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n'
    await expect(
      consume(
        turn.messagesStreamToResponses(chunks(start, toolStart, toolDelta, toolStop, delta, stop)),
      ),
    ).resolves.toBeUndefined()
  })

  it('rejects an unknown production tool caller variant', async () => {
    const turn = prepareResponsesTurn(clone())
    await expect(
      consume(
        turn.messagesStreamToResponses(
          chunks(
            start,
            'data: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"c","name":"exec","input":{},"caller":{"type":"indirect"}}}\n\n',
          ),
        ),
      ),
    ).rejects.toThrow('unsupported_messages_event')
  })

  it('accepts a grammar-valid pragma with leading horizontal whitespace', async () => {
    const code =
      ' \t// @exec: {"yield_time_ms":1000}\nawait tools.mcp__wiswork__wiswork_read_document({})'
    const turn = prepareResponsesTurn(clone())
    const toolStart =
      'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"c","name":"exec","input":{}}}\n\n'
    const toolDelta = `event: content_block_delta\ndata: ${JSON.stringify({ type: 'content_block_delta', index: 0, delta: { type: 'input_json_delta', partial_json: JSON.stringify({ code }) } })}\n\n`
    const toolStop = 'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n'
    await expect(
      consume(
        turn.messagesStreamToResponses(chunks(start, toolStart, toolDelta, toolStop, delta, stop)),
      ),
    ).resolves.toBeUndefined()
  })

  it('accepts standard Anthropic output_tokens at message_start and ping events', async () => {
    const standardStart =
      'event: message_start\ndata: {"type":"message_start","message":{"id":"msg_1","type":"message","role":"assistant","content":[],"model":"openai/gpt-5.6-sol","stop_reason":null,"stop_sequence":null,"usage":{"input_tokens":1,"output_tokens":0}}}\n\n'
    const ping = 'event: ping\ndata: {"type":"ping"}\n\n'
    const turn = prepareResponsesTurn(clone())
    await expect(
      consume(turn.messagesStreamToResponses(chunks(ping, standardStart, ping, delta, ping, stop))),
    ).resolves.toBeUndefined()
  })

  it('accepts WisUsage data-only SSE and an omitted terminal usage object', async () => {
    const turn = prepareResponsesTurn(clone())
    const output: string[] = []
    for await (const frame of turn.messagesStreamToResponses(
      chunks(
        'data: {"type":"message_start","message":{"id":"msg_1","model":"openai/gpt-5.6-sol","usage":{"input_tokens":1}}}\n\n',
        'data: {"type":"content_block_start","index":0,"content_block":{"type":"text","text":""}}\n\n',
        'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"ok"}}\n\n',
        'data: {"type":"content_block_stop","index":0}\n\n',
        'data: {"type":"message_delta","delta":{"stop_reason":"end_turn"}}\n\n',
        'data: {"type":"message_stop"}\n\n',
        'event: data\ndata: [DONE]\n\n',
      ),
    )) {
      output.push(frame)
    }
    expect(output.join('')).toContain('event: response.completed')
    expect(output.join('')).toContain('"output_tokens":0')
  })

  it('accepts the bounded production WisUsage metadata envelope without forwarding it', async () => {
    const turn = prepareResponsesTurn(clone())
    const output: string[] = []
    for await (const frame of turn.messagesStreamToResponses(
      chunks(
        'data: {"type":"message_start","message":{"id":"msg_1","type":"message","role":"assistant","container":null,"content":[],"model":"openai/gpt-5.6-sol","stop_reason":null,"stop_details":null,"stop_sequence":null,"usage":{"input_tokens":12,"output_tokens":0,"output_tokens_details":null,"cache_creation_input_tokens":null,"cache_read_input_tokens":null,"cache_creation":null,"inference_geo":null,"server_tool_use":null,"service_tier":null,"speed":"standard"},"provider":"openai"}}\n\n',
        'data: {"type":"content_block_start","index":0,"content_block":{"type":"text","text":"","citations":[]}}\n\n',
        'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"ok"}}\n\n',
        'data: {"type":"content_block_stop","index":0}\n\n',
        'data: {"type":"message_delta","delta":{"container":null,"stop_details":null,"stop_reason":"end_turn","stop_sequence":null},"usage":{"input_tokens":12,"output_tokens":2,"output_tokens_details":{"thinking_tokens":0},"cache_creation_input_tokens":null,"cache_read_input_tokens":0,"cache_creation":null,"server_tool_use":null,"service_tier":"standard","speed":"standard","cost":0.001,"is_byok":false,"cost_details":{"upstream_inference_cost":0.001,"upstream_inference_prompt_cost":0.0004,"upstream_inference_completions_cost":0.0006}},"context_management":null}\n\n',
        'data: {"type":"message_stop"}\n\n',
      ),
    )) {
      output.push(frame)
    }
    const serialized = output.join('')
    expect(serialized).toContain('event: response.completed')
    expect(serialized).toContain('"output_tokens":2')
    expect(serialized).not.toContain('provider')
    expect(serialized).not.toContain('cost_details')
  })

  it('still rejects a mismatched explicit SSE event name', async () => {
    const turn = prepareResponsesTurn(clone())
    await expect(
      consume(
        turn.messagesStreamToResponses(
          chunks(
            'event: message_stop\ndata: {"type":"message_start","message":{"id":"msg_1","model":"openai/gpt-5.6-sol","usage":{"input_tokens":1}}}\n\n',
          ),
        ),
      ),
    ).rejects.toThrow('invalid_messages_sse')
  })

  it('accepts a bounded batch of MCP calls in one response', async () => {
    const code = 'await tools.mcp__wiswork__wiswork_read_document({})'
    const call = (index: number, id: string) =>
      `event: content_block_start\ndata: ${JSON.stringify({ type: 'content_block_start', index, content_block: { type: 'tool_use', id, name: 'exec', input: {} } })}\n\n` +
      `event: content_block_delta\ndata: ${JSON.stringify({ type: 'content_block_delta', index, delta: { type: 'input_json_delta', partial_json: JSON.stringify({ code }) } })}\n\n` +
      `event: content_block_stop\ndata: ${JSON.stringify({ type: 'content_block_stop', index })}\n\n`
    const turn = prepareResponsesTurn(clone())
    await expect(
      consume(
        turn.messagesStreamToResponses(
          chunks(
            start,
            call(0, 'c1'),
            call(1, 'c2'),
            'event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"tool_use"},"usage":{"output_tokens":1}}\n\n',
            'event: message_stop\ndata: {"type":"message_stop"}\n\n',
          ),
        ),
      ),
    ).resolves.toBeUndefined()
  })

  it('rejects an MCP response above the bounded batch limit', async () => {
    const code = 'await tools.mcp__wiswork__wiswork_read_document({})'
    const call = (index: number) =>
      `event: content_block_start\ndata: ${JSON.stringify({ type: 'content_block_start', index, content_block: { type: 'tool_use', id: `c${index}`, name: 'exec', input: {} } })}\n\n` +
      `event: content_block_delta\ndata: ${JSON.stringify({ type: 'content_block_delta', index, delta: { type: 'input_json_delta', partial_json: JSON.stringify({ code }) } })}\n\n` +
      `event: content_block_stop\ndata: ${JSON.stringify({ type: 'content_block_stop', index })}\n\n`
    const turn = prepareResponsesTurn(clone())
    await expect(
      consume(
        turn.messagesStreamToResponses(
          chunks(start, ...Array.from({ length: 17 }, (_, index) => call(index))),
        ),
      ),
    ).rejects.toThrow('tool_call_limit_exceeded')
  })

  it('accepts a bounded historical assistant tool batch', () => {
    const body = clone()
    const code = 'await tools.mcp__wiswork__wiswork_read_document({})'
    body.input.push(
      { type: 'custom_tool_call', call_id: 'c1', name: 'exec', input: code },
      { type: 'custom_tool_call', call_id: 'c2', name: 'exec', input: code },
    )
    body.input.push(
      { type: 'custom_tool_call_output', call_id: 'c1', output: 'one' },
      { type: 'custom_tool_call_output', call_id: 'c2', output: 'two' },
    )
    expect(() => prepareResponsesTurn(body)).not.toThrow()
  })

  it.each([
    '// @exec: []\nawait tools.mcp__wiswork__wiswork_read_document({})',
    '// @exec: {"unknown":1}\nawait tools.mcp__wiswork__wiswork_read_document({})',
    '// @exec: {"yield_time_ms":0}\nawait tools.mcp__wiswork__wiswork_read_document({})',
    '// @exec: {"yield_time_ms":1,"yield_time_ms":2}\nawait tools.mcp__wiswork__wiswork_read_document({})',
    '// @exec: {}\n// @exec: {}\nawait tools.mcp__wiswork__wiswork_read_document({})',
  ])('rejects unsafe pragma without echoing code', async (code) => {
    const body = clone()
    body.input.push({ type: 'custom_tool_call', call_id: 'c', name: 'exec', input: code })
    try {
      prepareResponsesTurn(body)
      throw new Error('expected failure')
    } catch (error) {
      expect(error).toBeInstanceOf(bridge.ProtocolCompatibilityError)
      expect(String(error)).not.toContain(code)
    }
  })

  it('rejects extra custom call and result fields', () => {
    const call = clone()
    call.input.push({
      type: 'custom_tool_call',
      call_id: 'c',
      name: 'exec',
      input: 'await tools.mcp__wiswork__wiswork_read_document({})',
      extra: 'secret',
    })
    expect(() => prepareResponsesTurn(call)).toThrowError('unsupported_input_item')

    const result = clone()
    result.input.push(
      {
        type: 'custom_tool_call',
        call_id: 'c',
        name: 'exec',
        input: 'await tools.mcp__wiswork__wiswork_read_document({})',
      },
      { type: 'custom_tool_call_output', call_id: 'c', output: 'ok', extra: 'secret' },
    )
    expect(() => prepareResponsesTurn(result)).toThrowError('unsupported_input_item')
  })

  it.each([
    [
      'message_start extra',
      'event: message_start\ndata: {"type":"message_start","extra":1,"message":{"id":"msg_1","model":"openai/gpt-5.6-sol","usage":{"input_tokens":1}}}\n\n',
    ],
    [
      'message usage extra',
      'event: message_start\ndata: {"type":"message_start","message":{"id":"msg_1","model":"openai/gpt-5.6-sol","usage":{"input_tokens":1,"extra":1}}}\n\n',
    ],
    [
      'content block extra',
      start +
        'event: content_block_start\ndata: {"type":"content_block_start","index":0,"extra":1,"content_block":{"type":"text","text":""}}\n\n',
    ],
    [
      'content delta extra',
      start +
        'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"text","text":""}}\n\n' +
        'event: content_block_delta\ndata: {"type":"content_block_delta","index":0,"extra":1,"delta":{"type":"text_delta","text":"x"}}\n\n',
    ],
    [
      'message delta usage extra',
      start +
        'event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"end_turn"},"usage":{"output_tokens":1,"extra":1}}\n\n',
    ],
    [
      'message stop extra',
      start + delta + 'event: message_stop\ndata: {"type":"message_stop","extra":1}\n\n',
    ],
  ])('rejects exact-shape SSE violation: %s', async (_label, frame) => {
    const turn = prepareResponsesTurn(clone())
    await expect(consume(turn.messagesStreamToResponses(chunks(frame)))).rejects.toThrow(
      'invalid_messages_event',
    )
  })
})
