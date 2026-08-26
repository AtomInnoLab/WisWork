import { describe, expect, it } from 'vitest'
import { messagesSseToResponses } from '../src/index.js'

async function* chunks(...values: string[]): AsyncGenerator<string> {
  yield* values
}

async function collect(
  source: AsyncIterable<string>,
): Promise<Array<{ event: string; data: unknown }>> {
  const frames: Array<{ event: string; data: unknown }> = []
  for await (const output of source) {
    const event = /^event: ([^\n]+)$/m.exec(output)?.[1]
    const data = /^data: (.+)$/m.exec(output)?.[1]
    if (event && data) frames.push({ event, data: JSON.parse(data) })
  }
  return frames
}

const messageStart =
  'event: message_start\ndata: {"type":"message_start","message":{"id":"msg_1","model":"openai/gpt-5.6-sol","usage":{"input_tokens":10,"cache_read_input_tokens":4,"cache_creation_input_tokens":2}}}\n\n'

describe('messagesSseToResponses', () => {
  it('converts text, usage, and end-to-end completion state', async () => {
    const events = await collect(
      messagesSseToResponses(
        chunks(
          messageStart,
          'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"text","text":""}}\n\n',
          'event: content_block_delta\ndata: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"Hello"}}\n\n',
          'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n',
          'event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"end_turn"},"usage":{"output_tokens":3}}\n\n',
          'event: message_stop\ndata: {"type":"message_stop"}\n\n',
        ),
      ),
    )

    expect(events.map(({ event }) => event)).toEqual([
      'response.created',
      'response.output_item.added',
      'response.content_part.added',
      'response.output_text.delta',
      'response.output_text.done',
      'response.content_part.done',
      'response.output_item.done',
      'response.completed',
    ])
    expect(events.find(({ event }) => event === 'response.output_text.delta')?.data).toMatchObject({
      delta: 'Hello',
    })
    expect(events.at(-1)?.data).toMatchObject({
      response: {
        status: 'completed',
        usage: {
          input_tokens: 16,
          output_tokens: 3,
          total_tokens: 19,
          input_tokens_details: { cached_tokens: 4 },
        },
      },
    })
  })

  it('preserves tool call IDs and argument delta order', async () => {
    const events = await collect(
      messagesSseToResponses(
        chunks(
          messageStart,
          'event: content_block_start\ndata: {"type":"content_block_start","index":1,"content_block":{"type":"tool_use","id":"call_9","name":"edit","input":{}}}\n\n',
          'event: content_block_delta\ndata: {"type":"content_block_delta","index":1,"delta":{"type":"input_json_delta","partial_json":"{\\"a\\":"}}\n\n',
          'event: content_block_delta\ndata: {"type":"content_block_delta","index":1,"delta":{"type":"input_json_delta","partial_json":"1}"}}\n\n',
          'event: content_block_stop\ndata: {"type":"content_block_stop","index":1}\n\n',
          'event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"tool_use"},"usage":{"output_tokens":7}}\n\n',
          'event: message_stop\ndata: {"type":"message_stop"}\n\n',
        ),
      ),
    )

    expect(events.find(({ event }) => event === 'response.output_item.added')?.data).toMatchObject({
      item: { type: 'function_call', call_id: 'call_9', name: 'edit' },
    })
    expect(
      events
        .filter(({ event }) => event === 'response.function_call_arguments.delta')
        .map(({ data }) => (data as { delta: string }).delta),
    ).toEqual(['{"a":', '1}'])
    expect(
      events.find(({ event }) => event === 'response.function_call_arguments.done')?.data,
    ).toMatchObject({ item_id: 'item_1', arguments: '{"a":1}' })
    expect(events.find(({ event }) => event === 'response.output_item.done')?.data).toMatchObject({
      item: { call_id: 'call_9', arguments: '{"a":1}' },
    })
    expect(events.at(-1)?.data).toMatchObject({ response: { status: 'completed' } })
  })

  it('converts exec tool use to custom tool call events and raw input', async () => {
    const events = await collect(
      messagesSseToResponses(
        chunks(
          messageStart,
          'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"custom_7","name":"exec","input":{}}}\n\n',
          'event: content_block_delta\ndata: {"type":"content_block_delta","index":0,"delta":{"type":"input_json_delta","partial_json":"{\\"code\\":\\"text(42)\\"}"}}\n\n',
          'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n',
          'event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"tool_use"},"usage":{"output_tokens":4}}\n\n',
          'event: message_stop\ndata: {"type":"message_stop"}\n\n',
        ),
      ),
    )

    expect(events.find(({ event }) => event === 'response.output_item.added')?.data).toMatchObject({
      item: { type: 'custom_tool_call', call_id: 'custom_7', name: 'exec', input: '' },
    })
    expect(
      events.find(({ event }) => event === 'response.custom_tool_call_input.delta')?.data,
    ).toMatchObject({
      delta: 'text(42)',
    })
    expect(
      events.find(({ event }) => event === 'response.custom_tool_call_input.done')?.data,
    ).toMatchObject({
      input: 'text(42)',
    })
    expect(events.find(({ event }) => event === 'response.output_item.done')?.data).toMatchObject({
      item: { type: 'custom_tool_call', call_id: 'custom_7', name: 'exec', input: 'text(42)' },
    })
  })

  it('fails closed on malformed exec input without echoing code', async () => {
    let caught: unknown
    try {
      await collect(
        messagesSseToResponses(
          chunks(
            messageStart,
            'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"custom_7","name":"exec","input":{}}}\n\n',
            'event: content_block_delta\ndata: {"type":"content_block_delta","index":0,"delta":{"type":"input_json_delta","partial_json":"{\\"code\\":\\"secret prompt\\",\\"extra\\":true}"}}\n\n',
            'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n',
          ),
        ),
      )
    } catch (error) {
      caught = error
    }
    expect(caught).toEqual(expect.objectContaining({ message: 'invalid_custom_tool_input' }))
    expect(String(caught)).not.toContain('secret prompt')
  })

  it('handles SSE frames split at arbitrary byte boundaries', async () => {
    const encoded = new TextEncoder().encode(
      messageStart +
        'event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"max_tokens"},"usage":{"output_tokens":1}}\n\n' +
        'event: message_stop\ndata: {"type":"message_stop"}\n\n',
    )
    async function* byteChunks(): AsyncGenerator<Uint8Array> {
      for (const byte of encoded) yield new Uint8Array([byte])
    }

    const events = await collect(messagesSseToResponses(byteChunks()))
    expect(events.at(-1)?.data).toMatchObject({
      response: { status: 'incomplete', incomplete_details: { reason: 'max_output_tokens' } },
    })
  })

  it('terminates a completed Responses stream with [DONE]', async () => {
    const output: string[] = []
    for await (const frame of messagesSseToResponses(
      chunks(
        messageStart,
        'event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"end_turn"},"usage":{"output_tokens":0}}\n\n',
        'event: message_stop\ndata: {"type":"message_stop"}\n\n',
      ),
    )) {
      output.push(frame)
    }
    expect(output.at(-1)).toBe('data: [DONE]\n\n')
  })

  it.each([
    [
      'upstream error',
      'event: error\ndata: {"type":"error","error":{"type":"overloaded_error","message":"secret prompt"}}\n\n',
      'upstream_error',
    ],
    [
      'unknown event variant',
      'event: ping\ndata: {"type":"ping","message":"secret prompt"}\n\n',
      'unsupported_messages_event',
    ],
  ])('fails closed for %s without echoing upstream data', async (_label, frame, code) => {
    let caught: unknown
    try {
      await collect(messagesSseToResponses(chunks(frame)))
    } catch (error) {
      caught = error
    }
    expect(caught).toEqual(expect.objectContaining({ message: code }))
    expect(String(caught)).not.toContain('secret prompt')
  })
})
