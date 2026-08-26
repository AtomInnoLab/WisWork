import { describe, expect, it } from 'vitest'
import { prepareResponsesTurn } from '../src/index.js'
import captured from './fixtures/codex-0147-request.json'

const noToolTurn = () => prepareResponsesTurn({ model: 'gpt-5.6-sol', input: 'Hello' })
const customToolTurn = () => prepareResponsesTurn(structuredClone(captured))

function convert(
  source: AsyncIterable<string | Uint8Array>,
  custom = false,
): AsyncGenerator<string> {
  return (custom ? customToolTurn() : noToolTurn()).messagesStreamToResponses(source)
}

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
      convert(
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

  it('converts exec tool use to custom tool call events and raw input', async () => {
    const events = await collect(
      convert(
        chunks(
          messageStart,
          'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"custom_7","name":"exec","input":{}}}\n\n',
          'event: content_block_delta\ndata: {"type":"content_block_delta","index":0,"delta":{"type":"input_json_delta","partial_json":"{\\"code\\":\\"text(await tools.mcp__wiswork__wiswork_read_document({}))\\"}"}}\n\n',
          'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n',
          'event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"tool_use"},"usage":{"output_tokens":4}}\n\n',
          'event: message_stop\ndata: {"type":"message_stop"}\n\n',
        ),
        true,
      ),
    )

    expect(events.find(({ event }) => event === 'response.output_item.added')?.data).toMatchObject({
      item: { type: 'custom_tool_call', call_id: 'custom_7', name: 'exec', input: '' },
    })
    expect(
      events.find(({ event }) => event === 'response.custom_tool_call_input.delta')?.data,
    ).toMatchObject({
      delta: 'text(await tools.mcp__wiswork__wiswork_read_document({}))',
    })
    expect(
      events.find(({ event }) => event === 'response.custom_tool_call_input.done')?.data,
    ).toMatchObject({
      input: 'text(await tools.mcp__wiswork__wiswork_read_document({}))',
    })
    expect(events.find(({ event }) => event === 'response.output_item.done')?.data).toMatchObject({
      item: {
        type: 'custom_tool_call',
        call_id: 'custom_7',
        name: 'exec',
        input: 'text(await tools.mcp__wiswork__wiswork_read_document({}))',
      },
    })
  })

  it('fails closed on malformed exec input without echoing code', async () => {
    let caught: unknown
    try {
      await collect(
        convert(
          chunks(
            messageStart,
            'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"custom_7","name":"exec","input":{}}}\n\n',
            'event: content_block_delta\ndata: {"type":"content_block_delta","index":0,"delta":{"type":"input_json_delta","partial_json":"{\\"code\\":\\"secret prompt\\",\\"extra\\":true}"}}\n\n',
            'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n',
          ),
          true,
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

    const events = await collect(convert(byteChunks()))
    expect(events.at(-1)?.data).toMatchObject({
      response: { status: 'incomplete', incomplete_details: { reason: 'max_output_tokens' } },
    })
  })

  it('terminates a completed Responses stream with [DONE]', async () => {
    const output: string[] = []
    for await (const frame of convert(
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
      'malformed ping event',
      'event: ping\ndata: {"type":"ping","message":"secret prompt"}\n\n',
      'invalid_messages_event',
    ],
  ])('fails closed for %s without echoing upstream data', async (_label, frame, code) => {
    let caught: unknown
    try {
      await collect(convert(chunks(frame)))
    } catch (error) {
      caught = error
    }
    expect(caught).toEqual(expect.objectContaining({ message: code }))
    expect(String(caught)).not.toContain('secret prompt')
  })
})
