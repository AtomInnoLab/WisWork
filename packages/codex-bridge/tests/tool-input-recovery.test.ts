import { runInNewContext } from 'node:vm'
import { describe, expect, it, vi } from 'vitest'
import { prepareResponsesTurn } from '../src/index.js'
import captured from './fixtures/codex-0147-request.json'
import { prepareCarrierTurn } from './fixtures/carrier-authorization.js'

const safeCode = 'text(await tools.mcp__wiswork__wiswork_read_document({}))'
const frame = (value: object) => `data: ${JSON.stringify(value)}\n\n`

async function* response(inputs: string[], stopReason = 'tool_use') {
  yield frame({
    type: 'message_start',
    message: { id: 'response-1', model: 'openai/gpt-5.6-sol', usage: { input_tokens: 1 } },
  })
  for (const [index, input] of inputs.entries()) {
    yield frame({
      type: 'content_block_start',
      index,
      content_block: { type: 'tool_use', id: `new-${index}`, name: 'exec', input: {} },
    })
    yield frame({
      type: 'content_block_delta',
      index,
      delta: { type: 'input_json_delta', partial_json: input },
    })
    yield frame({ type: 'content_block_stop', index })
  }
  yield frame({ type: 'message_delta', delta: { stop_reason: stopReason } })
  yield frame({ type: 'message_stop' })
}

async function collect(source: AsyncIterable<string>) {
  const events: Record<string, any>[] = []
  for await (const value of source) {
    const data = /^data: (\{.+\})$/m.exec(value)?.[1]
    if (data) events.push(JSON.parse(data))
  }
  return events
}

const calls = (events: Record<string, any>[]) =>
  events.filter((event) => event.type === 'response.output_item.done').map((event) => event.item)

function recoveryResult(code: string) {
  const text = vi.fn()
  // No tools, process or other application authority: a recovery call must only emit feedback.
  runInNewContext(
    code,
    { text },
    { timeout: 100, contextCodeGeneration: { strings: false, wasm: false } },
  )
  expect(text).toHaveBeenCalledOnce()
  return text.mock.calls[0]![0]
}

function appendHistory(request: { input: any[] }, code: string, id: string) {
  request.input.push(
    { type: 'custom_tool_call', call_id: id, name: 'exec', input: code },
    { type: 'custom_tool_call_output', call_id: id, output: 'Tool output' },
  )
}

describe('recoverable exec input errors', () => {
  it.each([
    ['malformed JSON', '{"code":"SECRET'],
    ['missing code', '{"unexpected":"SECRET"}'],
    ['non-string code', '{"code":7}'],
    [
      'unsupported wrapper',
      JSON.stringify({
        code: 'const r = await tools.mcp__wiswork__wiswork_read_document({}); text(r);',
      }),
    ],
    [
      'invalid JSON literal',
      JSON.stringify({
        code: 'text(await tools.mcp__wiswork__wiswork_read_document({token:"SECRET"}))',
      }),
    ],
    ['unapproved code', JSON.stringify({ code: 'process.exit(); // SECRET' })],
  ])('returns actionable feedback without dispatching %s', async (_label, input) => {
    const events = await collect(
      prepareCarrierTurn(structuredClone(captured)).messagesStreamToResponses(response([input])),
    )
    expect(events.at(-1)?.type).toBe('response.completed')
    const [call] = calls(events)
    expect(call).toMatchObject({ type: 'custom_tool_call', call_id: 'new-0', name: 'exec' })
    expect(call.input).not.toContain('SECRET')
    const result = recoveryResult(call.input)
    expect(result).toMatchObject({ isError: true, error: 'invalid_tool_input' })
    expect(result.message).toContain('No document tool was executed')
    expect(result.message).toContain('JSON object')
  })

  it('keeps rejected-call feedback in history and accepts the corrected call on the next request', async () => {
    const request = structuredClone(captured)
    const rejected = calls(
      await collect(prepareCarrierTurn(request).messagesStreamToResponses(response(['{']))),
    )[0]
    appendHistory(request, rejected.input, 'rejected-1')
    const next = prepareCarrierTurn(request)
    expect(JSON.stringify(next.messagesRequest.messages)).toContain('No document tool was executed')
    const events = await collect(
      next.messagesStreamToResponses(response([JSON.stringify({ code: safeCode })])),
    )
    expect(calls(events)[0]).toMatchObject({ input: safeCode })
    expect(events.at(-1)?.type).toBe('response.completed')
  })

  it('preserves IDs and valid calls in a batch containing a rejected call', async () => {
    const events = await collect(
      prepareCarrierTurn(structuredClone(captured)).messagesStreamToResponses(
        response([JSON.stringify({ code: safeCode }), '{', JSON.stringify({ code: safeCode })]),
      ),
    )
    const output = calls(events)
    expect(output.map((call) => call.call_id)).toEqual(['new-0', 'new-1', 'new-2'])
    expect(output[0].input).toBe(safeCode)
    expect(recoveryResult(output[1].input).isError).toBe(true)
    expect(output[2].input).toBe(safeCode)
  })

  it('bounds recovery across requests and only resets the budget on a new user message', async () => {
    const request = structuredClone(captured)
    for (let index = 0; index < 3; index++) {
      const rejected = calls(
        await collect(prepareCarrierTurn(request).messagesStreamToResponses(response(['{']))),
      )[0]
      appendHistory(request, rejected.input, `rejected-${index}`)
      // A valid call and its result are not a new user instruction and cannot reset the cap.
      appendHistory(request, safeCode, `read-${index}`)
    }
    await expect(
      collect(prepareCarrierTurn(request).messagesStreamToResponses(response(['{']))),
    ).rejects.toThrow('tool_input_retry_limit_exceeded')
    request.input.push({
      type: 'message',
      role: 'user',
      content: [{ type: 'input_text', text: 'Continue' }],
    } as any)
    const events = await collect(
      prepareCarrierTurn(request).messagesStreamToResponses(response(['{'])),
    )
    expect(recoveryResult(calls(events)[0].input).isError).toBe(true)
  })

  it('does not dispatch a batch exceeding the recovery budget', async () => {
    const events: Record<string, any>[] = []
    await expect(
      (async () => {
        for await (const output of prepareCarrierTurn(
          structuredClone(captured),
        ).messagesStreamToResponses(response(['{', '{', '{', '{']))) {
          const data = /^data: (\{.+\})$/m.exec(output)?.[1]
          if (data) events.push(JSON.parse(data))
        }
      })(),
    ).rejects.toThrow('tool_input_retry_limit_exceeded')
    expect(calls(events)).toEqual([])
  })

  it('does not broaden the history validator to arbitrary or appended code', async () => {
    const rejected = calls(
      await collect(
        prepareCarrierTurn(structuredClone(captured)).messagesStreamToResponses(response(['{'])),
      ),
    )[0]
    for (const code of [rejected.input + '; text("extra")', 'text({"isError":true})']) {
      const request = structuredClone(captured)
      appendHistory(request, code, 'bad-history')
      expect(() => prepareCarrierTurn(request)).toThrow('unsafe_custom_tool_input')
    }
  })

  it('still rejects missing authority and oversized arguments and discards truncated calls', async () => {
    await expect(
      collect(
        prepareResponsesTurn({ model: 'gpt-5.6-sol', input: 'hello' }).messagesStreamToResponses(
          response(['{']),
        ),
      ),
    ).rejects.toThrow('unadvertised_tool_call')
    await expect(
      collect(
        prepareCarrierTurn(structuredClone(captured), {
          maxToolArguments: 10,
        }).messagesStreamToResponses(response(['x'.repeat(11)])),
      ),
    ).rejects.toThrow('tool_arguments_limit_exceeded')
    await expect(
      collect(
        prepareCarrierTurn(structuredClone(captured), {
          maxToolArguments: 10,
        }).messagesStreamToResponses(response(['{'])),
      ),
    ).rejects.toThrow('tool_arguments_limit_exceeded')
    const events = await collect(
      prepareCarrierTurn(structuredClone(captured)).messagesStreamToResponses(
        response(['{'], 'max_tokens'),
      ),
    )
    expect(calls(events)).toEqual([])
    expect(events.at(-1)?.type).toBe('response.incomplete')
  })
})
