import { describe, expect, it } from 'vitest'
import {
  ProtocolRecorder,
  parseProtocolRecording,
  replayProtocolRecording,
} from '../src/protocol-recording.js'
import { prepareResponsesTurn } from '../src/index.js'
import redacted from './fixtures/protocol-redacted-max-tokens.json'
import incomplete from './fixtures/protocol-incomplete-tool.json'

const frame = (data: unknown) => `data: ${JSON.stringify(data)}\n\n`
const start = {
  type: 'message_start',
  message: { id: 'private-id', model: 'openai/gpt-5.6-sol', usage: { input_tokens: 1 } },
}
async function* stream(values: unknown[]) {
  for (const value of values) yield frame(value)
}

describe('structural protocol recording', () => {
  it('runs fixtures through real protocol logic, including changed stop reason and missing terminal', async () => {
    expect((await replayProtocolRecording(redacted)).events).toContain('response.incomplete')
    expect((await replayProtocolRecording(incomplete)).error).toBe('invalid_custom_tool_input')
    const limited = structuredClone(incomplete)
    limited.frames[4]!.delta!.stop_reason = 'max_tokens'
    expect((await replayProtocolRecording(limited)).events).toContain('response.incomplete')
    expect(
      (await replayProtocolRecording({ ...redacted, frames: redacted.frames.slice(0, -1) })).error,
    ).toBe('premature_messages_eof')
  })
  it('records the production parser path and replays redacted reasoning with max_tokens', async () => {
    const recorder = new ProtocolRecorder()
    const turn = prepareResponsesTurn({ model: 'gpt-5.6-sol', input: 'private prompt' })
    for await (const _ of turn.messagesStreamToResponses(
      stream([
        start,
        {
          type: 'content_block_start',
          index: 0,
          content_block: { type: 'redacted_thinking', data: 'SECRET JWT encrypted thinking' },
        },
        { type: 'content_block_stop', index: 0 },
        { type: 'message_delta', delta: { stop_reason: 'max_tokens' } },
        { type: 'message_stop' },
      ]),
      recorder,
    )) {
      /* consume */
    }
    const recording = recorder.snapshot()
    expect(recording.frames).toHaveLength(5)
    expect(JSON.stringify(recording)).not.toMatch(/SECRET|private|JWT|encrypted thinking/)
    const result = await replayProtocolRecording(recording)
    expect(result.events).toContain('response.incomplete')
    expect(result.error).toBeUndefined()
  })
  it('sanitizes unknown nested keys and values without accepting them on import', () => {
    const recorder = new ProtocolRecorder()
    recorder.recordFrame(
      frame({
        type: 'message_start',
        message: { ...start.message, '/private/secret': { jwt: 'SECRET' } },
      }),
    )
    expect(JSON.stringify(recorder.snapshot())).not.toMatch(/private|secret|jwt|SECRET/)
    expect(() => parseProtocolRecording({ ...recorder.snapshot(), jwt: 'SECRET' })).toThrow(
      'invalid_protocol_recording',
    )
    expect(() =>
      parseProtocolRecording({
        ...recorder.snapshot(),
        frames: [{ type: 'ping', payload: 'SECRET' }],
      }),
    ).toThrow('invalid_protocol_recording')
  })
  it('replays malformed stream and bounds capture without evicting its initial state', async () => {
    const recorder = new ProtocolRecorder()
    recorder.recordFrame('data: not json')
    expect((await replayProtocolRecording(recorder.snapshot())).error).toBe('invalid_messages_sse')
    for (let i = 0; i < 300; i++) recorder.recordFrame(frame({ type: 'ping' }))
    expect(recorder.snapshot().frames).toHaveLength(256)
    expect(recorder.snapshot().truncated).toBe(true)
  })
})
