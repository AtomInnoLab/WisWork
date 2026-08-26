import { describe, expect, it } from 'vitest'
import captured from './fixtures/codex-0147-request.json'
import { ProtocolCompatibilityError, responsesToMessagesWithContext } from '../src/index.js'

const cloneCaptured = (): Record<string, unknown> =>
  structuredClone(captured) as Record<string, unknown>

function expectCode(request: unknown, code: string): void {
  expect(() => responsesToMessagesWithContext(request)).toThrowError(
    expect.objectContaining<Partial<ProtocolCompatibilityError>>({ message: code }),
  )
}

describe('strict Responses request conversion', () => {
  it('accepts the captured Codex 0.147 envelope and returns stream context', () => {
    const converted = responsesToMessagesWithContext(cloneCaptured())
    expect(converted.request).toMatchObject({
      model: 'openai/gpt-5.6-sol',
      system: 'System',
      messages: [{ role: 'user', content: [{ type: 'text', text: 'Hello' }] }],
      tool_choice: { type: 'auto', disable_parallel_tool_use: true },
    })
    expect(converted.context).toEqual({
      advertisedTools: { exec: 'custom' },
      usedCallIds: [],
      allowedExecMethods: ['mcp__wiswork__read_document'],
    })
    const exec = converted.request.tools?.find((tool) => tool.name === 'exec')
    expect(exec?.description).toContain('text(await tools.mcp__wiswork__read_document({...}))')
    expect(exec?.description).not.toContain('ALL_TOOLS')
    expect(exec?.description).not.toContain('apply_patch')
  })

  it.each([
    ['empty input', { model: 'gpt-5.6-sol', input: [] }, 'invalid_conversation'],
    [
      'developer-only',
      {
        model: 'gpt-5.6-sol',
        input: [
          { type: 'message', role: 'developer', content: [{ type: 'input_text', text: 'x' }] },
        ],
      },
      'invalid_conversation',
    ],
    [
      'empty content',
      { model: 'gpt-5.6-sol', input: [{ type: 'message', role: 'user', content: [] }] },
      'invalid_conversation',
    ],
    [
      'late developer',
      {
        model: 'gpt-5.6-sol',
        input: [
          { type: 'message', role: 'user', content: [{ type: 'input_text', text: 'x' }] },
          {
            type: 'message',
            role: 'developer',
            content: [{ type: 'input_text', text: 'late secret' }],
          },
        ],
      },
      'invalid_conversation',
    ],
    [
      'orphan result',
      {
        model: 'gpt-5.6-sol',
        input: [
          { type: 'message', role: 'user', content: [{ type: 'input_text', text: 'x' }] },
          { type: 'function_call_output', call_id: 'orphan', output: 'secret' },
        ],
      },
      'invalid_tool_result_batch',
    ],
    [
      'mismatched result batch',
      {
        model: 'gpt-5.6-sol',
        input: [
          { type: 'message', role: 'user', content: [{ type: 'input_text', text: 'x' }] },
          { type: 'function_call', call_id: 'a', name: 'read', arguments: '{}' },
          { type: 'function_call', call_id: 'b', name: 'read', arguments: '{}' },
          { type: 'function_call_output', call_id: 'b', output: 'secret' },
          { type: 'function_call_output', call_id: 'a', output: 'secret' },
        ],
        tools: [{ type: 'function', name: 'read', parameters: {} }],
      },
      'invalid_tool_result_batch',
    ],
    [
      'duplicate call ID',
      {
        model: 'gpt-5.6-sol',
        input: [
          { type: 'message', role: 'user', content: [{ type: 'input_text', text: 'x' }] },
          { type: 'function_call', call_id: 'a', name: 'read', arguments: '{}' },
          { type: 'function_call', call_id: 'a', name: 'read', arguments: '{}' },
        ],
        tools: [{ type: 'function', name: 'read', parameters: {} }],
      },
      'duplicate_call_id',
    ],
  ])('rejects %s', (_label, request, code) => expectCode(request, code))

  it('preserves a valid alternating conversation and exact tool batches', () => {
    const { request, context } = responsesToMessagesWithContext({
      model: 'gpt-5.6-sol',
      input: [
        { type: 'message', role: 'user', content: [{ type: 'input_text', text: 'one' }] },
        { type: 'message', role: 'user', content: [{ type: 'input_text', text: 'two' }] },
        {
          type: 'message',
          role: 'assistant',
          content: [{ type: 'output_text', text: 'checking' }],
        },
        { type: 'function_call', call_id: 'a', name: 'read', arguments: '{}' },
        { type: 'function_call', call_id: 'b', name: 'read', arguments: '{}' },
        { type: 'function_call_output', call_id: 'a', output: 'A' },
        { type: 'function_call_output', call_id: 'b', output: 'B' },
      ],
      tools: [{ type: 'function', name: 'read', parameters: {} }],
    })
    expect(request.messages.map((message) => message.role)).toEqual(['user', 'assistant', 'user'])
    expect(request.messages[0]?.content).toHaveLength(2)
    expect(context.usedCallIds).toEqual(['a', 'b'])
    expect(context.advertisedTools).toEqual({ read: 'function' })
  })

  it.each([
    ['reserved exec', [{ type: 'function', name: 'exec', parameters: {} }]],
    [
      'duplicate tool',
      [
        { type: 'function', name: 'read', parameters: {} },
        { type: 'function', name: 'read', parameters: {} },
      ],
    ],
  ])('rejects %s registry entries', (_label, tools) => {
    expectCode(
      {
        model: 'gpt-5.6-sol',
        input: [{ type: 'message', role: 'user', content: [{ type: 'input_text', text: 'x' }] }],
        tools,
      },
      'invalid_tool_registry',
    )
  })

  it.each([
    ['item count', { maxRequestItems: 1 }, 'request_item_limit_exceeded'],
    ['content count', { maxContentParts: 1 }, 'request_content_limit_exceeded'],
    ['string length', { maxStringLength: 3 }, 'request_string_limit_exceeded'],
    ['output tokens', { maxOutputTokens: 10 }, 'max_output_tokens_limit_exceeded'],
    ['prompt cache key', { maxPromptCacheKeyLength: 3 }, 'prompt_cache_key_limit_exceeded'],
    ['tool count', { maxTools: 1 }, 'request_tool_limit_exceeded'],
    ['description', { maxDescriptionLength: 3 }, 'unsupported_additional_tool'],
    ['schema', { maxSchemaBytes: 1 }, 'tool_schema_limit_exceeded'],
  ])('enforces configurable %s', (_label, limits, code) => {
    const request = cloneCaptured()
    request.max_output_tokens = 11
    expect(() => responsesToMessagesWithContext(request, limits)).toThrowError(code)
  })

  it('rejects reasoning history instead of inventing a thinking dialect', () => {
    expectCode(
      {
        model: 'gpt-5.6-sol',
        input: [
          { type: 'message', role: 'user', content: [{ type: 'input_text', text: 'x' }] },
          { type: 'reasoning', encrypted_content: 'secret' },
        ],
      },
      'unsupported_reasoning_input',
    )
  })

  it.each([
    [
      'grammar drift',
      (body: any) => (body.input[0].tools[0].tools[0].format.definition += ' drift'),
    ],
    ['namespace drift', (body: any) => (body.input[0].tools[0].name = 'unknown')],
    ['child drift', (body: any) => (body.input[0].tools[0].tools[1].name = 'unknown')],
    ['metadata key drift', (body: any) => (body.client_metadata.unknown = 'secret')],
    [
      'turn metadata drift',
      (body: any) => {
        const metadata = JSON.parse(body.client_metadata['x-codex-turn-metadata'])
        metadata.unknown = 'secret'
        body.client_metadata['x-codex-turn-metadata'] = JSON.stringify(metadata)
      },
    ],
    [
      'duplicate turn metadata key',
      (body: any) => {
        body.client_metadata['x-codex-turn-metadata'] = body.client_metadata[
          'x-codex-turn-metadata'
        ].replace('{', '{"thread_id":"duplicate",')
      },
    ],
  ])('rejects captured-envelope %s without echoing data', (_label, mutate) => {
    const request = cloneCaptured()
    mutate(request)
    try {
      responsesToMessagesWithContext(request)
      throw new Error('expected compatibility failure')
    } catch (error) {
      expect(error).toBeInstanceOf(ProtocolCompatibilityError)
      expect(String(error)).not.toContain('secret')
      expect(String(error)).not.toContain('drift')
    }
  })
})
