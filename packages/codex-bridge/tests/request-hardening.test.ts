import { describe, expect, it } from 'vitest'
import captured from './fixtures/codex-0147-request.json'
import { prepareResponsesTurn, ProtocolCompatibilityError } from '../src/index.js'
import { prepareCarrierTurn } from './fixtures/carrier-authorization.js'

const cloneCaptured = (): Record<string, unknown> =>
  structuredClone(captured) as Record<string, unknown>

function expectCode(request: unknown, code: string): void {
  expect(() => prepareResponsesTurn(request)).toThrowError(
    expect.objectContaining<Partial<ProtocolCompatibilityError>>({ message: code }),
  )
}

describe('strict Responses request conversion', () => {
  it.each([
    Object.defineProperty({}, 'model', { get: () => 'gpt-5.6-sol' }),
    Object.assign({ model: 'gpt-5.6-sol', input: 'hello' }, { [Symbol('secret')]: true }),
    { model: 'gpt-5.6-sol', input: Object.defineProperty([], '0', { get: () => 'secret' }) },
    new Proxy(
      { model: 'gpt-5.6-sol', input: 'hello' },
      {
        ownKeys: () => {
          throw new Error('secret proxy')
        },
      },
    ),
  ])('rejects accessor, symbol, exotic array, and hostile proxy request graphs', (value) => {
    expect(() => prepareResponsesTurn(value)).toThrowError(
      expect.objectContaining<Partial<ProtocolCompatibilityError>>({ message: 'invalid_request' }),
    )
  })

  it('strictly parses limit overrides without invoking accessors', () => {
    const limits = Object.defineProperty({}, 'maxRequestBytes', { get: () => 1 })
    expect(() => prepareResponsesTurn({ model: 'gpt-5.6-sol', input: 'hello' }, limits)).toThrow(
      'invalid_protocol_limits',
    )
  })
  it('accepts the captured Codex 0.147 envelope and returns stream context', () => {
    const converted = prepareCarrierTurn(cloneCaptured())
    expect(converted.messagesRequest).toMatchObject({
      model: 'openai/gpt-5.6-sol',
      system: 'System',
      messages: [{ role: 'user', content: [{ type: 'text', text: 'Hello' }] }],
    })
    expect(converted.messagesRequest).not.toHaveProperty('tool_choice')
    const exec = converted.messagesRequest.tools?.find((tool) => tool.name === 'exec')
    expect(exec?.description).toContain(
      'text(await tools.mcp__wiswork__wiswork_read_document({...}))',
    )
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
          { type: 'custom_tool_call_output', call_id: 'orphan', output: 'secret' },
        ],
      },
      'invalid_tool_result_batch',
    ],
  ])('rejects %s', (_label, request, code) => expectCode(request, code))

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
      'unsupported_top_level_tools',
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
    expect(() => prepareResponsesTurn(request, limits)).toThrowError(code)
  })

  it('round-trips only the exact encrypted reasoning history shape', () => {
    const turn = prepareResponsesTurn({
      model: 'gpt-5.6-sol',
      input: [
        { type: 'message', role: 'user', content: [{ type: 'input_text', text: 'x' }] },
        {
          type: 'reasoning',
          id: 'item_0',
          summary: [],
          content: null,
          encrypted_content: 'opaque-reasoning',
        },
      ],
    })
    expect(turn.messagesRequest.messages).toEqual([
      { role: 'user', content: [{ type: 'text', text: 'x' }] },
      {
        role: 'assistant',
        content: [{ type: 'redacted_thinking', data: 'opaque-reasoning' }],
      },
    ])
  })

  it('accepts the exact empty reasoning history item produced after discarded plaintext thinking', () => {
    const turn = prepareResponsesTurn({
      model: 'gpt-5.6-sol',
      input: [
        { type: 'message', role: 'user', content: [{ type: 'input_text', text: 'x' }] },
        {
          type: 'reasoning',
          id: 'item_0',
          summary: [],
          content: null,
          encrypted_content: null,
        },
      ],
    })
    expect(turn.messagesRequest.messages).toEqual([
      { role: 'user', content: [{ type: 'text', text: 'x' }] },
    ])
  })

  it.each([
    { type: 'reasoning', encrypted_content: 'secret' },
    {
      type: 'reasoning',
      id: 'item_0',
      summary: [{ type: 'summary_text', text: 'secret prompt' }],
      content: null,
      encrypted_content: 'opaque-reasoning',
    },
  ])('rejects unsupported reasoning history variants without exposing content', (reasoning) => {
    expectCode(
      {
        model: 'gpt-5.6-sol',
        input: [
          { type: 'message', role: 'user', content: [{ type: 'input_text', text: 'x' }] },
          reasoning,
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
      prepareCarrierTurn(request)
      throw new Error('expected compatibility failure')
    } catch (error) {
      expect(error).toBeInstanceOf(ProtocolCompatibilityError)
      expect(String(error)).not.toContain('secret')
      expect(String(error)).not.toContain('drift')
    }
  })
})
