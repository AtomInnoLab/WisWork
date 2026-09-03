import { describe, expect, it } from 'vitest'
import {
  CODEX_0147_EXEC_GRAMMAR,
  createDocumentCarrierIssuer,
  prepareResponsesTurn,
  ProtocolCompatibilityError,
} from '../src/index.js'

const responsesToMessages = (input: unknown) => {
  if (typeof input !== 'object' || input === null || !('client_metadata' in input)) {
    return prepareResponsesTurn(input).messagesRequest
  }
  const issuer = createDocumentCarrierIssuer(
    { host: 'slides', documentId: 'document-1', sessionId: 'session-1', generation: 1 },
    (capability) => capability === 'fixture-proof',
  )
  const handle = issuer.issueForTurn({
    turnId: 'turn_1',
    sourceNonce: 'N'.repeat(43),
    capability: 'fixture-proof',
    method: 'mcp__wiswork__read_document',
    toolName: 'read_document',
    schemaDigest: 'a'.repeat(64),
  })
  return issuer.prepareTurn(input, {}, handle).messagesRequest
}

const metadata = {
  'x-codex-turn-metadata': JSON.stringify({
    session_id: 'session-1',
    thread_id: 'thread_1',
    turn_id: 'turn_1',
    code_mode_tool_names: {
      mcp__wiswork__read_document: {
        name: 'read_document',
        namespace: 'mcp__wiswork',
      },
    },
  }),
}

const additionalTools = {
  type: 'additional_tools',
  role: 'developer',
  tools: [
    {
      type: 'namespace',
      name: 'functions',
      tools: [
        {
          type: 'custom',
          name: 'exec',
          description: 'Run JavaScript code with tools.mcp__wiswork__read_document and ALL_TOOLS.',
          format: { type: 'grammar', syntax: 'lark', definition: CODEX_0147_EXEC_GRAMMAR },
        },
        { type: 'function', name: 'wait', description: 'Wait', strict: false, parameters: {} },
        {
          type: 'function',
          name: 'request_user_input',
          description: 'Ask',
          strict: false,
          parameters: {},
        },
      ],
    },
    {
      type: 'namespace',
      name: 'collaboration',
      tools: [
        'followup_task',
        'interrupt_agent',
        'list_agents',
        'send_message',
        'spawn_agent',
        'wait_agent',
      ].map((name) => ({
        type: 'function',
        name,
        description: name,
        strict: false,
        parameters: {},
      })),
    },
  ],
}

describe('responsesToMessages', () => {
  it('maps instructions, text, model, and max tokens', () => {
    expect(
      responsesToMessages({
        model: 'gpt-5.6-sol',
        instructions: 'Be concise.',
        input: 'Hello',
        max_output_tokens: 123,
        stream: true,
      }),
    ).toEqual({
      model: 'openai/gpt-5.6-sol',
      system: 'Be concise.',
      messages: [{ role: 'user', content: [{ type: 'text', text: 'Hello' }] }],
      max_tokens: 123,
      stream: true,
    })
  })

  it('accepts and safely omits the fixed Codex 0.147 request metadata', () => {
    expect(
      responsesToMessages({
        model: 'gpt-5.6-sol',
        input: [
          additionalTools,
          {
            type: 'message',
            id: 'msg_developer_1',
            role: 'developer',
            content: [{ type: 'input_text', text: 'System from Codex' }],
          },
          {
            type: 'message',
            id: 'msg_user_1',
            role: 'user',
            content: [{ type: 'input_text', text: 'Hello' }],
          },
        ],
        tool_choice: 'auto',
        parallel_tool_calls: false,
        reasoning: { effort: 'medium', context: 'all_turns' },
        store: false,
        stream: true,
        include: ['reasoning.encrypted_content'],
        prompt_cache_key: 'cache-secret',
        text: { verbosity: 'low' },
        client_metadata: metadata,
      }),
    ).toMatchObject({
      model: 'openai/gpt-5.6-sol',
      system: 'System from Codex',
      messages: [{ role: 'user', content: [{ type: 'text', text: 'Hello' }] }],
      max_tokens: 32_768,
      stream: true,
      tools: [
        {
          name: 'exec',
          input_schema: {
            type: 'object',
            properties: { code: { type: 'string' } },
            required: ['code'],
            additionalProperties: false,
          },
        },
      ],
    })
    expect(
      responsesToMessages({
        model: 'gpt-5.6-sol',
        input: [
          additionalTools,
          { type: 'message', role: 'user', content: [{ type: 'input_text', text: 'Hello' }] },
        ],
        tool_choice: 'auto',
        parallel_tool_calls: false,
        client_metadata: metadata,
      }),
    ).not.toHaveProperty('tool_choice')
  })

  it('maps custom exec call history and preserves call IDs', () => {
    expect(
      responsesToMessages({
        model: 'gpt-5.6-sol',
        input: [
          additionalTools,
          { type: 'message', role: 'user', content: [{ type: 'input_text', text: 'Run it' }] },
          {
            type: 'custom_tool_call',
            id: 'item_call_1',
            call_id: 'custom_1',
            name: 'exec',
            input: 'text(await tools.mcp__wiswork__read_document({}))',
            status: 'completed',
          },
          {
            type: 'custom_tool_call_output',
            id: 'item_output_1',
            call_id: 'custom_1',
            output: [{ type: 'input_text', text: '42' }],
            status: 'completed',
          },
        ],
        client_metadata: metadata,
      }).messages,
    ).toEqual([
      { role: 'user', content: [{ type: 'text', text: 'Run it' }] },
      {
        role: 'assistant',
        content: [
          {
            type: 'tool_use',
            id: 'custom_1',
            name: 'exec',
            input: { code: 'text(await tools.mcp__wiswork__read_document({}))' },
          },
        ],
      },
      {
        role: 'user',
        content: [{ type: 'tool_result', tool_use_id: 'custom_1', content: '42' }],
      },
    ])
  })

  it.each([
    ['invalid model', { model: 'gpt-4o', input: 'secret prompt' }, 'unsupported_model'],
    [
      'unsupported top-level field',
      { model: 'gpt-5.6-sol', input: 'secret prompt', previous_response_id: 'resp_1' },
      'unsupported_request_field',
    ],
    [
      'unsupported input item',
      { model: 'gpt-5.6-sol', input: [{ type: 'computer_call', secret: 'secret prompt' }] },
      'unsupported_input_item',
    ],
    [
      'unsupported top-level tools',
      {
        model: 'gpt-5.6-sol',
        input: [
          { type: 'message', role: 'user', content: [{ type: 'input_text', text: 'x' }] },
          { type: 'function_call', call_id: 'call_1', name: 'x', arguments: 'secret prompt {' },
        ],
        tools: [{ type: 'function', name: 'x', parameters: {} }],
      },
      'unsupported_top_level_tools',
    ],
    [
      'unknown additional_tools child',
      {
        model: 'gpt-5.6-sol',
        input: [
          {
            type: 'additional_tools',
            role: 'developer',
            tools: [
              {
                type: 'namespace',
                name: 'functions',
                tools: [{ type: 'function', name: 'secret prompt', parameters: {} }],
              },
            ],
          },
        ],
      },
      'unsupported_additional_tool',
    ],
    [
      'unknown Codex metadata values',
      {
        model: 'gpt-5.6-sol',
        input: 'secret prompt',
        tool_choice: 'required',
        reasoning: { effort: 'high', context: 'secret prompt' },
        store: true,
        include: ['secret prompt'],
        text: { verbosity: 'high' },
      },
      'unsupported_request_field',
    ],
    [
      'non-object client metadata',
      { model: 'gpt-5.6-sol', input: 'secret prompt', client_metadata: 'secret prompt' },
      'invalid_client_metadata',
    ],
  ])('fails closed for %s without echoing request data', (_label, request, code) => {
    expect(() => responsesToMessages(request)).toThrowError(
      expect.objectContaining<Partial<ProtocolCompatibilityError>>({ message: code }),
    )
    try {
      responsesToMessages(request)
    } catch (error) {
      expect(String(error)).not.toContain('secret prompt')
    }
  })
})
