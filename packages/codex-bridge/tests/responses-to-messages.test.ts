import { describe, expect, it } from 'vitest'
import { ProtocolCompatibilityError, responsesToMessages } from '../src/index.js'

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
          format: { type: 'grammar', syntax: 'lark', definition: 'start: /[\\s\\S]+/' },
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
  it('maps instructions, text, model, max tokens, and tools', () => {
    expect(
      responsesToMessages({
        model: 'gpt-5.6-sol',
        instructions: 'Be concise.',
        input: 'Hello',
        max_output_tokens: 123,
        parallel_tool_calls: true,
        stream: true,
        tools: [
          {
            type: 'function',
            name: 'read_document',
            description: 'Read the current document',
            parameters: { type: 'object', properties: {} },
          },
        ],
      }),
    ).toEqual({
      model: 'openai/gpt-5.6-sol',
      system: 'Be concise.',
      messages: [{ role: 'user', content: [{ type: 'text', text: 'Hello' }] }],
      max_tokens: 123,
      stream: true,
      tools: [
        {
          name: 'read_document',
          description: 'Read the current document',
          input_schema: { type: 'object', properties: {} },
        },
      ],
    })
  })

  it('maps text, images, calls, and tool results in their original order', () => {
    const converted = responsesToMessages({
      model: 'gpt-5.6-sol',
      input: [
        {
          role: 'user',
          content: [
            { type: 'input_text', text: 'Inspect this' },
            { type: 'input_image', image_url: 'https://example.test/image.png' },
          ],
        },
        { type: 'function_call', call_id: 'call_17', name: 'inspect', arguments: '{"page":2}' },
        { type: 'function_call', call_id: 'call_18', name: 'inspect', arguments: '{"page":3}' },
        { type: 'function_call_output', call_id: 'call_17', output: 'done' },
        { type: 'function_call_output', call_id: 'call_18', output: 'also done' },
        { role: 'assistant', content: [{ type: 'output_text', text: 'Finished' }] },
      ],
    })

    expect(converted.messages).toEqual([
      {
        role: 'user',
        content: [
          { type: 'text', text: 'Inspect this' },
          { type: 'image', source: { type: 'url', url: 'https://example.test/image.png' } },
        ],
      },
      {
        role: 'assistant',
        content: [
          { type: 'tool_use', id: 'call_17', name: 'inspect', input: { page: 2 } },
          { type: 'tool_use', id: 'call_18', name: 'inspect', input: { page: 3 } },
        ],
      },
      {
        role: 'user',
        content: [
          { type: 'tool_result', tool_use_id: 'call_17', content: 'done' },
          { type: 'tool_result', tool_use_id: 'call_18', content: 'also done' },
        ],
      },
      { role: 'assistant', content: [{ type: 'text', text: 'Finished' }] },
    ])
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
        client_metadata: { 'x-codex-turn-metadata': '{"thread_id":"secret"}' },
      }),
    ).toEqual({
      model: 'openai/gpt-5.6-sol',
      system: 'System from Codex',
      messages: [{ role: 'user', content: [{ type: 'text', text: 'Hello' }] }],
      max_tokens: 4096,
      stream: true,
      tools: [
        {
          name: 'exec',
          description: 'Run JavaScript code with tools.mcp__wiswork__read_document and ALL_TOOLS.',
          input_schema: {
            type: 'object',
            properties: { code: { type: 'string' } },
            required: ['code'],
            additionalProperties: false,
          },
        },
      ],
      tool_choice: { type: 'auto', disable_parallel_tool_use: true },
    })
  })

  it('maps custom exec call history and preserves call IDs', () => {
    expect(
      responsesToMessages({
        model: 'gpt-5.6-sol',
        input: [
          { type: 'custom_tool_call', call_id: 'custom_1', name: 'exec', input: 'text(42)' },
          { type: 'custom_tool_call_output', call_id: 'custom_1', output: '42' },
        ],
      }).messages,
    ).toEqual([
      {
        role: 'assistant',
        content: [{ type: 'tool_use', id: 'custom_1', name: 'exec', input: { code: 'text(42)' } }],
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
      'invalid tool arguments',
      {
        model: 'gpt-5.6-sol',
        input: [
          { type: 'function_call', call_id: 'call_1', name: 'x', arguments: 'secret prompt {' },
        ],
      },
      'invalid_tool_arguments',
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
