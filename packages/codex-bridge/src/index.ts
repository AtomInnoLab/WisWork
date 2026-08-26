import type { MessagesRequest, ResponsesRequest } from './types.js'

export * from './types.js'

export class ProtocolCompatibilityError extends Error {
  constructor(readonly code: string) {
    super(code)
    this.name = 'ProtocolCompatibilityError'
  }
}

type UnknownRecord = Record<string, unknown>

function fail(code: string): never {
  throw new ProtocolCompatibilityError(code)
}

function isRecord(value: unknown): value is UnknownRecord {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function hasOnlyKeys(value: UnknownRecord, allowed: readonly string[]): boolean {
  const keySet = new Set(allowed)
  return Object.keys(value).every((key) => keySet.has(key))
}

function requireString(value: unknown, code: string): string {
  return typeof value === 'string' ? value : fail(code)
}

function isJsonValue(value: unknown, seen = new WeakSet<object>()): boolean {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return true
  if (typeof value === 'number') return Number.isFinite(value)
  if (typeof value !== 'object' || seen.has(value)) return false
  seen.add(value)
  if (Array.isArray(value)) return value.every((item) => isJsonValue(item, seen))
  const prototype = Object.getPrototypeOf(value)
  if (prototype !== Object.prototype && prototype !== null) return false
  return Object.values(value).every((item) => isJsonValue(item, seen))
}

function convertInputItem(item: unknown): MessagesRequest['messages'][number] {
  if (!isRecord(item)) fail('unsupported_input_item')

  if ('role' in item) {
    if (!hasOnlyKeys(item, ['type', 'id', 'role', 'content'])) fail('unsupported_input_item')
    if (item.type !== undefined && item.type !== 'message') fail('unsupported_input_item')
    if (item.id !== undefined && typeof item.id !== 'string') fail('invalid_input_item')
    if (item.role !== 'user' && item.role !== 'assistant') fail('unsupported_input_item')
    if (!Array.isArray(item.content)) fail('unsupported_input_item')
    const content = item.content.map((part): Record<string, unknown> => {
      if (!isRecord(part) || typeof part.type !== 'string') fail('unsupported_input_content')
      if (part.type === 'input_text' || part.type === 'output_text') {
        if (!hasOnlyKeys(part, ['type', 'text'])) fail('unsupported_input_content')
        return { type: 'text', text: requireString(part.text, 'invalid_input_content') }
      }
      if (part.type === 'input_image') {
        if (!hasOnlyKeys(part, ['type', 'image_url'])) fail('unsupported_input_content')
        return {
          type: 'image',
          source: { type: 'url', url: requireString(part.image_url, 'invalid_input_content') },
        }
      }
      return fail('unsupported_input_content')
    })
    return { role: item.role, content }
  }

  if (item.type === 'function_call') {
    if (!hasOnlyKeys(item, ['type', 'call_id', 'name', 'arguments'])) fail('unsupported_input_item')
    const argumentText = requireString(item.arguments, 'invalid_tool_arguments')
    let parsed: unknown
    try {
      parsed = JSON.parse(argumentText)
    } catch {
      fail('invalid_tool_arguments')
    }
    if (!isRecord(parsed)) fail('invalid_tool_arguments')
    return {
      role: 'assistant',
      content: [
        {
          type: 'tool_use',
          id: requireString(item.call_id, 'invalid_tool_call'),
          name: requireString(item.name, 'invalid_tool_call'),
          input: parsed,
        },
      ],
    }
  }

  if (item.type === 'function_call_output') {
    if (!hasOnlyKeys(item, ['type', 'call_id', 'output'])) fail('unsupported_input_item')
    return {
      role: 'user',
      content: [
        {
          type: 'tool_result',
          tool_use_id: requireString(item.call_id, 'invalid_tool_result'),
          content: requireString(item.output, 'invalid_tool_result'),
        },
      ],
    }
  }

  if (item.type === 'custom_tool_call') {
    if (!hasOnlyKeys(item, ['type', 'call_id', 'name', 'input']) || item.name !== 'exec') {
      fail('unsupported_input_item')
    }
    return {
      role: 'assistant',
      content: [
        {
          type: 'tool_use',
          id: requireString(item.call_id, 'invalid_tool_call'),
          name: 'exec',
          input: { code: requireString(item.input, 'invalid_custom_tool_input') },
        },
      ],
    }
  }

  if (item.type === 'custom_tool_call_output') {
    if (!hasOnlyKeys(item, ['type', 'call_id', 'output'])) fail('unsupported_input_item')
    return {
      role: 'user',
      content: [
        {
          type: 'tool_result',
          tool_use_id: requireString(item.call_id, 'invalid_tool_result'),
          content: requireString(item.output, 'invalid_tool_result'),
        },
      ],
    }
  }

  return fail('unsupported_input_item')
}

const OMITTED_FUNCTION_TOOLS = new Set(['wait', 'request_user_input'])
const OMITTED_COLLABORATION_TOOLS = new Set([
  'followup_task',
  'interrupt_agent',
  'list_agents',
  'send_message',
  'spawn_agent',
  'wait_agent',
])

function validateOmittedFunction(tool: unknown, allowed: Set<string>): void {
  if (
    !isRecord(tool) ||
    !hasOnlyKeys(tool, ['type', 'name', 'description', 'strict', 'parameters']) ||
    tool.type !== 'function' ||
    typeof tool.name !== 'string' ||
    !allowed.has(tool.name) ||
    (tool.description !== undefined && typeof tool.description !== 'string') ||
    tool.strict !== false ||
    !isRecord(tool.parameters)
  ) {
    fail('unsupported_additional_tool')
  }
}

function validateAdditionalTools(item: UnknownRecord): string | undefined {
  if (
    !hasOnlyKeys(item, ['type', 'role', 'tools']) ||
    item.role !== 'developer' ||
    !Array.isArray(item.tools)
  ) {
    fail('unsupported_additional_tool')
  }
  let hasExec = false
  let execDescription: string | undefined
  const seenNamespaces = new Set<string>()
  for (const namespace of item.tools) {
    if (
      !isRecord(namespace) ||
      !hasOnlyKeys(namespace, ['type', 'name', 'description', 'tools']) ||
      namespace.type !== 'namespace' ||
      typeof namespace.name !== 'string' ||
      (namespace.description !== undefined && typeof namespace.description !== 'string') ||
      !Array.isArray(namespace.tools)
    ) {
      fail('unsupported_additional_tool')
    }
    if (seenNamespaces.has(namespace.name)) fail('unsupported_additional_tool')
    seenNamespaces.add(namespace.name)
    const seenTools = new Set<string>()
    if (namespace.name === 'functions') {
      for (const tool of namespace.tools) {
        if (!isRecord(tool) || typeof tool.name !== 'string' || seenTools.has(tool.name)) {
          fail('unsupported_additional_tool')
        }
        seenTools.add(tool.name)
        if (isRecord(tool) && tool.type === 'custom' && tool.name === 'exec') {
          if (
            hasExec ||
            !hasOnlyKeys(tool, ['type', 'name', 'description', 'format']) ||
            typeof tool.description !== 'string' ||
            !isRecord(tool.format) ||
            !hasOnlyKeys(tool.format, ['type', 'syntax', 'definition']) ||
            tool.format.type !== 'grammar' ||
            tool.format.syntax !== 'lark' ||
            typeof tool.format.definition !== 'string'
          ) {
            fail('unsupported_additional_tool')
          }
          hasExec = true
          execDescription = tool.description
        } else {
          validateOmittedFunction(tool, OMITTED_FUNCTION_TOOLS)
        }
      }
    } else if (namespace.name === 'collaboration') {
      for (const tool of namespace.tools) {
        if (!isRecord(tool) || typeof tool.name !== 'string' || seenTools.has(tool.name)) {
          fail('unsupported_additional_tool')
        }
        seenTools.add(tool.name)
        validateOmittedFunction(tool, OMITTED_COLLABORATION_TOOLS)
      }
    } else {
      fail('unsupported_additional_tool')
    }
  }
  return execDescription
}

function validateRequest(input: unknown): ResponsesRequest {
  if (!isRecord(input)) fail('invalid_request')
  if (
    !hasOnlyKeys(input, [
      'model',
      'input',
      'instructions',
      'tools',
      'tool_choice',
      'max_output_tokens',
      'parallel_tool_calls',
      'reasoning',
      'store',
      'stream',
      'include',
      'prompt_cache_key',
      'text',
      'client_metadata',
    ])
  ) {
    fail('unsupported_request_field')
  }
  if (input.model !== 'gpt-5.6-sol') fail('unsupported_model')
  if (input.tool_choice !== undefined && input.tool_choice !== 'auto') {
    fail('unsupported_request_field')
  }
  if (input.stream !== undefined && input.stream !== true) fail('unsupported_request_field')
  if (input.parallel_tool_calls !== undefined && typeof input.parallel_tool_calls !== 'boolean') {
    fail('unsupported_request_field')
  }
  if (
    input.reasoning !== undefined &&
    (!isRecord(input.reasoning) ||
      !hasOnlyKeys(input.reasoning, ['effort', 'context']) ||
      input.reasoning.effort !== 'medium' ||
      input.reasoning.context !== 'all_turns')
  ) {
    fail('unsupported_request_field')
  }
  if (input.store !== undefined && input.store !== false) fail('unsupported_request_field')
  if (
    input.include !== undefined &&
    (!Array.isArray(input.include) ||
      input.include.length !== 1 ||
      input.include[0] !== 'reasoning.encrypted_content')
  ) {
    fail('unsupported_request_field')
  }
  if (input.prompt_cache_key !== undefined && typeof input.prompt_cache_key !== 'string') {
    fail('invalid_request')
  }
  if (
    input.text !== undefined &&
    (!isRecord(input.text) ||
      !hasOnlyKeys(input.text, ['verbosity']) ||
      input.text.verbosity !== 'low')
  ) {
    fail('unsupported_request_field')
  }
  if (
    input.client_metadata !== undefined &&
    (!isRecord(input.client_metadata) || !isJsonValue(input.client_metadata))
  ) {
    fail('invalid_client_metadata')
  }
  if (input.instructions !== undefined && typeof input.instructions !== 'string')
    fail('invalid_request')
  if (
    input.max_output_tokens !== undefined &&
    (!Number.isSafeInteger(input.max_output_tokens) || (input.max_output_tokens as number) <= 0)
  ) {
    fail('invalid_max_output_tokens')
  }
  if (typeof input.input !== 'string' && !Array.isArray(input.input)) fail('invalid_request')
  if (input.tools !== undefined && !Array.isArray(input.tools)) fail('invalid_tools')
  return input as unknown as ResponsesRequest
}

export function responsesToMessages(input: unknown): MessagesRequest {
  const request = validateRequest(input)
  const sourceItems: unknown[] =
    typeof request.input === 'string'
      ? [{ role: 'user', content: [{ type: 'input_text', text: request.input }] }]
      : request.input
  const developerInstructions: string[] = []
  let execToolDescription: string | undefined
  const messageItems = sourceItems.filter((item) => {
    if (isRecord(item) && item.type === 'additional_tools') {
      const description = validateAdditionalTools(item)
      if (description !== undefined) {
        if (execToolDescription !== undefined) fail('unsupported_additional_tool')
        execToolDescription = description
      }
      return false
    }
    if (!isRecord(item) || item.role !== 'developer') return true
    if (
      item.type !== 'message' ||
      !hasOnlyKeys(item, ['type', 'id', 'role', 'content']) ||
      (item.id !== undefined && typeof item.id !== 'string') ||
      !Array.isArray(item.content)
    ) {
      fail('unsupported_input_item')
    }
    for (const part of item.content) {
      if (!isRecord(part) || !hasOnlyKeys(part, ['type', 'text']) || part.type !== 'input_text') {
        fail('unsupported_developer_content')
      }
      developerInstructions.push(requireString(part.text, 'invalid_input_content'))
    }
    return false
  })
  const messages = messageItems
    .map(convertInputItem)
    .reduce<MessagesRequest['messages']>((all, next) => {
      const previous = all.at(-1)
      if (previous?.role === next.role) previous.content.push(...next.content)
      else all.push(next)
      return all
    }, [])
  const converted: MessagesRequest = {
    model: 'openai/gpt-5.6-sol',
    messages,
    max_tokens: request.max_output_tokens ?? 4096,
    stream: true,
  }

  const systemParts = [request.instructions, ...developerInstructions].filter(
    (value): value is string => value !== undefined,
  )
  if (systemParts.length > 0) converted.system = systemParts.join('\n\n')
  if (request.tools !== undefined) {
    converted.tools = request.tools.map((tool) => {
      if (!isRecord(tool) || !hasOnlyKeys(tool, ['type', 'name', 'description', 'parameters'])) {
        fail('unsupported_tool')
      }
      if (tool.type !== 'function' || !isRecord(tool.parameters)) fail('unsupported_tool')
      if (tool.description !== undefined && typeof tool.description !== 'string')
        fail('invalid_tool')
      return {
        name: requireString(tool.name, 'invalid_tool'),
        ...(tool.description === undefined ? {} : { description: tool.description }),
        input_schema: tool.parameters,
      }
    })
  }
  if (execToolDescription !== undefined) {
    if (converted.tools?.some((tool) => tool.name === 'exec')) fail('duplicate_tool_name')
    converted.tools ??= []
    converted.tools.push({
      name: 'exec',
      description: execToolDescription,
      input_schema: {
        type: 'object',
        properties: { code: { type: 'string' } },
        required: ['code'],
        additionalProperties: false,
      },
    })
  }
  if (request.parallel_tool_calls === false && converted.tools !== undefined) {
    converted.tool_choice = { type: 'auto', disable_parallel_tool_use: true }
  }

  return converted
}

type StreamBlock =
  | { kind: 'text'; itemId: string; text: string }
  | { kind: 'tool'; itemId: string; callId: string; name: string; arguments: string }
  | { kind: 'custom'; itemId: string; callId: string; arguments: string }

interface StreamState {
  responseId?: string
  inputTokens: number
  cachedTokens: number
  outputTokens: number
  stopReason?: string
  blocks: Map<number, StreamBlock>
  output: Array<Record<string, unknown>>
}

function sse(event: string, data: UnknownRecord): string {
  return `event: ${event}\ndata: ${JSON.stringify({ type: event, ...data })}\n\n`
}

function numeric(value: unknown): number {
  return typeof value === 'number' && Number.isFinite(value) && value >= 0 ? value : 0
}

function responseSnapshot(state: StreamState, status: string): UnknownRecord {
  return {
    id: state.responseId,
    object: 'response',
    model: 'gpt-5.6-sol',
    status,
    output: state.output,
  }
}

function parseFrame(frame: string): { event: string; data: UnknownRecord } {
  let event: string | undefined
  const dataLines: string[] = []
  for (const line of frame.split(/\r?\n/)) {
    if (line.startsWith('event:')) event = line.slice(6).trim()
    else if (line.startsWith('data:')) dataLines.push(line.slice(5).trimStart())
    else if (line !== '' && !line.startsWith(':')) fail('invalid_messages_sse')
  }
  if (!event || dataLines.length === 0) fail('invalid_messages_sse')
  let data: unknown
  try {
    data = JSON.parse(dataLines.join('\n'))
  } catch {
    fail('invalid_messages_sse')
  }
  if (!isRecord(data) || data.type !== event) fail('invalid_messages_sse')
  return { event, data }
}

function requireBlock(state: StreamState, index: unknown): StreamBlock {
  if (!Number.isSafeInteger(index)) fail('invalid_messages_event')
  return state.blocks.get(index as number) ?? fail('invalid_messages_event_order')
}

function convertEvent(event: string, data: UnknownRecord, state: StreamState): string[] {
  if (event === 'error') fail('upstream_error')
  if (
    ![
      'message_start',
      'content_block_start',
      'content_block_delta',
      'content_block_stop',
      'message_delta',
      'message_stop',
    ].includes(event)
  ) {
    fail('unsupported_messages_event')
  }

  if (event === 'message_start') {
    if (state.responseId !== undefined || !isRecord(data.message))
      fail('invalid_messages_event_order')
    const id = requireString(data.message.id, 'invalid_messages_event')
    if (data.message.model !== 'openai/gpt-5.6-sol') fail('unsupported_upstream_model')
    const usage = isRecord(data.message.usage) ? data.message.usage : {}
    const ordinary = numeric(usage.input_tokens)
    const created = numeric(usage.cache_creation_input_tokens)
    const cached = numeric(usage.cache_read_input_tokens)
    state.responseId = id
    state.inputTokens = ordinary + created + cached
    state.cachedTokens = cached
    return [sse('response.created', { response: responseSnapshot(state, 'in_progress') })]
  }

  if (state.responseId === undefined) fail('invalid_messages_event_order')

  if (event === 'content_block_start') {
    if (!Number.isSafeInteger(data.index) || !isRecord(data.content_block))
      fail('invalid_messages_event')
    const index = data.index as number
    if (state.blocks.has(index)) fail('invalid_messages_event_order')
    const itemId = `item_${index}`
    if (data.content_block.type === 'text') {
      if (data.content_block.text !== '') fail('unsupported_messages_event')
      state.blocks.set(index, { kind: 'text', itemId, text: '' })
      const item = {
        id: itemId,
        type: 'message',
        role: 'assistant',
        status: 'in_progress',
        content: [],
      }
      return [
        sse('response.output_item.added', { output_index: index, item }),
        sse('response.content_part.added', {
          item_id: itemId,
          output_index: index,
          content_index: 0,
          part: { type: 'output_text', text: '', annotations: [] },
        }),
      ]
    }
    if (data.content_block.type === 'tool_use') {
      if (
        !isRecord(data.content_block.input) ||
        Object.keys(data.content_block.input).length !== 0
      ) {
        fail('unsupported_messages_event')
      }
      const callId = requireString(data.content_block.id, 'invalid_messages_event')
      const name = requireString(data.content_block.name, 'invalid_messages_event')
      if (name === 'exec') {
        state.blocks.set(index, { kind: 'custom', itemId, callId, arguments: '' })
        return [
          sse('response.output_item.added', {
            output_index: index,
            item: {
              id: itemId,
              type: 'custom_tool_call',
              status: 'in_progress',
              call_id: callId,
              name,
              input: '',
            },
          }),
        ]
      }
      state.blocks.set(index, { kind: 'tool', itemId, callId, name, arguments: '' })
      return [
        sse('response.output_item.added', {
          output_index: index,
          item: {
            id: itemId,
            type: 'function_call',
            status: 'in_progress',
            call_id: callId,
            name,
            arguments: '',
          },
        }),
      ]
    }
    return fail('unsupported_content_block')
  }

  if (event === 'content_block_delta') {
    if (!isRecord(data.delta)) fail('invalid_messages_event')
    const block = requireBlock(state, data.index)
    if (data.delta.type === 'text_delta' && block.kind === 'text') {
      const delta = requireString(data.delta.text, 'invalid_messages_event')
      block.text += delta
      return [
        sse('response.output_text.delta', {
          item_id: block.itemId,
          output_index: data.index,
          content_index: 0,
          delta,
        }),
      ]
    }
    if (
      data.delta.type === 'input_json_delta' &&
      (block.kind === 'tool' || block.kind === 'custom')
    ) {
      const delta = requireString(data.delta.partial_json, 'invalid_messages_event')
      block.arguments += delta
      if (block.kind === 'custom') return []
      return [
        sse('response.function_call_arguments.delta', {
          item_id: block.itemId,
          output_index: data.index,
          delta,
        }),
      ]
    }
    return fail('unsupported_content_delta')
  }

  if (event === 'content_block_stop') {
    const index = data.index
    const block = requireBlock(state, index)
    state.blocks.delete(index as number)
    if (block.kind === 'text') {
      const part = { type: 'output_text', text: block.text, annotations: [] }
      const item = {
        id: block.itemId,
        type: 'message',
        role: 'assistant',
        status: 'completed',
        content: [part],
      }
      state.output.push(item)
      return [
        sse('response.output_text.done', {
          item_id: block.itemId,
          output_index: index,
          content_index: 0,
          text: block.text,
        }),
        sse('response.content_part.done', {
          item_id: block.itemId,
          output_index: index,
          content_index: 0,
          part,
        }),
        sse('response.output_item.done', { output_index: index, item }),
      ]
    }
    if (block.kind === 'custom') {
      let parsed: unknown
      try {
        parsed = JSON.parse(block.arguments)
      } catch {
        fail('invalid_custom_tool_input')
      }
      if (!isRecord(parsed) || !hasOnlyKeys(parsed, ['code']) || typeof parsed.code !== 'string') {
        fail('invalid_custom_tool_input')
      }
      const item = {
        id: block.itemId,
        type: 'custom_tool_call',
        status: 'completed',
        call_id: block.callId,
        name: 'exec',
        input: parsed.code,
      }
      state.output.push(item)
      return [
        sse('response.custom_tool_call_input.delta', {
          item_id: block.itemId,
          output_index: index,
          delta: parsed.code,
        }),
        sse('response.custom_tool_call_input.done', {
          item_id: block.itemId,
          output_index: index,
          input: parsed.code,
        }),
        sse('response.output_item.done', { output_index: index, item }),
      ]
    }
    try {
      const parsed = JSON.parse(block.arguments)
      if (!isRecord(parsed)) fail('invalid_tool_arguments')
    } catch (error) {
      if (error instanceof ProtocolCompatibilityError) throw error
      fail('invalid_tool_arguments')
    }
    const item = {
      id: block.itemId,
      type: 'function_call',
      status: 'completed',
      call_id: block.callId,
      name: block.name,
      arguments: block.arguments,
    }
    state.output.push(item)
    return [
      sse('response.function_call_arguments.done', {
        item_id: block.itemId,
        output_index: index,
        arguments: block.arguments,
      }),
      sse('response.output_item.done', { output_index: index, item }),
    ]
  }

  if (event === 'message_delta') {
    if (!isRecord(data.delta) || !isRecord(data.usage)) fail('invalid_messages_event')
    state.stopReason = requireString(data.delta.stop_reason, 'invalid_messages_event')
    state.outputTokens = numeric(data.usage.output_tokens)
    return []
  }

  if (event === 'message_stop') {
    if (state.blocks.size !== 0 || state.stopReason === undefined)
      fail('invalid_messages_event_order')
    const completed = ['end_turn', 'stop_sequence', 'tool_use'].includes(state.stopReason)
    if (!completed && state.stopReason !== 'max_tokens') fail('unsupported_stop_reason')
    const response = {
      ...responseSnapshot(state, completed ? 'completed' : 'incomplete'),
      ...(completed ? {} : { incomplete_details: { reason: 'max_output_tokens' } }),
      usage: {
        input_tokens: state.inputTokens,
        output_tokens: state.outputTokens,
        total_tokens: state.inputTokens + state.outputTokens,
        input_tokens_details: { cached_tokens: state.cachedTokens },
        output_tokens_details: { reasoning_tokens: 0 },
      },
    }
    return [
      sse(completed ? 'response.completed' : 'response.incomplete', { response }),
      'data: [DONE]\n\n',
    ]
  }

  return fail('unsupported_messages_event')
}

export async function* messagesSseToResponses(
  chunks: AsyncIterable<string | Uint8Array>,
): AsyncGenerator<string> {
  const decoder = new TextDecoder()
  const state: StreamState = {
    inputTokens: 0,
    cachedTokens: 0,
    outputTokens: 0,
    blocks: new Map(),
    output: [],
  }
  let buffer = ''

  for await (const chunk of chunks) {
    buffer += typeof chunk === 'string' ? chunk : decoder.decode(chunk, { stream: true })
    while (true) {
      const boundary = /\r?\n\r?\n/.exec(buffer)
      if (!boundary || boundary.index === undefined) break
      const frame = buffer.slice(0, boundary.index)
      buffer = buffer.slice(boundary.index + boundary[0].length)
      if (frame.trim() === '') continue
      const parsed = parseFrame(frame)
      for (const converted of convertEvent(parsed.event, parsed.data, state)) yield converted
    }
  }
  buffer += decoder.decode()
  if (buffer.trim() !== '') fail('invalid_messages_sse')
}
