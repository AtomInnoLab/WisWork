import type {
  MessagesRequest,
  PreparedResponsesTurn,
  ProtocolLimits,
  ResponsesRequest,
} from './types.js'

export * from './types.js'
export * from './local-server.js'

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

export const CODEX_0147_EXEC_GRAMMAR =
  '\nstart: pragma_source | plain_source\npragma_source: PRAGMA_LINE NEWLINE SOURCE\nplain_source: SOURCE\n\nPRAGMA_LINE: /[ \\t]*\\/\\/ @exec:[^\\r\\n]*/\nNEWLINE: /\\r?\\n/\nSOURCE: /[\\s\\S]+/\n'

export const DEFAULT_PROTOCOL_LIMITS: Readonly<ProtocolLimits> = Object.freeze({
  maxRequestItems: 512,
  maxRequestBytes: 8_000_000,
  maxRequestNodes: 100_000,
  maxNestingDepth: 64,
  maxContentParts: 1024,
  maxTools: 256,
  maxStringLength: 1_000_000,
  maxDescriptionLength: 256_000,
  maxSchemaBytes: 1_000_000,
  maxOutputTokens: 128_000,
  maxPromptCacheKeyLength: 512,
  maxClientMetadataBytes: 128_000,
  maxWorkspaces: 32,
  maxSseFrameBytes: 1_000_000,
  maxSseBufferBytes: 2_000_000,
  maxSseFrames: 100_000,
  maxBlocks: 512,
  maxAccumulatedText: 16_000_000,
  maxToolArguments: 4_000_000,
  maxTotalOutput: 32_000_000,
})

function resolveLimits(overrides: Partial<ProtocolLimits> = {}): ProtocolLimits {
  if (!Object.keys(overrides).every((key) => key in DEFAULT_PROTOCOL_LIMITS))
    fail('invalid_protocol_limits')
  const limits = { ...DEFAULT_PROTOCOL_LIMITS, ...overrides }
  for (const value of Object.values(limits)) {
    if (!Number.isSafeInteger(value) || value <= 0) fail('invalid_protocol_limits')
  }
  return limits
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

const utf8Length = (value: string): number => new TextEncoder().encode(value).byteLength

function inspectJsonGraph(value: unknown, limits: ProtocolLimits): void {
  const pending: Array<{ value: unknown; depth: number }> = [{ value, depth: 0 }]
  const seen = new WeakSet<object>()
  let nodes = 0
  let bytes = 0
  while (pending.length > 0) {
    const current = pending.pop()!
    nodes += 1
    if (nodes > limits.maxRequestNodes) fail('request_nodes_limit_exceeded')
    if (current.depth > limits.maxNestingDepth) fail('request_nesting_limit_exceeded')
    const item = current.value
    if (item === null) bytes += 4
    else if (typeof item === 'string') bytes += utf8Length(JSON.stringify(item))
    else if (typeof item === 'boolean') bytes += item ? 4 : 5
    else if (typeof item === 'number') {
      if (!Number.isFinite(item)) fail('invalid_request')
      bytes += String(item).length
    } else if (typeof item === 'object') {
      if (seen.has(item)) fail('invalid_request')
      seen.add(item)
      if (!Array.isArray(item)) {
        const prototype = Object.getPrototypeOf(item)
        if (prototype !== Object.prototype && prototype !== null) fail('invalid_request')
      }
      const entries = Array.isArray(item)
        ? item.map((child, index) => [String(index), child] as const)
        : Object.entries(item)
      bytes += 2 + Math.max(0, entries.length - 1)
      for (const [key, child] of entries) {
        if (!Array.isArray(item)) {
          if (utf8Length(key) > limits.maxStringLength) fail('request_string_limit_exceeded')
          bytes += utf8Length(JSON.stringify(key)) + 1
        }
        pending.push({ value: child, depth: current.depth + 1 })
      }
    } else fail('invalid_request')
    if (bytes > limits.maxRequestBytes) fail('request_bytes_limit_exceeded')
  }
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
  if (input.client_metadata !== undefined && !isRecord(input.client_metadata)) {
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
  if (input.tools !== undefined && input.tools.length > 0) fail('unsupported_top_level_tools')
  return input as unknown as ResponsesRequest
}

const PINNED_FUNCTION_CHILDREN = new Set(['exec', 'wait', 'request_user_input'])
const PINNED_COLLABORATION_CHILDREN = new Set([
  'followup_task',
  'interrupt_agent',
  'list_agents',
  'send_message',
  'spawn_agent',
  'wait_agent',
])
const PINNED_BUILTIN_CODE_METHODS = new Set([
  'apply_patch',
  'exec_command',
  'list_mcp_resource_templates',
  'list_mcp_resources',
  'read_mcp_resource',
  'update_plan',
  'view_image',
  'write_stdin',
])
const CLIENT_METADATA_KEYS = new Set([
  'session_id',
  'x-codex-installation-id',
  'turn_id',
  'x-codex-window-id',
  'thread_id',
  'x-codex-turn-metadata',
])
const TURN_METADATA_KEYS = new Set([
  'installation_id',
  'session_id',
  'thread_id',
  'turn_id',
  'window_id',
  'request_kind',
  'sandbox',
  'workspaces',
  'code_mode_tool_names',
  'turn_started_at_unix_ms',
])

function enforceStringLimits(value: unknown, limit: number): void {
  const pending = [value]
  while (pending.length > 0) {
    const item = pending.pop()
    if (typeof item === 'string') {
      if (utf8Length(item) > limit) fail('request_string_limit_exceeded')
    } else if (typeof item === 'object' && item !== null) {
      pending.push(...(Array.isArray(item) ? item : Object.values(item)))
    }
  }
}

function boundedJsonSize(value: unknown, limit: number, code: string): void {
  let serialized: string
  try {
    serialized = JSON.stringify(value)
  } catch {
    fail(code)
  }
  if (utf8Length(serialized) > limit) fail(code)
}

function extractAllowedExecMethods(
  metadata: Record<string, unknown> | undefined,
  limits: ProtocolLimits,
): string[] {
  if (metadata === undefined) return []
  if (!isRecord(metadata) || !Object.keys(metadata).every((key) => CLIENT_METADATA_KEYS.has(key))) {
    fail('unsupported_client_metadata')
  }
  boundedJsonSize(metadata, limits.maxClientMetadataBytes, 'client_metadata_limit_exceeded')
  for (const value of Object.values(metadata)) {
    if (typeof value !== 'string' || utf8Length(value) > limits.maxStringLength) {
      fail('unsupported_client_metadata')
    }
  }
  const packedValue = metadata['x-codex-turn-metadata']
  if (packedValue === undefined) return []
  if (typeof packedValue !== 'string') fail('unsupported_client_metadata')
  const packed = packedValue
  let parsed: unknown
  try {
    parsed = JSON.parse(packed)
  } catch {
    fail('unsupported_client_metadata')
  }
  inspectJsonGraph(parsed, limits)
  if (!isRecord(parsed) || !Object.keys(parsed).every((key) => TURN_METADATA_KEYS.has(key))) {
    fail('unsupported_client_metadata')
  }
  const jsonKeyCount = (key: string): number => {
    const quoted = JSON.stringify(key).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    return packed.match(new RegExp(`${quoted}\\s*:`, 'g'))?.length ?? 0
  }
  for (const key of Object.keys(parsed)) {
    if (jsonKeyCount(key) !== 1) fail('unsupported_client_metadata')
  }
  for (const key of [
    'installation_id',
    'session_id',
    'thread_id',
    'turn_id',
    'window_id',
    'request_kind',
    'sandbox',
  ]) {
    if (parsed[key] !== undefined && typeof parsed[key] !== 'string')
      fail('unsupported_client_metadata')
  }
  if (
    parsed.turn_started_at_unix_ms !== undefined &&
    (!Number.isSafeInteger(parsed.turn_started_at_unix_ms) ||
      (parsed.turn_started_at_unix_ms as number) < 0)
  ) {
    fail('unsupported_client_metadata')
  }
  if (parsed.workspaces !== undefined) {
    if (!isRecord(parsed.workspaces)) fail('unsupported_client_metadata')
    const workspaces = Object.entries(parsed.workspaces)
    if (workspaces.length > limits.maxWorkspaces) fail('client_metadata_limit_exceeded')
    for (const [path, workspace] of workspaces) {
      if (
        path === '' ||
        utf8Length(path) > limits.maxStringLength ||
        !isRecord(workspace) ||
        !hasOnlyKeys(workspace, ['associated_remote_urls', 'latest_git_commit_hash', 'has_changes'])
      ) {
        fail('unsupported_client_metadata')
      }
      if (workspace.associated_remote_urls !== undefined) {
        if (!isRecord(workspace.associated_remote_urls)) fail('unsupported_client_metadata')
        const remotes = Object.entries(workspace.associated_remote_urls)
        if (remotes.length > limits.maxWorkspaces) fail('client_metadata_limit_exceeded')
        for (const [name, url] of remotes) {
          if (
            name === '' ||
            utf8Length(name) > limits.maxStringLength ||
            typeof url !== 'string' ||
            utf8Length(url) > limits.maxStringLength
          ) {
            fail('unsupported_client_metadata')
          }
        }
      }
      if (
        workspace.latest_git_commit_hash !== undefined &&
        (typeof workspace.latest_git_commit_hash !== 'string' ||
          utf8Length(workspace.latest_git_commit_hash) > limits.maxStringLength)
      ) {
        fail('unsupported_client_metadata')
      }
      if (workspace.has_changes !== undefined && typeof workspace.has_changes !== 'boolean') {
        fail('unsupported_client_metadata')
      }
    }
  }
  if (parsed.code_mode_tool_names === undefined) return []
  if (!isRecord(parsed.code_mode_tool_names)) fail('unsupported_client_metadata')
  const allowed: string[] = []
  for (const [method, descriptor] of Object.entries(parsed.code_mode_tool_names)) {
    if (jsonKeyCount(method) !== 1) fail('unsupported_client_metadata')
    if (
      !isRecord(descriptor) ||
      !hasOnlyKeys(descriptor, ['name', 'namespace']) ||
      typeof descriptor.name !== 'string' ||
      (descriptor.namespace !== null && typeof descriptor.namespace !== 'string')
    ) {
      fail('unsupported_client_metadata')
    }
    if (descriptor.namespace === null) {
      if (!PINNED_BUILTIN_CODE_METHODS.has(method) || descriptor.name !== method) {
        fail('unsupported_client_metadata')
      }
      continue
    }
    if (
      descriptor.namespace !== 'mcp__wiswork' ||
      method !== `mcp__wiswork__${descriptor.name}` ||
      allowed.includes(method)
    ) {
      fail('unsupported_client_metadata')
    }
    allowed.push(method)
  }
  return allowed.sort()
}

function validatePinnedFunction(tool: unknown, allowed: Set<string>, limits: ProtocolLimits): void {
  if (
    !isRecord(tool) ||
    !hasOnlyKeys(tool, ['type', 'name', 'description', 'strict', 'parameters']) ||
    tool.type !== 'function' ||
    typeof tool.name !== 'string' ||
    !allowed.has(tool.name) ||
    typeof tool.description !== 'string' ||
    utf8Length(tool.description) > limits.maxDescriptionLength ||
    tool.strict !== false ||
    !isRecord(tool.parameters)
  ) {
    fail('unsupported_additional_tool')
  }
  boundedJsonSize(tool.parameters, limits.maxSchemaBytes, 'tool_schema_limit_exceeded')
}

function validatePinnedAdditionalTools(item: UnknownRecord, limits: ProtocolLimits): boolean {
  if (
    !hasOnlyKeys(item, ['type', 'role', 'tools']) ||
    item.role !== 'developer' ||
    !Array.isArray(item.tools) ||
    item.tools.length < 1 ||
    item.tools.length > 2
  ) {
    fail('unsupported_additional_tool')
  }
  const seenNamespaces = new Set<string>()
  let hasExec = false
  let nestedToolCount = 0
  for (const namespace of item.tools) {
    if (
      !isRecord(namespace) ||
      !hasOnlyKeys(namespace, ['type', 'name', 'description', 'tools']) ||
      namespace.type !== 'namespace' ||
      typeof namespace.name !== 'string' ||
      (namespace.description !== undefined && typeof namespace.description !== 'string') ||
      !Array.isArray(namespace.tools) ||
      seenNamespaces.has(namespace.name)
    ) {
      fail('unsupported_additional_tool')
    }
    seenNamespaces.add(namespace.name)
    const expected =
      namespace.name === 'functions'
        ? PINNED_FUNCTION_CHILDREN
        : namespace.name === 'collaboration'
          ? PINNED_COLLABORATION_CHILDREN
          : fail('unsupported_additional_tool')
    if (namespace.tools.length !== expected.size) fail('unsupported_additional_tool')
    nestedToolCount += namespace.tools.length
    if (nestedToolCount > limits.maxTools) fail('request_tool_limit_exceeded')
    const seenChildren = new Set<string>()
    for (const tool of namespace.tools) {
      if (!isRecord(tool) || typeof tool.name !== 'string' || seenChildren.has(tool.name)) {
        fail('unsupported_additional_tool')
      }
      seenChildren.add(tool.name)
      if (namespace.name === 'functions' && tool.name === 'exec') {
        if (
          !hasOnlyKeys(tool, ['type', 'name', 'description', 'format']) ||
          tool.type !== 'custom' ||
          typeof tool.description !== 'string' ||
          utf8Length(tool.description) > limits.maxDescriptionLength ||
          !isRecord(tool.format) ||
          !hasOnlyKeys(tool.format, ['type', 'syntax', 'definition']) ||
          tool.format.type !== 'grammar' ||
          tool.format.syntax !== 'lark' ||
          tool.format.definition !== CODEX_0147_EXEC_GRAMMAR
        ) {
          fail('unsupported_additional_tool')
        }
        hasExec = true
      } else {
        validatePinnedFunction(tool, expected, limits)
      }
    }
  }
  if (!seenNamespaces.has('functions') || !hasExec) fail('unsupported_additional_tool')
  return true
}

function safeExecDescription(methods: readonly string[]): string {
  return `Execute exactly one document MCP call. An optional first line // @exec: {"yield_time_ms":1000,"max_output_tokens":100} is allowed. Allowed syntax: text(await tools.${methods.join(
    '({...})) or text(await tools.',
  )}({...})). Arguments must be a JSON object literal. No other JavaScript is allowed.`
}

function parseSafeExecCode(code: string, methods: readonly string[], limits: ProtocolLimits): void {
  if (utf8Length(code) > limits.maxToolArguments) fail('tool_arguments_limit_exceeded')
  let source = code
  const pragmaPrefix = /^[\t ]*\/\/ @exec:/.exec(source)
  if (pragmaPrefix) {
    const newline = source.indexOf('\n')
    if (newline < 0) fail('unsafe_custom_tool_input')
    const pragmaText = source.slice(pragmaPrefix[0].length, newline).trim()
    if (utf8Length(pragmaText) > 1024) fail('unsafe_custom_tool_input')
    let pragma: unknown
    try {
      pragma = JSON.parse(pragmaText)
    } catch {
      fail('unsafe_custom_tool_input')
    }
    if (!isRecord(pragma) || !hasOnlyKeys(pragma, ['yield_time_ms', 'max_output_tokens'])) {
      fail('unsafe_custom_tool_input')
    }
    for (const key of Object.keys(pragma)) {
      const quoted = JSON.stringify(key).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
      if ((pragmaText.match(new RegExp(`${quoted}\\s*:`, 'g'))?.length ?? 0) !== 1)
        fail('unsafe_custom_tool_input')
    }
    for (const [key, value] of Object.entries(pragma)) {
      const cap = key === 'yield_time_ms' ? 300_000 : limits.maxOutputTokens
      if (!Number.isSafeInteger(value) || (value as number) <= 0 || (value as number) > cap) {
        fail('unsafe_custom_tool_input')
      }
    }
    source = source.slice(newline + 1)
  }
  const direct = /^await\s+tools\.([A-Za-z_][A-Za-z0-9_]*)\((\{[\s\S]*\})\);?$/
  const wrapped = /^text\(\s*await\s+tools\.([A-Za-z_][A-Za-z0-9_]*)\((\{[\s\S]*\})\)\s*\);?$/
  const match = wrapped.exec(source.trim()) ?? direct.exec(source.trim())
  if (!match || !methods.includes(match[1]!)) fail('unsafe_custom_tool_input')
  let argument: unknown
  try {
    argument = JSON.parse(match[2]!)
  } catch {
    fail('unsafe_custom_tool_input')
  }
  if (!isRecord(argument)) fail('unsafe_custom_tool_input')
  inspectJsonGraph(argument, limits)
}

function convertMessageContent(
  item: UnknownRecord,
  limits: ProtocolLimits,
  count: { parts: number },
): { role: 'user' | 'assistant'; content: Array<Record<string, unknown>> } {
  if (
    !hasOnlyKeys(item, ['type', 'id', 'role', 'content']) ||
    item.type !== 'message' ||
    (item.id !== undefined && typeof item.id !== 'string') ||
    (item.role !== 'user' && item.role !== 'assistant') ||
    !Array.isArray(item.content) ||
    item.content.length === 0
  ) {
    fail('invalid_conversation')
  }
  count.parts += item.content.length
  if (count.parts > limits.maxContentParts) fail('request_content_limit_exceeded')
  const content = item.content.map((part): Record<string, unknown> => {
    if (!isRecord(part)) fail('unsupported_input_content')
    if (part.type === 'input_text' || part.type === 'output_text') {
      if (
        !hasOnlyKeys(part, ['type', 'text']) ||
        typeof part.text !== 'string' ||
        part.text === ''
      ) {
        fail('invalid_conversation')
      }
      return { type: 'text', text: part.text }
    }
    if (part.type === 'input_image') {
      if (
        !hasOnlyKeys(part, ['type', 'image_url']) ||
        typeof part.image_url !== 'string' ||
        part.image_url === ''
      ) {
        fail('invalid_conversation')
      }
      return { type: 'image', source: { type: 'url', url: part.image_url } }
    }
    fail('unsupported_input_content')
  })
  return { role: item.role, content }
}

interface PrivateStreamContext {
  usedCallIds: readonly string[]
  allowedExecMethods: readonly string[]
}

function convertResponsesRequest(
  input: unknown,
  limitOverrides: Partial<ProtocolLimits> = {},
): { request: MessagesRequest; context: PrivateStreamContext; limits: ProtocolLimits } {
  const limits = resolveLimits(limitOverrides)
  inspectJsonGraph(input, limits)
  enforceStringLimits(input, limits.maxStringLength)
  const request = validateRequest(input)
  if (
    request.max_output_tokens !== undefined &&
    request.max_output_tokens > limits.maxOutputTokens
  ) {
    fail('max_output_tokens_limit_exceeded')
  }
  if (
    request.prompt_cache_key !== undefined &&
    utf8Length(request.prompt_cache_key) > limits.maxPromptCacheKeyLength
  ) {
    fail('prompt_cache_key_limit_exceeded')
  }
  const sourceItems: unknown[] =
    typeof request.input === 'string'
      ? [{ type: 'message', role: 'user', content: [{ type: 'input_text', text: request.input }] }]
      : request.input
  if (sourceItems.length === 0) fail('invalid_conversation')
  if (sourceItems.length > limits.maxRequestItems) fail('request_item_limit_exceeded')

  const allowedExecMethods = extractAllowedExecMethods(request.client_metadata, limits)
  const upstreamTools: NonNullable<MessagesRequest['tools']> = []

  const developerInstructions: string[] = []
  const conversationItems: unknown[] = []
  const contentCount = { parts: 0 }
  let leading = true
  let sawAdditionalTools = false
  let sawDeveloper = false
  let exposesExec = false
  for (const item of sourceItems) {
    if (isRecord(item) && item.type === 'additional_tools') {
      if (!leading || sawAdditionalTools || sawDeveloper) fail('invalid_conversation')
      validatePinnedAdditionalTools(item, limits)
      sawAdditionalTools = true
      exposesExec = allowedExecMethods.length > 0
      continue
    }
    if (isRecord(item) && item.role === 'developer') {
      if (!leading) fail('invalid_conversation')
      sawDeveloper = true
      if (
        item.type !== 'message' ||
        !hasOnlyKeys(item, ['type', 'id', 'role', 'content']) ||
        !Array.isArray(item.content) ||
        item.content.length === 0
      ) {
        fail('invalid_conversation')
      }
      contentCount.parts += item.content.length
      if (contentCount.parts > limits.maxContentParts) fail('request_content_limit_exceeded')
      for (const part of item.content) {
        if (
          !isRecord(part) ||
          !hasOnlyKeys(part, ['type', 'text']) ||
          part.type !== 'input_text' ||
          typeof part.text !== 'string' ||
          part.text === ''
        ) {
          fail('invalid_conversation')
        }
        developerInstructions.push(part.text)
      }
      continue
    }
    leading = false
    conversationItems.push(item)
  }
  if (exposesExec) {
    upstreamTools.push({
      name: 'exec',
      description: safeExecDescription(allowedExecMethods),
      input_schema: {
        type: 'object',
        properties: { code: { type: 'string' } },
        required: ['code'],
        additionalProperties: false,
      },
    })
  }
  if (upstreamTools.length > limits.maxTools) fail('request_tool_limit_exceeded')

  const messages: MessagesRequest['messages'] = []
  const usedCallIds: string[] = []
  const used = new Set<string>()
  let pending: Array<{ id: string }> | undefined
  let resultIndex = 0
  let sawMessage = false
  const append = (role: 'user' | 'assistant', content: Array<Record<string, unknown>>): void => {
    const previous = messages.at(-1)
    if (previous?.role === role) previous.content.push(...content)
    else messages.push({ role, content: [...content] })
  }

  for (const rawItem of conversationItems) {
    if (!isRecord(rawItem)) fail('unsupported_input_item')
    const isResult = rawItem.type === 'custom_tool_call_output'
    const isCall = rawItem.type === 'custom_tool_call'
    if (pending && resultIndex === pending.length && !isResult) {
      pending = undefined
      resultIndex = 0
    }
    if (rawItem.type === 'reasoning') fail('unsupported_reasoning_input')
    if (rawItem.role === 'developer' || rawItem.type === 'additional_tools')
      fail('invalid_conversation')

    if (rawItem.type === 'message') {
      if (pending) fail('invalid_tool_result_batch')
      const converted = convertMessageContent(rawItem, limits, contentCount)
      if (!sawMessage && converted.role !== 'user') fail('invalid_conversation')
      sawMessage = true
      append(converted.role, converted.content)
      continue
    }
    if (isCall) {
      if (!sawMessage || (pending && resultIndex !== 0)) fail('invalid_conversation')
      if (pending && pending.length >= 1) fail('tool_call_limit_exceeded')
      const id = requireString(rawItem.call_id, 'invalid_tool_call')
      if (id === '') fail('invalid_tool_call')
      if (used.has(id)) fail('duplicate_call_id')
      used.add(id)
      usedCallIds.push(id)
      if (!hasOnlyKeys(rawItem, ['type', 'call_id', 'name', 'input']))
        fail('unsupported_input_item')
      if (rawItem.name !== 'exec' || !exposesExec) fail('unadvertised_tool_call')
      const code = requireString(rawItem.input, 'invalid_custom_tool_input')
      parseSafeExecCode(code, allowedExecMethods, limits)
      pending ??= []
      pending.push({ id })
      append('assistant', [{ type: 'tool_use', id, name: 'exec', input: { code } }])
      continue
    }
    if (isResult) {
      if (!hasOnlyKeys(rawItem, ['type', 'call_id', 'output'])) fail('unsupported_input_item')
      if (!pending || resultIndex >= pending.length) fail('invalid_tool_result_batch')
      const id = requireString(rawItem.call_id, 'invalid_tool_result')
      if (id === '') fail('invalid_tool_result')
      const expected = pending[resultIndex]!
      if (id !== expected.id) fail('invalid_tool_result_batch')
      const output = requireString(rawItem.output, 'invalid_tool_result')
      append('user', [{ type: 'tool_result', tool_use_id: id, content: output }])
      resultIndex += 1
      continue
    }
    fail('unsupported_input_item')
  }
  if (pending && resultIndex !== pending.length) fail('invalid_tool_result_batch')
  if (!sawMessage || messages.length === 0 || messages[0]?.role !== 'user') {
    fail('invalid_conversation')
  }

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
  if (upstreamTools.length > 0) converted.tools = upstreamTools
  if (request.parallel_tool_calls === false && upstreamTools.length > 0) {
    converted.tool_choice = { type: 'auto', disable_parallel_tool_use: true }
  }
  return {
    request: converted,
    context: {
      usedCallIds: Object.freeze([...usedCallIds]),
      allowedExecMethods: Object.freeze([...allowedExecMethods]),
    },
    limits,
  }
}

export function prepareResponsesTurn(
  input: unknown,
  limitOverrides: Partial<ProtocolLimits> = {},
): PreparedResponsesTurn {
  const prepared = convertResponsesRequest(input, limitOverrides)
  return Object.freeze({
    messagesRequest: prepared.request,
    messagesStreamToResponses: (chunks: AsyncIterable<string | Uint8Array>) =>
      convertMessagesStream(chunks, prepared.context, prepared.limits),
  })
}

function sse(event: string, data: UnknownRecord): string {
  return `event: ${event}\ndata: ${JSON.stringify({ type: event, ...data })}\n\n`
}

async function* convertMessagesStream(
  chunks: AsyncIterable<string | Uint8Array>,
  context: PrivateStreamContext,
  limits: ProtocolLimits,
): AsyncGenerator<string> {
  const decoder = new TextDecoder('utf-8', { fatal: true })
  type StrictBlock =
    | { kind: 'text'; itemId: string; text: string }
    | {
        kind: 'tool'
        itemId: string
        callId: string
        name: string
        arguments: string
      }
  type Phase = 'await_start' | 'content' | 'await_stop' | 'terminal'
  const strict = {
    phase: 'await_start' as Phase,
    responseId: undefined as string | undefined,
    active: undefined as StrictBlock | undefined,
    activeIndex: undefined as number | undefined,
    nextIndex: 0,
    inputTokens: 0,
    cachedTokens: 0,
    cacheWriteTokens: 0,
    outputTokens: 0,
    stopReason: undefined as string | undefined,
    output: [] as Array<Record<string, unknown>>,
    usedCalls: new Set(context.usedCallIds),
    toolCalls: 0,
    frames: 0,
    totalOutput: 0,
    terminalFrames: [] as string[],
  }
  let buffer = ''

  const byteLength = (value: string): number => new TextEncoder().encode(value).byteLength
  const usageInteger = (value: unknown, optional: boolean): number => {
    if (value === undefined && optional) return 0
    if (!Number.isSafeInteger(value) || (value as number) < 0) fail('invalid_messages_usage')
    return value as number
  }
  const snapshot = (status: 'in_progress' | 'completed' | 'incomplete'): UnknownRecord => ({
    id: strict.responseId,
    object: 'response',
    model: 'gpt-5.6-sol',
    status,
    output: strict.output,
  })
  const parseStrictFrame = (frame: string): { event: string; data: UnknownRecord } | undefined => {
    let event: string | undefined
    const dataLines: string[] = []
    for (const line of frame.split(/\r\n|\r|\n/)) {
      if (line.startsWith(':')) continue
      if (line.startsWith('event:')) event = line.slice(6).trim()
      else if (line.startsWith('data:')) dataLines.push(line.slice(5).trimStart())
      else if (line !== '') fail('invalid_messages_sse')
    }
    if (event === undefined && dataLines.length === 0) return undefined
    if (dataLines.join('\n') === '[DONE]') fail('unsupported_messages_sse')
    if (event === undefined || dataLines.length === 0) fail('invalid_messages_sse')
    let data: unknown
    try {
      data = JSON.parse(dataLines.join('\n'))
    } catch {
      fail('invalid_messages_sse')
    }
    if (!isRecord(data) || data.type !== event) fail('invalid_messages_sse')
    return { event, data }
  }
  const validateIndex = (value: unknown, starting: boolean): number => {
    if (!Number.isSafeInteger(value) || (value as number) < 0) fail('invalid_messages_block_index')
    const index = value as number
    if ((starting && index !== strict.nextIndex) || (!starting && index !== strict.activeIndex)) {
      fail('invalid_messages_block_index')
    }
    return index
  }
  const processStrictEvent = (event: string, data: UnknownRecord): string[] => {
    if (strict.phase === 'terminal') fail('post_terminal_messages_event')
    if (event === 'error') fail('upstream_error')
    if (event === 'ping') {
      if (!hasOnlyKeys(data, ['type'])) fail('invalid_messages_event')
      return []
    }
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
      if (strict.phase !== 'await_start' || !isRecord(data.message)) {
        fail('invalid_messages_event_order')
      }
      if (!hasOnlyKeys(data, ['type', 'message'])) fail('invalid_messages_event')
      if (
        !hasOnlyKeys(data.message, [
          'id',
          'type',
          'role',
          'content',
          'model',
          'stop_reason',
          'stop_sequence',
          'usage',
        ])
      ) {
        fail('invalid_messages_event')
      }
      if (
        (data.message.type !== undefined && data.message.type !== 'message') ||
        (data.message.role !== undefined && data.message.role !== 'assistant') ||
        (data.message.content !== undefined &&
          (!Array.isArray(data.message.content) || data.message.content.length !== 0)) ||
        (data.message.stop_reason !== undefined && data.message.stop_reason !== null) ||
        (data.message.stop_sequence !== undefined && data.message.stop_sequence !== null)
      ) {
        fail('invalid_messages_event')
      }
      if (data.message.model !== 'openai/gpt-5.6-sol') fail('unsupported_upstream_model')
      strict.responseId = requireString(data.message.id, 'invalid_messages_event')
      if (strict.responseId === '') fail('invalid_messages_event')
      if (!isRecord(data.message.usage)) fail('invalid_messages_usage')
      if (
        !hasOnlyKeys(data.message.usage, [
          'input_tokens',
          'output_tokens',
          'cache_read_input_tokens',
          'cache_creation_input_tokens',
        ])
      ) {
        fail('invalid_messages_event')
      }
      const ordinary = usageInteger(data.message.usage.input_tokens, false)
      usageInteger(data.message.usage.output_tokens, true)
      strict.cachedTokens = usageInteger(data.message.usage.cache_read_input_tokens, true)
      strict.cacheWriteTokens = usageInteger(data.message.usage.cache_creation_input_tokens, true)
      strict.inputTokens = ordinary + strict.cachedTokens + strict.cacheWriteTokens
      if (!Number.isSafeInteger(strict.inputTokens)) fail('invalid_messages_usage')
      strict.phase = 'content'
      return [sse('response.created', { response: snapshot('in_progress') })]
    }
    if (strict.phase === 'await_start') fail('invalid_messages_event_order')
    if (event === 'content_block_start') {
      if (!hasOnlyKeys(data, ['type', 'index', 'content_block'])) fail('invalid_messages_event')
      if (
        strict.phase !== 'content' ||
        strict.active !== undefined ||
        !isRecord(data.content_block)
      ) {
        fail('invalid_messages_event_order')
      }
      if (strict.nextIndex >= limits.maxBlocks) fail('output_block_limit_exceeded')
      const index = validateIndex(data.index, true)
      const itemId = `item_${index}`
      if (
        data.content_block.type === 'thinking' ||
        data.content_block.type === 'redacted_thinking'
      ) {
        fail('unsupported_reasoning_block')
      }
      if (data.content_block.type === 'text') {
        if (!hasOnlyKeys(data.content_block, ['type', 'text']) || data.content_block.text !== '')
          fail('unsupported_messages_event')
        strict.active = { kind: 'text', itemId, text: '' }
        strict.activeIndex = index
        return [
          sse('response.output_item.added', {
            output_index: index,
            item: {
              id: itemId,
              type: 'message',
              role: 'assistant',
              status: 'in_progress',
              content: [],
            },
          }),
          sse('response.content_part.added', {
            item_id: itemId,
            output_index: index,
            content_index: 0,
            part: { type: 'output_text', text: '', annotations: [] },
          }),
        ]
      }
      if (data.content_block.type !== 'tool_use') fail('unsupported_content_block')
      if (
        !hasOnlyKeys(data.content_block, ['type', 'id', 'name', 'input']) ||
        !isRecord(data.content_block.input) ||
        Object.keys(data.content_block.input).length !== 0
      ) {
        fail('unsupported_messages_event')
      }
      const callId = requireString(data.content_block.id, 'invalid_messages_event')
      const name = requireString(data.content_block.name, 'invalid_messages_event')
      if (callId === '') fail('invalid_messages_event')
      if (strict.usedCalls.has(callId)) fail('duplicate_call_id')
      if (name !== 'exec' || context.allowedExecMethods.length === 0) fail('unadvertised_tool_call')
      if (strict.toolCalls >= 1) fail('tool_call_limit_exceeded')
      strict.toolCalls += 1
      strict.usedCalls.add(callId)
      strict.active = { kind: 'tool', itemId, callId, name, arguments: '' }
      strict.activeIndex = index
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
    if (event === 'content_block_delta') {
      if (!hasOnlyKeys(data, ['type', 'index', 'delta'])) fail('invalid_messages_event')
      if (strict.phase !== 'content' || strict.active === undefined) {
        fail('invalid_messages_event_order')
      }
      const index = validateIndex(data.index, false)
      if (!isRecord(data.delta)) fail('invalid_messages_event')
      if (strict.active.kind === 'text' && data.delta.type === 'text_delta') {
        if (!hasOnlyKeys(data.delta, ['type', 'text'])) fail('invalid_messages_event')
        const deltaText = requireString(data.delta.text, 'invalid_messages_event')
        strict.active.text += deltaText
        if (utf8Length(strict.active.text) > limits.maxAccumulatedText)
          fail('output_text_limit_exceeded')
        return [
          sse('response.output_text.delta', {
            item_id: strict.active.itemId,
            output_index: index,
            content_index: 0,
            delta: deltaText,
          }),
        ]
      }
      if (strict.active.kind === 'tool' && data.delta.type === 'input_json_delta') {
        if (!hasOnlyKeys(data.delta, ['type', 'partial_json'])) fail('invalid_messages_event')
        const argumentDelta = requireString(data.delta.partial_json, 'invalid_messages_event')
        strict.active.arguments += argumentDelta
        if (utf8Length(strict.active.arguments) > limits.maxToolArguments) {
          fail('tool_arguments_limit_exceeded')
        }
        return []
      }
      fail('unsupported_content_delta')
    }
    if (event === 'content_block_stop') {
      if (!hasOnlyKeys(data, ['type', 'index'])) fail('invalid_messages_event')
      if (strict.phase !== 'content' || strict.active === undefined) {
        fail('invalid_messages_event_order')
      }
      const index = validateIndex(data.index, false)
      const block = strict.active
      strict.active = undefined
      strict.activeIndex = undefined
      strict.nextIndex += 1
      if (block.kind === 'text') {
        const part = { type: 'output_text', text: block.text, annotations: [] }
        const item = {
          id: block.itemId,
          type: 'message',
          role: 'assistant',
          status: 'completed',
          content: [part],
        }
        strict.output.push(item)
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
      let parsed: unknown
      try {
        parsed = JSON.parse(block.arguments)
      } catch {
        fail('invalid_custom_tool_input')
      }
      if (!isRecord(parsed) || !hasOnlyKeys(parsed, ['code']) || typeof parsed.code !== 'string')
        fail('invalid_custom_tool_input')
      parseSafeExecCode(parsed.code, context.allowedExecMethods, limits)
      const item = {
        id: block.itemId,
        type: 'custom_tool_call',
        status: 'completed',
        call_id: block.callId,
        name: 'exec',
        input: parsed.code,
      }
      strict.output.push(item)
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
    if (event === 'message_delta') {
      if (!hasOnlyKeys(data, ['type', 'delta', 'usage'])) fail('invalid_messages_event')
      if (strict.phase !== 'content' || strict.active !== undefined || !isRecord(data.delta)) {
        fail('invalid_messages_event_order')
      }
      if (!isRecord(data.usage)) fail('invalid_messages_usage')
      if (!hasOnlyKeys(data.delta, ['stop_reason', 'stop_sequence'])) fail('invalid_messages_event')
      if (
        data.delta.stop_sequence !== undefined &&
        data.delta.stop_sequence !== null &&
        typeof data.delta.stop_sequence !== 'string'
      ) {
        fail('invalid_messages_event')
      }
      if (!hasOnlyKeys(data.usage, ['output_tokens'])) fail('invalid_messages_event')
      strict.stopReason = requireString(data.delta.stop_reason, 'invalid_messages_event')
      strict.outputTokens = usageInteger(data.usage.output_tokens, false)
      strict.phase = 'await_stop'
      return []
    }
    if (event === 'message_stop') {
      if (!hasOnlyKeys(data, ['type'])) fail('invalid_messages_event')
      if (strict.phase !== 'await_stop' || strict.stopReason === undefined) {
        fail('invalid_messages_event_order')
      }
      const completed = ['end_turn', 'stop_sequence', 'tool_use'].includes(strict.stopReason)
      if (!completed && strict.stopReason !== 'max_tokens') fail('unsupported_stop_reason')
      const response = {
        ...snapshot(completed ? 'completed' : 'incomplete'),
        ...(completed ? {} : { incomplete_details: { reason: 'max_output_tokens' } }),
        usage: {
          input_tokens: strict.inputTokens,
          output_tokens: strict.outputTokens,
          total_tokens: (() => {
            const total = strict.inputTokens + strict.outputTokens
            if (!Number.isSafeInteger(total)) fail('invalid_messages_usage')
            return total
          })(),
          input_tokens_details: {
            cached_tokens: strict.cachedTokens,
            cache_write_tokens: strict.cacheWriteTokens,
          },
          output_tokens_details: { reasoning_tokens: 0 },
        },
      }
      strict.phase = 'terminal'
      strict.terminalFrames = [
        sse(completed ? 'response.completed' : 'response.incomplete', { response }),
        'data: [DONE]\n\n',
      ]
      return []
    }
    fail('unsupported_messages_event')
  }

  const yieldBounded = async function* (frames: string[]): AsyncGenerator<string> {
    const addedBytes = frames.reduce((total, frame) => total + byteLength(frame), 0)
    if (strict.totalOutput + addedBytes > limits.maxTotalOutput) {
      fail('total_output_limit_exceeded')
    }
    for (const frame of frames) {
      strict.totalOutput += byteLength(frame)
      yield frame
    }
  }

  for await (const chunk of chunks) {
    try {
      buffer += typeof chunk === 'string' ? chunk : decoder.decode(chunk, { stream: true })
    } catch {
      fail('invalid_messages_utf8')
    }
    if (byteLength(buffer) > limits.maxSseBufferBytes) fail('sse_buffer_limit_exceeded')
    while (true) {
      const boundary = /(?:\r\n|\r(?!\n)|\n)(?:\r\n|\r(?!\n)|\n)/.exec(buffer)
      if (!boundary || boundary.index === undefined) break
      const frame = buffer.slice(0, boundary.index)
      buffer = buffer.slice(boundary.index + boundary[0].length)
      strict.frames += 1
      if (strict.frames > limits.maxSseFrames) fail('sse_frame_count_limit_exceeded')
      if (byteLength(frame) > limits.maxSseFrameBytes) fail('sse_frame_limit_exceeded')
      const parsed = parseStrictFrame(frame)
      if (parsed === undefined) continue
      for await (const converted of yieldBounded(processStrictEvent(parsed.event, parsed.data))) {
        yield converted
      }
    }
  }
  try {
    buffer += decoder.decode()
  } catch {
    fail('invalid_messages_utf8')
  }
  if (buffer !== '') fail('invalid_messages_sse')
  if (strict.phase !== 'terminal') fail('premature_messages_eof')
  for await (const terminal of yieldBounded(strict.terminalFrames)) yield terminal
}
