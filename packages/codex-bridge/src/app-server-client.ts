import { isAbsolute } from 'node:path'
import type {
  CodexAppServerNotification,
  InitializeParams,
  ThreadStartParams,
  TurnInterruptParams,
  TurnInterruptResponse,
  TurnStartParams,
} from './generated/index.js'
import { KNOWN_SERVER_NOTIFICATION_METHODS } from './generated/index.js'
import { CODEX_CLI_VERSION } from './generated/index.js'
import { JsonRpcClient, JsonRpcError, type JsonRpcNotification } from './json-rpc.js'

const KNOWN_NOTIFICATIONS = new Set<string>(KNOWN_SERVER_NOTIFICATION_METHODS)
// Only these lifecycle events cross the bridge. Every other version-known event is
// parsed as a bounded data object and discarded; it cannot expand host capabilities.
const FORWARDED_NOTIFICATIONS = new Set<string>([
  'item/agentMessage/delta',
  'turn/started',
  'turn/completed',
])
const INITIALIZE_RESPONSE_KEYS = ['userAgent', 'codexHome', 'platformFamily', 'platformOs'] as const
const THREAD_START_RESPONSE_KEYS = [
  'thread',
  'model',
  'modelProvider',
  'serviceTier',
  'cwd',
  'instructionSources',
  'approvalPolicy',
  'approvalsReviewer',
  'sandbox',
  'reasoningEffort',
  'activePermissionProfile',
  'multiAgentMode',
  'runtimeWorkspaceRoots',
] as const
const THREAD_START_REQUIRED_KEYS = [
  'thread',
  'model',
  'modelProvider',
  'cwd',
  'approvalPolicy',
  'approvalsReviewer',
  'sandbox',
] as const
const MAX_IDENTIFIER_BYTES = 256
const MAX_PROMPT_BYTES = 1_000_000
const MAX_POLICY_BYTES = 65_536
const MAX_PATH_BYTES = 4_096
const MAX_DELTA_BYTES = 1_000_000
const PINNED_VERSION = CODEX_CLI_VERSION.slice('codex-cli '.length)
const PINNED_TURN_COMPLETION_STATUSES = new Set(['completed', 'interrupted', 'failed'])

export interface CodexAppServerClientOptions {
  readonly rpc: JsonRpcClient
  readonly cwd: string
  readonly developerInstructions: string
  readonly diagnostics?: (code: string) => void
}

export class CodexAppServerError extends Error {
  constructor(readonly code: string) {
    super(code)
    this.name = 'CodexAppServerError'
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function hasExactKeys(value: Record<string, unknown>, expected: readonly string[]): boolean {
  const actual = Object.keys(value).sort()
  const sorted = [...expected].sort()
  return actual.length === sorted.length && actual.every((key, index) => key === sorted[index])
}

function hasAllowedRequiredKeys(
  value: Record<string, unknown>,
  allowed: readonly string[],
  required: readonly string[],
): boolean {
  const keys = Object.keys(value)
  return keys.every((key) => allowed.includes(key)) && required.every((key) => keys.includes(key))
}

export type CodexInitializeResult = Readonly<{
  userAgent: string
  platformFamily: string
  platformOs: string
}>
export type CodexThreadStartResult = Readonly<{ thread: Readonly<{ id: string }> }>
export type CodexTurnStartResult = Readonly<{ turn: Readonly<{ id: string }> }>

function hasNonemptyId(value: unknown, key: string): boolean {
  return isRecord(value) && boundedString(value[key], MAX_IDENTIFIER_BYTES)
}

function boundedString(value: unknown, maximumBytes: number): value is string {
  return (
    typeof value === 'string' && value !== '' && Buffer.byteLength(value, 'utf8') <= maximumBytes
  )
}

export class CodexAppServerClient {
  readonly #rpc: JsonRpcClient
  readonly #cwd: string
  readonly #developerInstructions: string
  readonly #notificationSubscribers = new Set<(notification: CodexAppServerNotification) => void>()
  readonly #diagnostics?: (code: string) => void
  readonly #unsubscribeRpc: () => void
  #state: 'new' | 'initializing' | 'ready' | 'closed' | 'failed' = 'new'
  #failureCode: string | undefined
  #shutdownPromise: Promise<void> | undefined

  constructor(options: CodexAppServerClientOptions) {
    if (
      !options?.rpc ||
      !boundedString(options.developerInstructions, MAX_POLICY_BYTES) ||
      !boundedString(options.cwd, MAX_PATH_BYTES) ||
      !isAbsolute(options.cwd)
    ) {
      throw new TypeError('invalid_app_server_client_options')
    }
    this.#rpc = options.rpc
    this.#cwd = options.cwd
    this.#developerInstructions = options.developerInstructions
    this.#diagnostics = options.diagnostics
    this.#unsubscribeRpc = this.#rpc.subscribe((notification) =>
      this.#handleNotification(notification),
    )
  }

  async initialize(): Promise<CodexInitializeResult> {
    if (this.#state !== 'new') {
      throw new CodexAppServerError(
        this.#state === 'closed' ? 'app_server_closed' : 'app_server_already_initialized',
      )
    }
    this.#state = 'initializing'
    const params: InitializeParams = {
      clientInfo: { name: 'wiswork', version: '0.1.0' },
      capabilities: null,
    }
    try {
      const response = await this.#rpc.request<unknown>('initialize', params)
      if (
        !isRecord(response) ||
        !hasExactKeys(response, INITIALIZE_RESPONSE_KEYS) ||
        !['userAgent', 'codexHome', 'platformFamily', 'platformOs'].every((key) =>
          boundedString(response[key], key === 'codexHome' ? MAX_PATH_BYTES : MAX_IDENTIFIER_BYTES),
        ) ||
        !(response.userAgent as string).includes(PINNED_VERSION)
      ) {
        this.#protocolError()
      }
      this.#rpc.notify('initialized', {})
      this.#state = 'ready'
      return {
        userAgent: response.userAgent as string,
        platformFamily: response.platformFamily as string,
        platformOs: response.platformOs as string,
      }
    } catch (error) {
      this.#state = 'failed'
      throw error
    }
  }

  async startThread(
    options: { readonly developerInstructions?: string } = {},
  ): Promise<CodexThreadStartResult> {
    this.#assertReady()
    const developerInstructions = options.developerInstructions ?? this.#developerInstructions
    if (!boundedString(developerInstructions, MAX_POLICY_BYTES))
      throw new CodexAppServerError('app_server_invalid_argument')
    const params: ThreadStartParams = {
      model: 'gpt-5.6-sol',
      modelProvider: 'wiswork',
      cwd: this.#cwd,
      approvalPolicy: 'never',
      sandbox: 'read-only',
      developerInstructions,
      ephemeral: true,
    }
    const response = await this.#rpc.request<unknown>('thread/start', params)
    if (
      !isRecord(response) ||
      !hasAllowedRequiredKeys(response, THREAD_START_RESPONSE_KEYS, THREAD_START_REQUIRED_KEYS) ||
      !hasNonemptyId(response.thread, 'id')
    ) {
      this.#protocolError()
    }
    return { thread: { id: (response.thread as Record<string, unknown>).id as string } }
  }

  async startTurn(
    threadId: string,
    text: string,
    options: { readonly signal?: AbortSignal } = {},
  ): Promise<CodexTurnStartResult> {
    this.#assertReady()
    if (!boundedString(threadId, MAX_IDENTIFIER_BYTES) || !boundedString(text, MAX_PROMPT_BYTES))
      throw new CodexAppServerError('app_server_invalid_argument')
    const params: TurnStartParams = {
      threadId,
      input: [{ type: 'text', text, text_elements: [] }],
      effort: 'medium',
    }
    const response = await this.#rpc.request<unknown>('turn/start', params, options)
    if (
      !isRecord(response) ||
      !hasExactKeys(response, ['turn']) ||
      !hasNonemptyId(response.turn, 'id')
    ) {
      this.#protocolError()
    }
    return { turn: { id: (response.turn as Record<string, unknown>).id as string } }
  }

  async interruptTurn(threadId: string, turnId: string): Promise<TurnInterruptResponse> {
    this.#assertReady()
    if (
      !boundedString(threadId, MAX_IDENTIFIER_BYTES) ||
      !boundedString(turnId, MAX_IDENTIFIER_BYTES)
    ) {
      throw new CodexAppServerError('app_server_invalid_argument')
    }
    const params: TurnInterruptParams = { threadId, turnId }
    const response = await this.#rpc.request<TurnInterruptResponse>('turn/interrupt', params)
    if (!isRecord(response) || Object.keys(response).length !== 0) {
      this.#protocolError()
    }
    return response
  }

  onNotification(listener: (notification: CodexAppServerNotification) => void): () => void {
    if (typeof listener !== 'function') throw new TypeError('invalid_notification_subscriber')
    if (this.#state === 'closed' || this.#state === 'failed') return () => undefined
    this.#notificationSubscribers.add(listener)
    let active = true
    return () => {
      if (!active) return
      active = false
      this.#notificationSubscribers.delete(listener)
    }
  }

  shutdown(code = 'app_server_closed'): Promise<void> {
    if (this.#shutdownPromise) return this.#shutdownPromise
    if (code !== 'app_server_closed') this.#failureCode = code
    this.#state = 'closed'
    this.#unsubscribeRpc()
    this.#notificationSubscribers.clear()
    this.#shutdownPromise = this.#rpc.close(code)
    return this.#shutdownPromise
  }

  #assertReady(): void {
    if (this.#state === 'ready') return
    if (this.#state === 'closed' || this.#state === 'failed') {
      throw new CodexAppServerError(this.#failureCode ?? 'app_server_closed')
    }
    throw new CodexAppServerError('app_server_not_initialized')
  }

  #handleNotification(notification: JsonRpcNotification): void {
    if (!KNOWN_NOTIFICATIONS.has(notification.method)) {
      this.#diagnostic('app_server_notification_unknown')
      void this.shutdown('app_server_protocol_error')
      return
    }
    this.#diagnostic(`app_server_${notification.method.replaceAll('/', '_')}`)
    if (notification.method === 'warning') {
      const message = isRecord(notification.params) ? notification.params.message : undefined
      if (typeof message === 'string') {
        const normalized = message.toLowerCase()
        if (normalized.includes('code mode')) this.#diagnostic('app_server_warning_code_mode')
        if (normalized.includes('host')) this.#diagnostic('app_server_warning_host')
        if (normalized.includes('sandbox')) this.#diagnostic('app_server_warning_sandbox')
        if (normalized.includes('install')) this.#diagnostic('app_server_warning_install')
        if (normalized.includes('bwrap')) this.#diagnostic('app_server_warning_bwrap')
      }
    }
    if (notification.method === 'item/started' || notification.method === 'item/completed') {
      const item = isRecord(notification.params) ? notification.params.item : undefined
      const itemType = isRecord(item) ? item.type : undefined
      if (typeof itemType === 'string' && /^[A-Za-z][A-Za-z0-9]{0,63}$/.test(itemType)) {
        this.#diagnostic(`app_server_item_type_${itemType}`)
      }
      const itemStatus = isRecord(item) ? item.status : undefined
      if (typeof itemStatus === 'string' && /^[a-z][a-zA-Z]{0,31}$/.test(itemStatus)) {
        this.#diagnostic(`app_server_item_status_${itemStatus}`)
      }
      if (isRecord(item) && item.server === 'wiswork')
        this.#diagnostic('app_server_mcp_server_exact')
      if (
        isRecord(item) &&
        (item.tool === 'wiswork_read' ||
          item.name === 'wiswork_read' ||
          item.tool === 'wiswork_propose' ||
          item.name === 'wiswork_propose')
      )
        this.#diagnostic('app_server_mcp_tool_exact')
      if (
        isRecord(item) &&
        [item.callId, item.call_id, item.id].some(
          (value) => typeof value === 'string' && value.length > 0 && value.length <= 256,
        )
      )
        this.#diagnostic('app_server_mcp_call_id_present')
      if (isRecord(item) && item.error !== undefined && item.error !== null)
        this.#diagnostic('app_server_mcp_error_present')
      const errorCode =
        isRecord(item) && isRecord(item.error)
          ? item.error.code
          : isRecord(item) && typeof item.error === 'string'
            ? item.error
            : undefined
      if (typeof errorCode === 'string' && /^[a-z][a-z0-9_]{0,63}$/.test(errorCode))
        this.#diagnostic(`app_server_mcp_error_${errorCode}`)
      if (
        typeof errorCode === 'number' &&
        Number.isSafeInteger(errorCode) &&
        errorCode >= -32_768 &&
        errorCode <= 32_767
      )
        this.#diagnostic(`app_server_mcp_error_code_${errorCode}`)
      const errorMessage =
        isRecord(item) && isRecord(item.error) && typeof item.error.message === 'string'
          ? item.error.message.toLowerCase()
          : isRecord(item) && typeof item.error === 'string'
            ? item.error.toLowerCase()
            : ''
      for (const [needle, code] of [
        ['method not found', 'method_not_found'],
        ['invalid params', 'invalid_params'],
        ['transport', 'transport'],
        ['timed out', 'timeout'],
        ['cancelled', 'cancelled'],
        ['connection closed', 'connection_closed'],
        ['tool error', 'tool_error'],
        ['code mode', 'code_mode'],
      ] as const) {
        if (errorMessage.includes(needle)) this.#diagnostic(`app_server_mcp_error_${code}`)
      }
    }
    if (notification.method === 'thread/status/changed') {
      const status = isRecord(notification.params) ? notification.params.status : undefined
      const statusType = isRecord(status) ? status.type : undefined
      if (typeof statusType === 'string' && /^[a-z][a-zA-Z]{0,31}$/.test(statusType)) {
        this.#diagnostic(`app_server_thread_status_${statusType}`)
      }
    }
    if (!FORWARDED_NOTIFICATIONS.has(notification.method)) {
      if (!isRecord(notification.params)) {
        this.#diagnostic('app_server_notification_invalid')
        void this.shutdown('app_server_protocol_error')
      }
      return
    }
    if (!this.#isValidImportantNotification(notification)) {
      this.#diagnostic('app_server_notification_invalid')
      void this.shutdown('app_server_protocol_error')
      return
    }
    for (const listener of [...this.#notificationSubscribers]) {
      try {
        listener(notification as CodexAppServerNotification)
      } catch {
        // Notification consumers cannot break transport or other consumers.
      }
    }
  }

  #diagnostic(code: string): void {
    try {
      this.#diagnostics?.(code)
    } catch {}
  }

  #isValidImportantNotification(notification: JsonRpcNotification): boolean {
    if (notification.method === 'item/agentMessage/delta') {
      return (
        isRecord(notification.params) &&
        ['threadId', 'turnId', 'itemId', 'delta'].every((key) =>
          boundedString(
            (notification.params as Record<string, unknown>)[key],
            key === 'delta' ? MAX_DELTA_BYTES : MAX_IDENTIFIER_BYTES,
          ),
        )
      )
    }
    if (notification.method === 'turn/started' || notification.method === 'turn/completed') {
      const turn = isRecord(notification.params) ? notification.params.turn : undefined
      return (
        isRecord(notification.params) &&
        boundedString(notification.params.threadId, MAX_IDENTIFIER_BYTES) &&
        hasNonemptyId(turn, 'id') &&
        (notification.method !== 'turn/completed' ||
          (isRecord(turn) && PINNED_TURN_COMPLETION_STATUSES.has(turn.status as string)))
      )
    }
    return true
  }

  #protocolError(): never {
    void this.shutdown('app_server_protocol_error')
    throw new CodexAppServerError('app_server_protocol_error')
  }
}

export function asCodexProcessError(error: unknown, fallback: string): CodexAppServerError {
  if (error instanceof CodexAppServerError) return error
  if (error instanceof JsonRpcError) return new CodexAppServerError(error.code)
  return new CodexAppServerError(fallback)
}
