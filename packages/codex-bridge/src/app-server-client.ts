import { isAbsolute } from 'node:path'
import type {
  CodexAppServerNotification,
  InitializeParams,
  InitializeResponse,
  ThreadStartParams,
  ThreadStartResponse,
  TurnInterruptParams,
  TurnInterruptResponse,
  TurnStartParams,
  TurnStartResponse,
} from './generated/index.js'
import { KNOWN_SERVER_NOTIFICATION_METHODS } from './generated/index.js'
import { JsonRpcClient, JsonRpcError, type JsonRpcNotification } from './json-rpc.js'

const KNOWN_NOTIFICATIONS = new Set<string>(KNOWN_SERVER_NOTIFICATION_METHODS)

export interface CodexAppServerClientOptions {
  readonly rpc: JsonRpcClient
  readonly cwd: string
  readonly developerInstructions: string
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

function hasNonemptyId(value: unknown, key: string): boolean {
  return isRecord(value) && typeof value[key] === 'string' && value[key] !== ''
}

export class CodexAppServerClient {
  readonly #rpc: JsonRpcClient
  readonly #cwd: string
  readonly #developerInstructions: string
  readonly #notificationSubscribers = new Set<(notification: CodexAppServerNotification) => void>()
  readonly #unsubscribeRpc: () => void
  #state: 'new' | 'initializing' | 'ready' | 'closed' | 'failed' = 'new'
  #failureCode: string | undefined
  #shutdownPromise: Promise<void> | undefined

  constructor(options: CodexAppServerClientOptions) {
    if (
      !options?.rpc ||
      typeof options.developerInstructions !== 'string' ||
      options.developerInstructions === '' ||
      typeof options.cwd !== 'string' ||
      !isAbsolute(options.cwd)
    ) {
      throw new TypeError('invalid_app_server_client_options')
    }
    this.#rpc = options.rpc
    this.#cwd = options.cwd
    this.#developerInstructions = options.developerInstructions
    this.#unsubscribeRpc = this.#rpc.subscribe((notification) =>
      this.#handleNotification(notification),
    )
  }

  async initialize(): Promise<InitializeResponse> {
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
      const response = await this.#rpc.request<InitializeResponse>('initialize', params)
      if (
        !isRecord(response) ||
        !['userAgent', 'codexHome', 'platformFamily', 'platformOs'].every(
          (key) => typeof response[key] === 'string' && response[key] !== '',
        )
      ) {
        this.#protocolError()
      }
      this.#rpc.notify('initialized', {})
      this.#state = 'ready'
      return response
    } catch (error) {
      this.#state = 'failed'
      throw error
    }
  }

  async startThread(): Promise<ThreadStartResponse> {
    this.#assertReady()
    const params: ThreadStartParams = {
      model: 'gpt-5.6-sol',
      modelProvider: 'wiswork',
      cwd: this.#cwd,
      approvalPolicy: 'never',
      sandbox: 'read-only',
      developerInstructions: this.#developerInstructions,
      ephemeral: true,
    }
    const response = await this.#rpc.request<ThreadStartResponse>('thread/start', params)
    if (!isRecord(response) || !hasNonemptyId(response.thread, 'id')) {
      this.#protocolError()
    }
    return response
  }

  async startTurn(
    threadId: string,
    text: string,
    options: { readonly signal?: AbortSignal } = {},
  ): Promise<TurnStartResponse> {
    this.#assertReady()
    if (threadId === '' || text === '') throw new CodexAppServerError('app_server_invalid_argument')
    const params: TurnStartParams = {
      threadId,
      input: [{ type: 'text', text, text_elements: [] }],
      effort: 'medium',
    }
    const response = await this.#rpc.request<TurnStartResponse>('turn/start', params, options)
    if (!isRecord(response) || !hasNonemptyId(response.turn, 'id')) {
      this.#protocolError()
    }
    return response
  }

  async interruptTurn(threadId: string, turnId: string): Promise<TurnInterruptResponse> {
    this.#assertReady()
    if (threadId === '' || turnId === '') {
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
      void this.shutdown('app_server_protocol_error')
      return
    }
    if (!this.#isValidImportantNotification(notification)) {
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

  #isValidImportantNotification(notification: JsonRpcNotification): boolean {
    if (notification.method === 'item/agentMessage/delta') {
      return (
        isRecord(notification.params) &&
        ['threadId', 'turnId', 'itemId', 'delta'].every(
          (key) => typeof (notification.params as Record<string, unknown>)[key] === 'string',
        )
      )
    }
    if (notification.method === 'turn/started' || notification.method === 'turn/completed') {
      return (
        isRecord(notification.params) &&
        typeof notification.params.threadId === 'string' &&
        hasNonemptyId(notification.params.turn, 'id')
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
