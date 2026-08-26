import { createHash } from 'node:crypto'
import { isAbsolute } from 'node:path'
import type { AgentToolCall, ToolExecution } from '@wiswork/agent-core'
import type { AgentEvent, AgentRuntimeKind } from '@wiswork/agent-runtime'
import { WISWORK_MESSAGES_URL, WISWORK_REQUEST_LOCATION } from '@wiswork/ai-provider'
import type { AuthClient } from '@wiswork/auth'
import {
  CodexProcessManager,
  startDocumentMcpServer,
  startResponsesBridge,
  type CodexAppServerClient,
  type CodexAppServerNotification,
  type CodexProcessManagerOptions,
  type DocumentMcpServer,
  type DocumentMcpSession,
  type MessagesRequest,
  type ResponsesBridge,
  type ToolSessionRegistration,
} from '@wiswork/codex-bridge'
import {
  CODEX_RUNTIME_CHANNELS,
  type CodexRuntimeStartRequest,
  type CodexRuntimeStatus,
} from '../shared/codex-api'

const MAX_DOCUMENT_ID_BYTES = 256
const MAX_TURN_TEXT_BYTES = 1_000_000
const MAX_ACTIVE_DOCUMENTS = 8
const FINISHED_TURN_IDS_LIMIT = 16
const FIXED_CODEX_MODEL = 'openai/gpt-5.6-sol'

export const WISWORK_CODEX_DEVELOPER_INSTRUCTIONS =
  'You are the WisWork document assistant. Use only the authenticated document tools provided by WisWork. Never access the shell, filesystem, network, credentials, or another document.'

interface CodexOwner {
  isDestroyed(): boolean
  send(channel: string, payload: unknown): void
}

interface AuthClientLike {
  getValidAccountStatus(): Promise<{ readonly loggedIn: boolean }>
  fetchWithAuth(request: (accessToken: string) => Promise<Response>): Promise<Response>
}

interface ProcessManagerLike {
  readonly crashed: Promise<unknown>
  start(): Promise<CodexAppServerClient>
  stop(): Promise<void>
}

interface IpcMainLike {
  handle(
    channel: string,
    handler: (event: { sender: CodexOwner }, ...args: unknown[]) => unknown,
  ): void
}

export interface ShellCodexRuntimeDiagnostic {
  readonly code: string
}

export interface ShellCodexRuntimeOptions {
  readonly runtimeKind: AgentRuntimeKind
  readonly executablePath?: string
  readonly authClient: AuthClientLike
  readonly startResponsesBridge?: typeof startResponsesBridge
  readonly startDocumentMcpServer?: typeof startDocumentMcpServer
  readonly createProcessManager?: (options: CodexProcessManagerOptions) => ProcessManagerLike
  /** Synchronously revokes renderer tool IPC when the child terminates unexpectedly. */
  readonly onProcessCrash?: (documentId: string) => void
  readonly diagnostics?: (diagnostic: ShellCodexRuntimeDiagnostic) => void
}

export interface RegisterCodexDocumentInput {
  readonly documentId: string
  readonly owner: CodexOwner
  readonly registration: ToolSessionRegistration
}

export interface RegisteredCodexDocument {
  close(): Promise<void>
}

interface ActiveTurn {
  turnId: string | undefined
  text: string
  cancelRequested: boolean
  terminal: boolean
  interruptPromise?: Promise<void>
  toolCall?: { fingerprint: string; approvalStarted: boolean; executed: boolean }
}

interface DocumentRecord extends RegisterCodexDocumentInput {
  state: 'starting' | 'ready' | 'closing' | 'closed'
  startPromise: Promise<void>
  closePromise?: Promise<void>
  bridge?: ResponsesBridge
  mcp?: DocumentMcpServer
  mcpSession?: DocumentMcpSession
  manager?: ProcessManagerLike
  client?: CodexAppServerClient
  threadId?: string
  unsubscribe?: () => void
  activeTurn?: ActiveTurn
  quiesced: boolean
  readonly finishedTurnIds: Set<string>
}

export class ShellCodexRuntimeError extends Error {
  constructor(readonly code: string) {
    super(code)
    this.name = 'ShellCodexRuntimeError'
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

function hasOnlyKeys(value: Record<string, unknown>, keys: readonly string[]): boolean {
  const allowed = new Set(keys)
  return Object.keys(value).every((key) => allowed.has(key))
}

function boundedNonempty(value: unknown, maximum: number): value is string {
  return typeof value === 'string' && value !== '' && Buffer.byteLength(value) <= maximum
}

function asRuntimeError(error: unknown, fallback: string): ShellCodexRuntimeError {
  return error instanceof ShellCodexRuntimeError ? error : new ShellCodexRuntimeError(fallback)
}

function safeTurnStatus(turn: unknown): 'completed' | 'interrupted' | 'failed' | undefined {
  if (!isRecord(turn)) return undefined
  return turn.status === 'completed' || turn.status === 'interrupted' || turn.status === 'failed'
    ? turn.status
    : undefined
}

/**
 * The only production adapter from the local Responses bridge to WisUsage.
 * Authentication stays inside AuthClient and the bridge-selected model is rechecked here.
 */
export function createWisUsageBridgeFetch(
  authClient: Pick<AuthClient, 'fetchWithAuth'>,
  upstreamFetch: typeof globalThis.fetch = globalThis.fetch,
): (request: MessagesRequest, signal: AbortSignal) => Promise<Response> {
  return async (request, signal) => {
    if (request.model !== FIXED_CODEX_MODEL) {
      throw new ShellCodexRuntimeError('codex_model_mismatch')
    }
    if (signal.aborted) throw new ShellCodexRuntimeError('codex_turn_cancelled')
    return authClient.fetchWithAuth((accessToken) => {
      if (signal.aborted) throw new ShellCodexRuntimeError('codex_turn_cancelled')
      return upstreamFetch(WISWORK_MESSAGES_URL, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
          'x-req-location': WISWORK_REQUEST_LOCATION,
        },
        body: JSON.stringify(request),
        signal,
      })
    })
  }
}

export class ShellCodexRuntime {
  readonly #runtimeKind: AgentRuntimeKind
  readonly #executablePath: string | undefined
  readonly #authClient: AuthClientLike
  readonly #startResponsesBridge: typeof startResponsesBridge
  readonly #startDocumentMcpServer: typeof startDocumentMcpServer
  readonly #createProcessManager: (options: CodexProcessManagerOptions) => ProcessManagerLike
  readonly #onProcessCrash?: (documentId: string) => void
  readonly #diagnostics?: (diagnostic: ShellCodexRuntimeDiagnostic) => void
  readonly #documents = new Map<string, DocumentRecord>()
  readonly #documentCloseGates = new Set<string>()
  #globalCloseGate = false
  #closed = false
  #acceptingDocuments = true
  #logoutPromise: Promise<void> | undefined
  #shutdownPromise: Promise<void> | undefined

  constructor(options: ShellCodexRuntimeOptions) {
    if (
      !options ||
      (options.runtimeKind !== 'legacy' && options.runtimeKind !== 'codex') ||
      !options.authClient ||
      typeof options.authClient.getValidAccountStatus !== 'function' ||
      typeof options.authClient.fetchWithAuth !== 'function'
    ) {
      throw new TypeError('invalid_shell_codex_runtime_options')
    }
    this.#runtimeKind = options.runtimeKind
    this.#executablePath = options.executablePath
    this.#authClient = options.authClient
    this.#startResponsesBridge = options.startResponsesBridge ?? startResponsesBridge
    this.#startDocumentMcpServer = options.startDocumentMcpServer ?? startDocumentMcpServer
    this.#createProcessManager =
      options.createProcessManager ?? ((processOptions) => new CodexProcessManager(processOptions))
    this.#onProcessCrash = options.onProcessCrash
    this.#diagnostics = options.diagnostics
  }

  get runtimeKind(): AgentRuntimeKind {
    return this.#runtimeKind
  }

  ownsDocument(owner: CodexOwner, documentId: string): boolean {
    const record = this.#documents.get(documentId)
    return record?.owner === owner && record.state !== 'closed' && !owner.isDestroyed()
  }

  async registerDocument(input: RegisterCodexDocumentInput): Promise<RegisteredCodexDocument> {
    if (this.#runtimeKind !== 'codex') throw new ShellCodexRuntimeError('codex_runtime_disabled')
    if (this.#closed) throw new ShellCodexRuntimeError('codex_runtime_closed')
    if (!this.#acceptingDocuments) throw new ShellCodexRuntimeError('codex_runtime_busy')
    if (
      !input ||
      !boundedNonempty(input.documentId, MAX_DOCUMENT_ID_BYTES) ||
      !input.owner ||
      input.owner.isDestroyed()
    ) {
      throw new ShellCodexRuntimeError('codex_document_unavailable')
    }
    if (!this.#executablePath || !isAbsolute(this.#executablePath)) {
      throw new ShellCodexRuntimeError('codex_executable_unavailable')
    }
    if (this.#globalCloseGate || this.#documentCloseGates.has(input.documentId)) {
      throw new ShellCodexRuntimeError('codex_runtime_busy')
    }
    if (this.#documents.has(input.documentId)) {
      throw new ShellCodexRuntimeError('codex_document_exists')
    }
    if (this.#documents.size >= MAX_ACTIVE_DOCUMENTS) {
      throw new ShellCodexRuntimeError('codex_document_limit')
    }
    const record: DocumentRecord = {
      ...input,
      state: 'starting',
      startPromise: Promise.resolve(),
      finishedTurnIds: new Set(),
      quiesced: false,
    }
    this.#documents.set(input.documentId, record)
    record.startPromise = this.#startDocument(record)
    try {
      await record.startPromise
      if (record.state !== 'ready') throw new ShellCodexRuntimeError('codex_session_closed')
      return { close: () => this.closeDocument(input.documentId) }
    } catch (error) {
      if (this.#documents.get(input.documentId) === record && record.state !== 'closing') {
        this.#documents.delete(input.documentId)
      }
      throw asRuntimeError(error, 'codex_session_start_failed')
    }
  }

  async startTurn(
    owner: CodexOwner,
    documentId: string,
    text: string,
  ): Promise<{ readonly turnId: string }> {
    const record = this.#availableDocument(owner, documentId)
    if (!boundedNonempty(text, MAX_TURN_TEXT_BYTES)) {
      throw new ShellCodexRuntimeError('codex_invalid_turn')
    }
    if (record.activeTurn) throw new ShellCodexRuntimeError('codex_turn_in_progress')
    const turn: ActiveTurn = {
      turnId: undefined,
      text: '',
      cancelRequested: false,
      terminal: false,
    }
    // Reserve synchronously: auth refresh is asynchronous and must not admit a second turn.
    record.activeTurn = turn
    try {
      let status: { readonly loggedIn: boolean }
      try {
        status = await this.#authClient.getValidAccountStatus()
      } catch {
        throw new ShellCodexRuntimeError('auth_required')
      }
      if (!status.loggedIn) throw new ShellCodexRuntimeError('auth_required')
      if (record.state !== 'ready' || record.quiesced || record.activeTurn !== turn) {
        throw new ShellCodexRuntimeError('codex_session_closed')
      }
      if (!record.client || !record.threadId) {
        throw new ShellCodexRuntimeError('codex_document_unavailable')
      }
      const response = await record.client.startTurn(record.threadId, text)
      if (record.state !== 'ready') {
        try {
          await record.client.interruptTurn(record.threadId, response.turn.id)
        } catch {
          this.#emitDiagnostic('codex_turn_interrupt_failed')
        }
        throw new ShellCodexRuntimeError('codex_session_closed')
      }
      if (record.activeTurn !== turn) {
        if (turn.terminal) return { turnId: response.turn.id }
        throw new ShellCodexRuntimeError('codex_session_closed')
      }
      if (turn.turnId !== undefined && turn.turnId !== response.turn.id) {
        throw new ShellCodexRuntimeError('codex_turn_protocol_error')
      }
      turn.turnId = response.turn.id
      if (turn.cancelRequested) await this.#interrupt(record, turn)
      return { turnId: response.turn.id }
    } catch (error) {
      if (record.activeTurn === turn) record.activeTurn = undefined
      const failure = asRuntimeError(error, 'codex_turn_start_failed')
      if (failure.code !== 'codex_session_closed') {
        this.#emitError(
          record,
          failure.code,
          failure.code === 'auth_required' ? 'Sign in to WisWork.' : 'Turn failed.',
        )
      }
      throw failure
    }
  }

  async cancelTurn(owner: CodexOwner, documentId: string): Promise<void> {
    const record = this.#availableDocument(owner, documentId)
    const turn = record.activeTurn
    if (!turn) return
    turn.cancelRequested = true
    if (turn.turnId) await this.#interrupt(record, turn)
  }

  async quiesceDocument(documentId: string): Promise<void> {
    let record = this.#documents.get(documentId)
    if (!record) return
    record.quiesced = true
    if (record.state === 'starting') {
      await record.startPromise.catch(() => undefined)
      record = this.#documents.get(documentId)
      if (!record || record.state !== 'ready') return
      record.quiesced = true
    }
    if (record.state !== 'ready') return
    const turn = record.activeTurn
    if (!turn) return
    turn.cancelRequested = true
    if (turn.turnId) await this.#interrupt(record, turn)
  }

  resumeDocument(documentId: string): void {
    const record = this.#documents.get(documentId)
    if (record?.state === 'ready') record.quiesced = false
  }

  async quiesceDocuments(): Promise<void> {
    const settled = await Promise.allSettled(
      [...this.#documents.keys()].map((id) => this.quiesceDocument(id)),
    )
    const failure = settled.find(
      (result): result is PromiseRejectedResult => result.status === 'rejected',
    )
    if (failure) throw failure.reason
  }

  resumeDocuments(): void {
    for (const record of this.#documents.values()) {
      if (record.state === 'ready') record.quiesced = false
    }
  }

  acquireDocumentCloseGate(documentId: string): () => void {
    if (this.#globalCloseGate || this.#documentCloseGates.has(documentId)) {
      throw new ShellCodexRuntimeError('codex_runtime_busy')
    }
    this.#documentCloseGates.add(documentId)
    let released = false
    return () => {
      if (released) return
      released = true
      this.#documentCloseGates.delete(documentId)
    }
  }

  acquireGlobalCloseGate(): () => void {
    if (this.#globalCloseGate || this.#documentCloseGates.size > 0) {
      throw new ShellCodexRuntimeError('codex_runtime_busy')
    }
    this.#globalCloseGate = true
    let released = false
    return () => {
      if (released) return
      released = true
      this.#globalCloseGate = false
    }
  }

  closeDocument(documentId: string): Promise<void> {
    const record = this.#documents.get(documentId)
    if (!record) return Promise.resolve()
    return this.#closeRecord(record)
  }

  async logout(): Promise<void> {
    if (this.#logoutPromise) return this.#logoutPromise
    this.#acceptingDocuments = false
    this.#logoutPromise = this.#closeAllDocuments().finally(() => {
      this.#logoutPromise = undefined
      if (!this.#closed) this.#acceptingDocuments = true
    })
    return this.#logoutPromise
  }

  shutdown(): Promise<void> {
    if (this.#shutdownPromise) return this.#shutdownPromise
    this.#closed = true
    this.#acceptingDocuments = false
    this.#shutdownPromise = this.#closeAllDocuments()
    return this.#shutdownPromise
  }

  async #startDocument(record: DocumentRecord): Promise<void> {
    try {
      let status: { readonly loggedIn: boolean }
      try {
        status = await this.#authClient.getValidAccountStatus()
      } catch {
        throw new ShellCodexRuntimeError('auth_required')
      }
      if (!status.loggedIn) throw new ShellCodexRuntimeError('auth_required')
      this.#assertStarting(record)

      try {
        record.bridge = await this.#startResponsesBridge({
          fetchWithAuth: createWisUsageBridgeFetch(this.#authClient),
        })
      } catch {
        throw new ShellCodexRuntimeError('codex_bridge_start_failed')
      }
      this.#assertStarting(record)

      try {
        record.mcp = await this.#startDocumentMcpServer()
      } catch {
        throw new ShellCodexRuntimeError('codex_mcp_start_failed')
      }
      this.#assertStarting(record)
      try {
        record.mcpSession = record.mcp.register(this.#instrumentRegistration(record))
      } catch {
        throw new ShellCodexRuntimeError('codex_mcp_registration_failed')
      }
      this.#assertStarting(record)

      record.manager = this.#createProcessManager({
        executablePath: this.#executablePath!,
        bridge: record.bridge,
        mcp: { url: record.mcpSession.url, secret: record.mcpSession.secret },
        developerInstructions: WISWORK_CODEX_DEVELOPER_INSTRUCTIONS,
        diagnostics: (diagnostic) => this.#emitDiagnostic(diagnostic.code),
      })
      void record.manager.crashed.then(
        () => this.#handleCrash(record),
        () => {
          this.#emitDiagnostic('codex_crash_monitor_failed')
          return this.#handleCrash(record)
        },
      )
      try {
        record.client = await record.manager.start()
      } catch {
        throw new ShellCodexRuntimeError('codex_process_start_failed')
      }
      this.#assertStarting(record)
      try {
        await record.client.initialize()
      } catch {
        throw new ShellCodexRuntimeError('codex_initialize_failed')
      }
      this.#assertStarting(record)
      record.unsubscribe = record.client.onNotification((notification) =>
        this.#handleNotification(record, notification),
      )
      try {
        const thread = await record.client.startThread()
        record.threadId = thread.thread.id
      } catch {
        throw new ShellCodexRuntimeError('codex_thread_start_failed')
      }
      this.#assertStarting(record)
      record.state = 'ready'
    } catch (error) {
      await this.#teardownResources(record)
      throw error
    }
  }

  #assertStarting(record: DocumentRecord): void {
    if (
      record.state !== 'starting' ||
      this.#closed ||
      record.owner.isDestroyed() ||
      this.#globalCloseGate ||
      this.#documentCloseGates.has(record.documentId)
    ) {
      throw new ShellCodexRuntimeError('codex_session_closed')
    }
  }

  #availableDocument(owner: CodexOwner, documentId: string): DocumentRecord {
    const record = this.#documents.get(documentId)
    if (
      !record ||
      record.owner !== owner ||
      owner.isDestroyed() ||
      record.state !== 'ready' ||
      record.quiesced
    ) {
      throw new ShellCodexRuntimeError('codex_document_unavailable')
    }
    return record
  }

  #instrumentRegistration(record: DocumentRecord): ToolSessionRegistration {
    const registration = record.registration
    const reserveToolCall = (call: AgentToolCall, phase: 'approval' | 'execute'): void => {
      const activeTurn = record.activeTurn
      if (!activeTurn) throw new ShellCodexRuntimeError('codex_turn_unavailable')
      let encoded: string
      try {
        encoded = JSON.stringify([call.id, call.name, call.input])
      } catch {
        throw new ShellCodexRuntimeError('codex_tool_call_limit')
      }
      const fingerprint = createHash('sha256').update(encoded).digest('hex')
      const budget = activeTurn.toolCall
      if (budget && budget.fingerprint !== fingerprint) {
        throw new ShellCodexRuntimeError('codex_tool_call_limit')
      }
      const owned = budget ?? { fingerprint, approvalStarted: false, executed: false }
      activeTurn.toolCall = owned
      if (phase === 'approval') {
        if (owned.approvalStarted || owned.executed) {
          throw new ShellCodexRuntimeError('codex_tool_call_limit')
        }
        owned.approvalStarted = true
      } else {
        if (owned.executed) throw new ShellCodexRuntimeError('codex_tool_call_limit')
        owned.executed = true
      }
    }
    const executeWithEvents = async (
      call: AgentToolCall,
      operation: () => ToolExecution | Promise<ToolExecution>,
      snapshotBefore?: string,
    ): Promise<ToolExecution> => {
      const activeTurn = record.activeTurn
      if (!activeTurn) throw new ShellCodexRuntimeError('codex_turn_unavailable')
      this.#emit(record, { type: 'tool-start', call })
      try {
        const execution = await operation()
        if (record.activeTurn === activeTurn) {
          this.#emit(record, {
            type: 'tool-executed',
            call,
            execution,
            ...(snapshotBefore === undefined ? {} : { snapshotBefore }),
          })
        }
        return execution
      } catch (error) {
        if (record.activeTurn === activeTurn) {
          this.#emit(record, {
            type: 'tool-executed',
            call,
            execution: {
              output: 'tool_execution_failed',
              summary: call.name,
              isError: true,
            },
          })
        }
        throw error
      } finally {
        if (record.activeTurn === activeTurn) this.#emit(record, { type: 'turn-end' })
      }
    }
    return {
      ...registration,
      skill: {
        ...registration.skill,
        executeTool: (call, signal) => {
          reserveToolCall(call, 'execute')
          return executeWithEvents(call, () => registration.skill.executeTool(call, signal))
        },
      },
      requestApproval: (call, expectedRevision, signal) => {
        reserveToolCall(call, 'approval')
        return registration.requestApproval(call, expectedRevision, signal)
      },
      ...(registration.executeMutation
        ? {
            executeMutation: (call, guard, signal) => {
              reserveToolCall(call, 'execute')
              return executeWithEvents(
                call,
                () => registration.executeMutation!(call, guard, signal),
                guard.snapshotId,
              )
            },
          }
        : {}),
    }
  }

  async #interrupt(record: DocumentRecord, turn: ActiveTurn): Promise<void> {
    if (!record.client || !record.threadId || !turn.turnId) return
    if (turn.interruptPromise) return turn.interruptPromise
    turn.interruptPromise = record.client
      .interruptTurn(record.threadId, turn.turnId)
      .then(() => undefined)
      .catch(() => {
        this.#emitError(record, 'codex_turn_interrupt_failed', 'Could not stop the turn.')
        throw new ShellCodexRuntimeError('codex_turn_interrupt_failed')
      })
    return turn.interruptPromise
  }

  #handleNotification(record: DocumentRecord, notification: CodexAppServerNotification): void {
    if (record.state !== 'ready' || !record.threadId) return
    const turn = record.activeTurn
    if (!turn) return
    if (notification.method === 'turn/started') {
      if (notification.params.threadId !== record.threadId) return
      const id = notification.params.turn.id
      if (record.finishedTurnIds.has(id)) return
      if (turn.turnId !== undefined && turn.turnId !== id) return
      turn.turnId = id
      if (turn.cancelRequested) void this.#interrupt(record, turn).catch(() => undefined)
      return
    }
    if (notification.method === 'item/agentMessage/delta') {
      if (
        notification.params.threadId !== record.threadId ||
        record.finishedTurnIds.has(notification.params.turnId) ||
        (turn.turnId !== undefined && turn.turnId !== notification.params.turnId)
      ) {
        return
      }
      turn.turnId = notification.params.turnId
      if (
        Buffer.byteLength(turn.text) + Buffer.byteLength(notification.params.delta) >
        MAX_TURN_TEXT_BYTES
      ) {
        this.#emitError(record, 'codex_event_limit', 'Turn output was too large.')
        void this.#closeRecord(record)
        return
      }
      turn.text += notification.params.delta
      this.#emit(record, { type: 'text', text: turn.text })
      return
    }
    if (notification.method === 'turn/completed') {
      if (
        notification.params.threadId !== record.threadId ||
        record.finishedTurnIds.has(notification.params.turn.id) ||
        (turn.turnId !== undefined && turn.turnId !== notification.params.turn.id)
      ) {
        return
      }
      const status = safeTurnStatus(notification.params.turn)
      if (!status) {
        this.#emitError(record, 'codex_turn_protocol_error', 'Turn failed.')
        void this.#closeRecord(record)
        return
      }
      record.activeTurn = undefined
      turn.terminal = true
      this.#rememberFinishedTurn(record, notification.params.turn.id)
      if (status === 'failed') {
        this.#emitError(record, 'codex_turn_failed', 'Turn failed.')
        return
      }
      this.#emit(record, {
        type: 'done',
        result: {
          text: turn.text,
          cancelled: status === 'interrupted' || turn.cancelRequested,
          turnLimit: false,
        },
      })
      return
    }
    if (notification.method === 'error' && isRecord(notification.params)) {
      if (
        notification.params.willRetry !== false ||
        notification.params.threadId !== record.threadId ||
        (turn.turnId !== undefined && notification.params.turnId !== turn.turnId)
      ) {
        return
      }
      record.activeTurn = undefined
      turn.terminal = true
      if (typeof notification.params.turnId === 'string') {
        this.#rememberFinishedTurn(record, notification.params.turnId)
      }
      this.#emitError(record, 'codex_turn_failed', 'Turn failed.')
    }
  }

  #rememberFinishedTurn(record: DocumentRecord, turnId: string): void {
    record.finishedTurnIds.add(turnId)
    while (record.finishedTurnIds.size > FINISHED_TURN_IDS_LIMIT) {
      record.finishedTurnIds.delete(record.finishedTurnIds.values().next().value!)
    }
  }

  async #handleCrash(record: DocumentRecord): Promise<void> {
    if (record.state === 'closing' || record.state === 'closed') return
    try {
      this.#onProcessCrash?.(record.documentId)
    } catch {
      this.#emitDiagnostic('codex_crash_cleanup_failed')
    }
    this.#emitError(record, 'codex_process_exited', 'Codex stopped.')
    await this.#closeRecord(record)
  }

  #closeRecord(record: DocumentRecord): Promise<void> {
    if (record.closePromise) return record.closePromise
    record.state = 'closing'
    record.closePromise = (async () => {
      try {
        await record.startPromise.catch(() => undefined)
        if (record.activeTurn?.turnId) {
          await this.#interrupt(record, record.activeTurn).catch(() => undefined)
        }
        record.activeTurn = undefined
        await this.#teardownResources(record)
      } finally {
        record.state = 'closed'
        if (this.#documents.get(record.documentId) === record) {
          this.#documents.delete(record.documentId)
        }
      }
    })()
    return record.closePromise
  }

  async #teardownResources(record: DocumentRecord): Promise<void> {
    record.unsubscribe?.()
    record.unsubscribe = undefined
    const manager = record.manager
    record.manager = undefined
    if (manager) {
      try {
        await manager.stop()
      } catch {
        this.#emitDiagnostic('codex_process_stop_failed')
      }
    }
    const mcpSession = record.mcpSession
    record.mcpSession = undefined
    if (mcpSession) {
      try {
        mcpSession.close()
      } catch {
        this.#emitDiagnostic('codex_mcp_session_close_failed')
      }
    }
    const mcp = record.mcp
    record.mcp = undefined
    if (mcp) {
      try {
        await mcp.close()
      } catch {
        this.#emitDiagnostic('codex_mcp_close_failed')
      }
    }
    const bridge = record.bridge
    record.bridge = undefined
    if (bridge) {
      try {
        await bridge.close()
      } catch {
        this.#emitDiagnostic('codex_bridge_close_failed')
      }
    }
    record.client = undefined
    record.threadId = undefined
  }

  async #closeAllDocuments(): Promise<void> {
    await Promise.all([...this.#documents.values()].map((record) => this.#closeRecord(record)))
  }

  #emit(record: DocumentRecord, event: AgentEvent<unknown>): void {
    if (record.owner.isDestroyed()) return
    try {
      record.owner.send(CODEX_RUNTIME_CHANNELS.event, { documentId: record.documentId, event })
    } catch {
      void this.#closeRecord(record)
    }
  }

  #emitError(record: DocumentRecord, code: string, message: string): void {
    this.#emit(record, { type: 'error', code, message })
  }

  #emitDiagnostic(code: string): void {
    try {
      this.#diagnostics?.({ code })
    } catch {
      // Diagnostics cannot affect lifecycle.
    }
  }
}

export interface CodexToolLifecycle {
  acquireDocumentCloseGate?(documentId: string): () => void
  acquireGlobalCloseGate?(): () => void
  quiesceDocument(documentId: string): Promise<void>
  resumeDocument(documentId: string): void
  quiesceSessions(): Promise<void>
  resumeSessions(): void
  closeDocument(documentId: string): void
  closeSessions(): void
}

export interface CodexClosePreparation {
  commit(): Promise<void>
  rollback(): void
  finalize(): void
}

export async function runCodexPreparedClose(
  preparation: Pick<CodexClosePreparation, 'commit' | 'rollback'>,
  checks: readonly (() => boolean | Promise<boolean>)[],
): Promise<boolean> {
  let committed = false
  try {
    for (const check of checks) {
      if ((await check()) !== true) return false
    }
    await preparation.commit()
    committed = true
    return true
  } finally {
    if (!committed) preparation.rollback()
  }
}

export async function prepareCodexDocumentClose(
  runtime: ShellCodexRuntime | null,
  tools: CodexToolLifecycle | null,
  documentId: string,
): Promise<CodexClosePreparation> {
  let releaseRuntime: (() => void) | undefined
  let releaseTools: (() => void) | undefined
  let runtimeOwned = false
  let toolsOwned = false
  let runtimeQuiesceStarted = false
  let toolsQuiesceStarted = false
  try {
    if (runtime) {
      releaseRuntime = runtime.acquireDocumentCloseGate(documentId)
      runtimeOwned = true
    }
    if (tools) {
      releaseTools = tools.acquireDocumentCloseGate?.(documentId)
      toolsOwned = true
    }
    if (runtime) {
      runtimeQuiesceStarted = true
      await runtime.quiesceDocument(documentId)
    }
    if (tools) {
      toolsQuiesceStarted = true
      await tools.quiesceDocument(documentId)
    }
  } catch (error) {
    if (toolsOwned && toolsQuiesceStarted) tools?.resumeDocument(documentId)
    if (runtimeOwned && runtimeQuiesceStarted) runtime?.resumeDocument(documentId)
    if (toolsOwned) releaseTools?.()
    if (runtimeOwned) releaseRuntime?.()
    throw error
  }
  let settled = false
  let committed = false
  return {
    async commit() {
      if (settled) return
      settled = true
      committed = true
      // Keep both gates as tombstones until the renderer/document identity is destroyed.
      if (toolsOwned) tools?.closeDocument(documentId)
      if (runtimeOwned) await runtime?.closeDocument(documentId)
    },
    rollback() {
      if (settled) return
      settled = true
      if (toolsOwned && toolsQuiesceStarted) tools?.resumeDocument(documentId)
      if (runtimeOwned && runtimeQuiesceStarted) runtime?.resumeDocument(documentId)
      if (toolsOwned) releaseTools?.()
      if (runtimeOwned) releaseRuntime?.()
    },
    finalize() {
      if (!committed) return
      committed = false
      if (toolsOwned) releaseTools?.()
      if (runtimeOwned) releaseRuntime?.()
    },
  }
}

export async function prepareCodexDocumentsClose(
  runtime: ShellCodexRuntime | null,
  tools: CodexToolLifecycle | null,
): Promise<CodexClosePreparation> {
  let releaseRuntime: (() => void) | undefined
  let releaseTools: (() => void) | undefined
  let runtimeOwned = false
  let toolsOwned = false
  let runtimeQuiesceStarted = false
  let toolsQuiesceStarted = false
  try {
    if (runtime) {
      releaseRuntime = runtime.acquireGlobalCloseGate()
      runtimeOwned = true
    }
    if (tools) {
      releaseTools = tools.acquireGlobalCloseGate?.()
      toolsOwned = true
    }
    if (runtime) {
      runtimeQuiesceStarted = true
      await runtime.quiesceDocuments()
    }
    if (tools) {
      toolsQuiesceStarted = true
      await tools.quiesceSessions()
    }
  } catch (error) {
    if (toolsOwned && toolsQuiesceStarted) tools?.resumeSessions()
    if (runtimeOwned && runtimeQuiesceStarted) runtime?.resumeDocuments()
    if (toolsOwned) releaseTools?.()
    if (runtimeOwned) releaseRuntime?.()
    throw error
  }
  let settled = false
  let committed = false
  return {
    async commit() {
      if (settled) return
      settled = true
      committed = true
      // The application/window is committed to close; retain the global tombstones.
      if (toolsOwned) tools?.closeSessions()
      if (runtimeOwned) await runtime?.logout()
    },
    rollback() {
      if (settled) return
      settled = true
      if (toolsOwned && toolsQuiesceStarted) tools?.resumeSessions()
      if (runtimeOwned && runtimeQuiesceStarted) runtime?.resumeDocuments()
      if (toolsOwned) releaseTools?.()
      if (runtimeOwned) releaseRuntime?.()
    },
    finalize() {
      if (!committed) return
      committed = false
      if (toolsOwned) releaseTools?.()
      if (runtimeOwned) releaseRuntime?.()
    },
  }
}

export async function logoutWithCodexClose(
  runtime: ShellCodexRuntime | null,
  tools: CodexToolLifecycle | null,
  logoutAuthority: () => Promise<void>,
): Promise<void> {
  const preparation = await prepareCodexDocumentsClose(runtime, tools)
  await preparation.commit()
  try {
    await logoutAuthority()
  } catch {
    throw new ShellCodexRuntimeError('auth_logout_failed')
  } finally {
    // The authority attempt has settled; resource cleanup stays committed, while the lease can
    // be released so a visible failed logout can be retried instead of permanently wedging IPC.
    preparation.finalize()
  }
}

export interface RegisterCodexRuntimeIpcOptions {
  readonly ipcMain: IpcMainLike
  readonly runtime: ShellCodexRuntime
  readonly documentIdForOwner: (owner: CodexOwner) => string | null
}

export function registerCodexRuntimeIpc(options: RegisterCodexRuntimeIpcOptions): void {
  options.ipcMain.handle(CODEX_RUNTIME_CHANNELS.status, (event, ...args): CodexRuntimeStatus => {
    if (args.length !== 0) throw new ShellCodexRuntimeError('invalid_codex_runtime_request')
    return {
      runtime: options.runtime.runtimeKind,
      documentId: options.documentIdForOwner(event.sender),
    }
  })
  options.ipcMain.handle(CODEX_RUNTIME_CHANNELS.startTurn, async (event, ...args) => {
    if (
      args.length !== 1 ||
      !isRecord(args[0]) ||
      !hasOnlyKeys(args[0], ['documentId', 'text']) ||
      !boundedNonempty(args[0].documentId, MAX_DOCUMENT_ID_BYTES) ||
      !boundedNonempty(args[0].text, MAX_TURN_TEXT_BYTES)
    ) {
      throw new ShellCodexRuntimeError('invalid_codex_runtime_request')
    }
    const request = args[0] as unknown as CodexRuntimeStartRequest
    if (options.documentIdForOwner(event.sender) !== request.documentId) {
      throw new ShellCodexRuntimeError('codex_document_unavailable')
    }
    return options.runtime.startTurn(event.sender, request.documentId, request.text)
  })
  options.ipcMain.handle(CODEX_RUNTIME_CHANNELS.cancelTurn, async (event, ...args) => {
    if (args.length !== 1 || !boundedNonempty(args[0], MAX_DOCUMENT_ID_BYTES)) {
      throw new ShellCodexRuntimeError('invalid_codex_runtime_request')
    }
    if (options.documentIdForOwner(event.sender) !== args[0]) {
      throw new ShellCodexRuntimeError('codex_document_unavailable')
    }
    await options.runtime.cancelTurn(event.sender, args[0])
  })
}

export interface CodexBeforeQuitOptions {
  readonly shutdown: () => Promise<void>
  readonly quit: () => void
  readonly diagnostics?: (diagnostic: ShellCodexRuntimeDiagnostic) => void
}

/** Electron before-quit adapter: finish the owned child chain before retrying app.quit(). */
export function createCodexBeforeQuitHandler(
  options: CodexBeforeQuitOptions,
): (event: { preventDefault(): void }) => void {
  let ready = false
  let shutdownPromise: Promise<void> | undefined
  return (event) => {
    if (ready) return
    event.preventDefault()
    if (shutdownPromise) return
    try {
      shutdownPromise = Promise.resolve(options.shutdown())
    } catch {
      shutdownPromise = Promise.reject(new Error('codex_shutdown_failed'))
    }
    void shutdownPromise
      .catch(() => {
        try {
          options.diagnostics?.({ code: 'codex_shutdown_failed' })
        } catch {
          // Diagnostics cannot prevent application shutdown.
        }
      })
      .finally(() => {
        ready = true
        options.quit()
      })
  }
}
