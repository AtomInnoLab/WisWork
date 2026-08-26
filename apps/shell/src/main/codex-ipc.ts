import { randomBytes } from 'node:crypto'
import type { AgentSkill, AgentToolCall, AgentToolDef, ToolExecution } from '@wiswork/agent-core'
import type { ToolSessionRegistration } from '@wiswork/codex-bridge'
import {
  CODEX_TOOL_CHANNELS,
  type CodexToolRegistrationRequest,
  type CodexToolRequest,
  type CodexToolResponse,
} from '../shared/codex-api'

const MAX_ID_BYTES = 256
const MAX_PROMPT_BYTES = 256_000
const MAX_TOOLS = 256
const MAX_SCHEMA_BYTES = 1_000_000
const MAX_CALL_BYTES = 1_000_000
const MAX_OUTPUT_BYTES = 1_000_000
const MAX_SUMMARY_BYTES = 4_096
const MAX_REVISION_BYTES = 4_096
const MAX_REGISTRATION_BYTES = 4_000_000
const DEFAULT_TIMEOUT_MS = 120_000

interface IpcMainLike {
  handle(
    channel: string,
    handler: (event: { sender: WebContentsLike }, ...args: unknown[]) => unknown,
  ): void
}

interface WebContentsLike {
  readonly id?: number
  isDestroyed(): boolean
  send(channel: string, payload: unknown): void
  once(event: 'destroyed', listener: () => void): this
  off(event: 'destroyed', listener: () => void): this
}

export interface CodexToolIpcDiagnostic {
  readonly code: string
}

export interface RegisteredCodexToolSession {
  close(): void | Promise<void>
}

export interface CodexToolIpcOptions {
  readonly ipcMain: IpcMainLike
  /** Authoritative live-document ownership check; renderer-provided IDs are never trusted alone. */
  readonly ownsDocument: (owner: WebContentsLike, documentId: string) => boolean
  readonly onRegister: (input: {
    readonly documentId: string
    readonly owner: WebContentsLike
    readonly registration: ToolSessionRegistration
  }) => RegisteredCodexToolSession | Promise<RegisteredCodexToolSession>
  readonly requestTimeoutMs?: number
  readonly randomId?: () => string
  readonly diagnostics?: (diagnostic: CodexToolIpcDiagnostic) => void
}

export interface CodexToolIpcController {
  acquireDocumentCloseGate(documentId: string): () => void
  acquireGlobalCloseGate(): () => void
  quiesceDocument(documentId: string): Promise<void>
  resumeDocument(documentId: string): void
  quiesceSessions(): Promise<void>
  resumeSessions(): void
  closeDocument(documentId: string): void
  closeSessions(): void
  close(): void
}

interface RemoteSession {
  readonly documentId: string
  readonly owner: WebContentsLike
  readonly registration: ToolSessionRegistration
  readonly closeRegistration: RegisteredCodexToolSession
  readonly onDestroyed: () => void
  closed: boolean
  quiesced: boolean
}

type ResponseType = Extract<CodexToolResponse, { ok: true }>['type']
type WithoutTransport<T> = T extends unknown ? Omit<T, 'requestId' | 'documentId'> : never
type OutboundToolRequest = WithoutTransport<CodexToolRequest>

interface PendingRequest {
  readonly owner: WebContentsLike
  readonly documentId: string
  readonly type: ResponseType
  readonly resolve: (response: Extract<CodexToolResponse, { ok: true }>) => void
  readonly reject: (error: Error) => void
  readonly timer: ReturnType<typeof setTimeout>
  readonly signal?: AbortSignal
  readonly onAbort?: () => void
  readonly mutation: boolean
}

interface PendingRegistration {
  readonly done: Promise<void>
  finish(): void
}

function isRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

function hasOnlyKeys(value: Record<string, unknown>, allowed: readonly string[]): boolean {
  const set = new Set(allowed)
  return Object.keys(value).every((key) => set.has(key))
}

function boundedString(value: unknown, maximum: number): value is string {
  return typeof value === 'string' && value !== '' && Buffer.byteLength(value) <= maximum
}

function boundedJson(value: unknown, maximum: number): boolean {
  try {
    const serialized = JSON.stringify(value)
    return serialized !== undefined && Buffer.byteLength(serialized) <= maximum
  } catch {
    return false
  }
}

function validTool(tool: unknown): tool is AgentToolDef {
  return (
    isRecord(tool) &&
    hasOnlyKeys(tool, ['name', 'description', 'inputSchema']) &&
    boundedString(tool.name, MAX_ID_BYTES) &&
    boundedString(tool.description, MAX_PROMPT_BYTES) &&
    isRecord(tool.inputSchema) &&
    boundedJson(tool.inputSchema, MAX_SCHEMA_BYTES)
  )
}

function validateRegistration(value: unknown): CodexToolRegistrationRequest {
  if (
    !isRecord(value) ||
    !hasOnlyKeys(value, ['documentId', 'skill', 'policy']) ||
    !boundedString(value.documentId, MAX_ID_BYTES) ||
    !isRecord(value.skill) ||
    !hasOnlyKeys(value.skill, ['id', 'systemPrompt', 'tools']) ||
    !boundedString(value.skill.id, MAX_ID_BYTES) ||
    typeof value.skill.systemPrompt !== 'string' ||
    Buffer.byteLength(value.skill.systemPrompt) > MAX_PROMPT_BYTES ||
    !Array.isArray(value.skill.tools) ||
    value.skill.tools.length === 0 ||
    value.skill.tools.length > MAX_TOOLS ||
    !value.skill.tools.every(validTool) ||
    !isRecord(value.policy) ||
    !boundedJson(value, MAX_REGISTRATION_BYTES)
  ) {
    throw new Error('invalid_codex_tool_registration')
  }
  const names = new Set(value.skill.tools.map((tool) => tool.name))
  const policies = Object.entries(value.policy)
  if (
    names.size !== value.skill.tools.length ||
    policies.length !== names.size ||
    policies.some(
      ([name, policy]) => !names.has(name) || (policy !== 'read' && policy !== 'mutate'),
    )
  ) {
    throw new Error('invalid_codex_tool_registration')
  }
  return value as unknown as CodexToolRegistrationRequest
}

function validCall(value: unknown): value is AgentToolCall {
  return (
    isRecord(value) &&
    hasOnlyKeys(value, ['id', 'name', 'input', 'inputError', 'truncated']) &&
    boundedString(value.id, MAX_ID_BYTES) &&
    boundedString(value.name, MAX_ID_BYTES) &&
    isRecord(value.input) &&
    boundedJson(value.input, MAX_CALL_BYTES) &&
    (value.inputError === undefined || typeof value.inputError === 'string') &&
    (value.truncated === undefined || typeof value.truncated === 'boolean')
  )
}

function validExecution(value: unknown): value is ToolExecution {
  return (
    isRecord(value) &&
    hasOnlyKeys(value, ['output', 'summary', 'isError', 'mutated', 'display']) &&
    typeof value.output === 'string' &&
    Buffer.byteLength(value.output) <= MAX_OUTPUT_BYTES &&
    typeof value.summary === 'string' &&
    Buffer.byteLength(value.summary) <= MAX_SUMMARY_BYTES &&
    (value.isError === undefined || typeof value.isError === 'boolean') &&
    (value.mutated === undefined || typeof value.mutated === 'boolean') &&
    (value.display === undefined || boundedJson(value.display, MAX_CALL_BYTES))
  )
}

function validSuccessResponse(
  value: unknown,
  type: ResponseType,
): value is Extract<CodexToolResponse, { ok: true }> {
  if (!isRecord(value) || value.ok !== true || value.type !== type) return false
  if (type === 'revision') {
    return (
      hasOnlyKeys(value, ['requestId', 'ok', 'type', 'revision']) &&
      boundedString(value.revision, MAX_REVISION_BYTES)
    )
  }
  if (type === 'approval') {
    return (
      hasOnlyKeys(value, ['requestId', 'ok', 'type', 'approved']) &&
      typeof value.approved === 'boolean'
    )
  }
  if (type === 'snapshot') {
    return (
      hasOnlyKeys(value, ['requestId', 'ok', 'type', 'snapshotId']) &&
      boundedString(value.snapshotId, MAX_REVISION_BYTES)
    )
  }
  if (type === 'execution') {
    return (
      hasOnlyKeys(value, ['requestId', 'ok', 'type', 'execution']) &&
      validExecution(value.execution)
    )
  }
  return hasOnlyKeys(value, ['requestId', 'ok', 'type'])
}

const FAILURE_CODES = new Set([
  'tool_cancelled',
  'tool_denied',
  'document_changed',
  'tool_session_closed',
  'tool_execution_failed',
  'snapshot_failed',
  'validation_failed',
])

export function registerCodexToolIpc(options: CodexToolIpcOptions): CodexToolIpcController {
  const timeoutMs = options.requestTimeoutMs ?? DEFAULT_TIMEOUT_MS
  if (!Number.isSafeInteger(timeoutMs) || timeoutMs <= 0)
    throw new TypeError('invalid_tool_ipc_timeout')
  const makeId = options.randomId ?? (() => randomBytes(16).toString('base64url'))
  const sessions = new Map<WebContentsLike, Map<string, RemoteSession>>()
  const activeDocuments = new Map<string, RemoteSession>()
  const pendingRegistrations = new Map<string, PendingRegistration>()
  const documentCloseGates = new Set<string>()
  let globalCloseGate = false
  const poisonedDocuments = new Set<string>()
  const pending = new Map<string, PendingRequest>()
  const drainWaiters = new Map<string, Set<() => void>>()
  let closed = false

  const diagnostic = (code: string): void => {
    try {
      options.diagnostics?.({ code })
    } catch {
      // Diagnostics cannot affect tool execution.
    }
  }

  const finishPending = (requestId: string, error: Error): void => {
    const item = pending.get(requestId)
    if (!item) return
    pending.delete(requestId)
    clearTimeout(item.timer)
    if (item.signal && item.onAbort) item.signal.removeEventListener('abort', item.onAbort)
    item.reject(error)
    if (![...pending.values()].some((candidate) => candidate.documentId === item.documentId)) {
      const waiters = drainWaiters.get(item.documentId)
      drainWaiters.delete(item.documentId)
      for (const resolve of waiters ?? []) resolve()
    }
  }

  const sendCancel = (session: RemoteSession, requestId: string): void => {
    if (session.owner.isDestroyed()) return
    try {
      session.owner.send(CODEX_TOOL_CHANNELS.cancel, {
        requestId,
        documentId: session.documentId,
      })
    } catch {
      // The lifecycle path still closes/rejects the request.
    }
  }

  const poisonDocument = (session: RemoteSession): void => {
    poisonedDocuments.add(session.documentId)
  }

  const closeSession = (session: RemoteSession): void => {
    if (session.closed) return
    session.closed = true
    session.owner.off('destroyed', session.onDestroyed)
    const ownerSessions = sessions.get(session.owner)
    ownerSessions?.delete(session.documentId)
    if (ownerSessions?.size === 0) sessions.delete(session.owner)
    if (activeDocuments.get(session.documentId) === session) {
      activeDocuments.delete(session.documentId)
    }
    for (const [requestId, item] of pending) {
      if (item.owner === session.owner && item.documentId === session.documentId) {
        if (item.mutation && !session.owner.isDestroyed()) poisonDocument(session)
        sendCancel(session, requestId)
        finishPending(requestId, new Error('tool_session_closed'))
      }
    }
    try {
      void Promise.resolve(session.closeRegistration.close()).catch(() =>
        diagnostic('tool_session_close_failed'),
      )
    } catch {
      diagnostic('tool_session_close_failed')
    }
  }

  const currentSession = (
    owner: WebContentsLike,
    documentId: string,
  ): RemoteSession | undefined => {
    const session = activeDocuments.get(documentId)
    if (session?.owner !== owner) return undefined
    return !session?.closed && !owner.isDestroyed() ? session : undefined
  }

  const requestRemote = <T extends Extract<CodexToolResponse, { ok: true }>>(
    session: RemoteSession,
    request: OutboundToolRequest,
    responseType: T['type'],
    signal?: AbortSignal,
  ): Promise<T> => {
    if (!currentSession(session.owner, session.documentId) || session.quiesced) {
      return Promise.reject(new Error('tool_session_closed'))
    }
    let requestId = makeId()
    for (let attempt = 0; attempt < 8 && pending.has(requestId); attempt++) requestId = makeId()
    if (!boundedString(requestId, MAX_ID_BYTES) || pending.has(requestId)) {
      return Promise.reject(new Error('tool_ipc_id_failed'))
    }
    return new Promise<T>((resolve, reject) => {
      const mutation = request.type === 'executeMutation'
      const timer = setTimeout(() => {
        diagnostic('tool_ipc_timeout')
        sendCancel(session, requestId)
        if (mutation) {
          poisonDocument(session)
          closeSession(session)
        } else {
          finishPending(requestId, new Error('tool_ipc_timeout'))
        }
      }, timeoutMs)
      timer.unref()
      const onAbort = (): void => {
        sendCancel(session, requestId)
        // A mutation has entered its guarded commit phase. Keep it serialized until the renderer
        // acknowledges cancellation or reports the actual commit result.
        if (!mutation) finishPending(requestId, new Error('tool_cancelled'))
      }
      pending.set(requestId, {
        owner: session.owner,
        documentId: session.documentId,
        type: responseType,
        resolve: resolve as PendingRequest['resolve'],
        reject,
        timer,
        signal,
        onAbort,
        mutation,
      })
      signal?.addEventListener('abort', onAbort, { once: true })
      if (signal?.aborted) {
        onAbort()
        return
      }
      try {
        session.owner.send(CODEX_TOOL_CHANNELS.request, {
          ...request,
          requestId,
          documentId: session.documentId,
        })
      } catch {
        finishPending(requestId, new Error('tool_session_closed'))
      }
    })
  }

  options.ipcMain.handle(CODEX_TOOL_CHANNELS.register, async (event, ...args) => {
    if (closed || event.sender.isDestroyed() || args.length !== 1) {
      throw new Error('invalid_codex_tool_registration')
    }
    const value = validateRegistration(args[0])
    let ownsDocument = false
    try {
      ownsDocument = options.ownsDocument(event.sender, value.documentId) === true
    } catch {
      // Ownership checks fail closed with a stable code below.
    }
    if (!ownsDocument) throw new Error('untrusted_codex_tool_registration')
    if (globalCloseGate || documentCloseGates.has(value.documentId)) {
      throw new Error('codex_tool_registration_busy')
    }
    if (pendingRegistrations.has(value.documentId)) {
      throw new Error('codex_tool_session_exists')
    }
    if (activeDocuments.has(value.documentId)) throw new Error('codex_tool_session_exists')
    if (poisonedDocuments.has(value.documentId)) {
      throw new Error('codex_tool_session_poisoned')
    }

    const sessionRef: { current?: RemoteSession } = {}
    const requireSession = (): RemoteSession => {
      if (!sessionRef.current) throw new Error('tool_session_closed')
      return sessionRef.current
    }
    const isOpen = (): boolean =>
      sessionRef.current !== undefined &&
      currentSession(event.sender, value.documentId) === sessionRef.current
    const skill: AgentSkill = {
      id: value.skill.id,
      systemPrompt: value.skill.systemPrompt,
      tools: [...value.skill.tools],
      executeTool: (call, signal) => {
        if (!validCall(call)) return Promise.reject(new Error('invalid_tool_call'))
        return requestRemote<Extract<CodexToolResponse, { ok: true; type: 'execution' }>>(
          requireSession(),
          { type: 'execute', call },
          'execution',
          signal,
        ).then((response) => response.execution)
      },
    }
    const registration: ToolSessionRegistration = {
      skill,
      policy: value.policy,
      isOpen,
      getRevision: (signal) =>
        requestRemote<Extract<CodexToolResponse, { ok: true; type: 'revision' }>>(
          requireSession(),
          { type: 'revision' },
          'revision',
          signal,
        ).then((response) => response.revision),
      requestApproval: (call, expectedRevision, signal) =>
        requestRemote<Extract<CodexToolResponse, { ok: true; type: 'approval' }>>(
          requireSession(),
          { type: 'approval', call, expectedRevision },
          'approval',
          signal,
        ).then((response) => response.approved),
      captureSnapshot: (call, expectedRevision, signal) =>
        requestRemote<Extract<CodexToolResponse, { ok: true; type: 'snapshot' }>>(
          requireSession(),
          { type: 'snapshot', call, expectedRevision },
          'snapshot',
          signal,
        ).then((response) => response.snapshotId),
      executeMutation: (call, guard, signal) =>
        requestRemote<Extract<CodexToolResponse, { ok: true; type: 'execution' }>>(
          requireSession(),
          { type: 'executeMutation', call, guard },
          'execution',
          signal,
        ).then((response) => response.execution),
      // executeMutation is a renderer transaction that includes validation before its ack.
      validateMutation: async () => undefined,
    }
    let finishPendingRegistration!: () => void
    const pendingRegistration: PendingRegistration = {
      done: new Promise<void>((resolve) => {
        finishPendingRegistration = resolve
      }),
      finish: () => finishPendingRegistration(),
    }
    pendingRegistrations.set(value.documentId, pendingRegistration)
    let closeRegistration: RegisteredCodexToolSession | undefined
    try {
      try {
        closeRegistration = await options.onRegister({
          documentId: value.documentId,
          owner: event.sender,
          registration,
        })
        if (!closeRegistration || typeof closeRegistration.close !== 'function') throw new Error()
      } catch {
        throw new Error('codex_tool_registration_failed')
      }
      let stillOwnsDocument = false
      try {
        stillOwnsDocument = options.ownsDocument(event.sender, value.documentId) === true
      } catch {
        // Ownership is rechecked after asynchronous startup and fails closed below.
      }
      if (
        closed ||
        globalCloseGate ||
        documentCloseGates.has(value.documentId) ||
        event.sender.isDestroyed() ||
        !stillOwnsDocument
      ) {
        try {
          await closeRegistration.close()
        } catch {
          diagnostic('tool_session_close_failed')
        }
        throw new Error('codex_tool_registration_failed')
      }
      const onDestroyed = (): void => {
        if (sessionRef.current) closeSession(sessionRef.current)
      }
      const session: RemoteSession = {
        documentId: value.documentId,
        owner: event.sender,
        registration,
        closeRegistration,
        onDestroyed,
        closed: false,
        quiesced: false,
      }
      sessionRef.current = session
      let ownerSessions = sessions.get(event.sender)
      if (!ownerSessions) {
        ownerSessions = new Map()
        sessions.set(event.sender, ownerSessions)
      }
      ownerSessions.set(value.documentId, session)
      activeDocuments.set(value.documentId, session)
      event.sender.once('destroyed', onDestroyed)
      return { registered: true as const }
    } finally {
      if (pendingRegistrations.get(value.documentId) === pendingRegistration) {
        pendingRegistrations.delete(value.documentId)
      }
      pendingRegistration.finish()
    }
  })

  options.ipcMain.handle(CODEX_TOOL_CHANNELS.unregister, async (event, ...args) => {
    if (args.length !== 1 || !boundedString(args[0], MAX_ID_BYTES)) {
      throw new Error('invalid_codex_tool_unregister')
    }
    const session = currentSession(event.sender, args[0])
    if (!session) throw new Error('untrusted_codex_tool_unregister')
    closeSession(session)
  })

  options.ipcMain.handle(CODEX_TOOL_CHANNELS.response, async (event, ...args) => {
    if (
      args.length !== 1 ||
      !isRecord(args[0]) ||
      !boundedString(args[0].requestId, MAX_ID_BYTES)
    ) {
      throw new Error('invalid_codex_tool_response')
    }
    const value = args[0]
    const requestId = value.requestId as string
    const item = pending.get(requestId)
    if (!item) return false
    if (item.owner !== event.sender) throw new Error('untrusted_codex_tool_response')
    pending.delete(requestId)
    clearTimeout(item.timer)
    if (item.signal && item.onAbort) item.signal.removeEventListener('abort', item.onAbort)
    if (![...pending.values()].some((candidate) => candidate.documentId === item.documentId)) {
      const waiters = drainWaiters.get(item.documentId)
      drainWaiters.delete(item.documentId)
      for (const resolve of waiters ?? []) resolve()
    }
    if (value.ok === false) {
      if (
        !hasOnlyKeys(value, ['requestId', 'ok', 'code']) ||
        typeof value.code !== 'string' ||
        !FAILURE_CODES.has(value.code)
      ) {
        item.reject(new Error('invalid_tool_ipc_response'))
      } else item.reject(new Error(value.code))
      return true
    }
    if (!validSuccessResponse(value, item.type)) {
      item.reject(new Error('invalid_tool_ipc_response'))
      return true
    }
    item.resolve(value)
    return true
  })

  return {
    acquireDocumentCloseGate(documentId) {
      if (globalCloseGate || documentCloseGates.has(documentId)) {
        throw new Error('codex_tool_registration_busy')
      }
      documentCloseGates.add(documentId)
      let released = false
      return () => {
        if (released) return
        released = true
        documentCloseGates.delete(documentId)
      }
    },
    acquireGlobalCloseGate() {
      if (globalCloseGate || documentCloseGates.size > 0) {
        throw new Error('codex_tool_registration_busy')
      }
      globalCloseGate = true
      let released = false
      return () => {
        if (released) return
        released = true
        globalCloseGate = false
      }
    },
    async quiesceDocument(documentId) {
      await pendingRegistrations.get(documentId)?.done
      const session = activeDocuments.get(documentId)
      if (!session || session.closed) return
      session.quiesced = true
      const matching = [...pending.entries()].filter(([, item]) => item.documentId === documentId)
      if (matching.length === 0) return
      const drained = new Promise<void>((resolve) => {
        let waiters = drainWaiters.get(documentId)
        if (!waiters) {
          waiters = new Set()
          drainWaiters.set(documentId, waiters)
        }
        waiters.add(resolve)
      })
      for (const [requestId, item] of matching) {
        sendCancel(session, requestId)
        if (!item.mutation) finishPending(requestId, new Error('tool_cancelled'))
      }
      await drained
      if (session.closed) throw new Error('tool_session_closed')
    },
    resumeDocument(documentId) {
      const session = activeDocuments.get(documentId)
      if (session && !session.closed) session.quiesced = false
    },
    async quiesceSessions() {
      await Promise.allSettled([...pendingRegistrations.values()].map(({ done }) => done))
      const settled = await Promise.allSettled(
        [...activeDocuments.keys()].map((id) => this.quiesceDocument(id)),
      )
      const failure = settled.find(
        (result): result is PromiseRejectedResult => result.status === 'rejected',
      )
      if (failure) throw failure.reason
    },
    resumeSessions() {
      for (const session of activeDocuments.values()) {
        if (!session.closed) session.quiesced = false
      }
    },
    closeDocument(documentId) {
      const session = activeDocuments.get(documentId)
      if (session) closeSession(session)
    },
    closeSessions() {
      for (const session of [...activeDocuments.values()]) closeSession(session)
    },
    close() {
      if (closed) return
      closed = true
      for (const session of [...activeDocuments.values()]) closeSession(session)
      sessions.clear()
      activeDocuments.clear()
      documentCloseGates.clear()
      globalCloseGate = false
      poisonedDocuments.clear()
    },
  }
}
