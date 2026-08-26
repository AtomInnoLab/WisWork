import { randomBytes as nodeRandomBytes, timingSafeEqual } from 'node:crypto'
import type { AgentSkill, AgentToolCall, AgentToolDef, ToolExecution } from '@wiswork/agent-core'

const TOKEN_BYTES = 32
const MAX_TOOLS = 256
const MAX_NAME_BYTES = 256
const MAX_DESCRIPTION_BYTES = 256_000
const MAX_SCHEMA_BYTES = 1_000_000
const MAX_INPUT_BYTES = 1_000_000
const MAX_OUTPUT_BYTES = 1_000_000
const MAX_SUMMARY_BYTES = 4_096
const MAX_REVISION_BYTES = 4_096
const MAX_SNAPSHOT_ID_BYTES = 4_096
const MAX_REGISTRATION_BYTES = 4_000_000

export type ToolMutability = 'read' | 'mutate'

export interface MutationExecutionGuard {
  readonly expectedRevision: string
  readonly snapshotId: string
}

export interface ToolSessionRegistration {
  readonly skill: AgentSkill
  readonly policy: Readonly<Record<string, ToolMutability>>
  readonly isOpen: () => boolean
  readonly getRevision: (signal?: AbortSignal) => string | Promise<string>
  readonly requestApproval: (
    call: AgentToolCall,
    expectedRevision: string,
    signal?: AbortSignal,
  ) => boolean | Promise<boolean>
  readonly captureSnapshot: (
    call: AgentToolCall,
    expectedRevision: string,
    signal?: AbortSignal,
  ) => string | Promise<string>
  /** Release a captured snapshot when guarded execution never begins. Must be idempotent. */
  readonly discardSnapshot?: (call: AgentToolCall, snapshotId: string) => void | Promise<void>
  readonly executeMutation?: (
    call: AgentToolCall,
    guard: MutationExecutionGuard,
    signal?: AbortSignal,
  ) => ToolExecution | Promise<ToolExecution>
  readonly validateMutation?: (
    call: AgentToolCall,
    execution: ToolExecution,
    guard: MutationExecutionGuard,
    signal?: AbortSignal,
  ) => void | Promise<void>
}

export interface ToolSessionCredentials {
  readonly sessionId: string
  readonly token: string
}

export interface McpToolDefinition extends AgentToolDef {
  readonly annotations: {
    readonly readOnlyHint: boolean
    readonly destructiveHint: boolean
  }
}

export interface DocumentToolRouterOptions {
  readonly randomBytes?: (size: number) => Uint8Array
}

export class ToolRouterError extends Error {
  constructor(readonly code: string) {
    super(code)
    this.name = 'ToolRouterError'
  }
}

interface ToolSession {
  readonly token: Buffer
  readonly registration: ToolSessionRegistration
  readonly tools: ReadonlyMap<string, AgentToolDef>
  readonly pending: Map<string, AbortController>
  closed: boolean
}

const utf8Length = (value: string): number => Buffer.byteLength(value, 'utf8')

function stableExecution(
  output: string,
  summary: string,
  isError: boolean,
  mutated = false,
): ToolExecution {
  return { output, summary, isError, mutated }
}

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

function jsonBytes(value: unknown, code: string): number {
  try {
    const serialized = JSON.stringify(value)
    if (serialized === undefined) throw new Error()
    return utf8Length(serialized)
  } catch {
    throw new ToolRouterError(code)
  }
}

function boundedString(value: unknown, maximum: number, code: string): string {
  if (typeof value !== 'string' || value === '' || utf8Length(value) > maximum) {
    throw new ToolRouterError(code)
  }
  return value
}

function validateRegistration(
  registration: ToolSessionRegistration,
): ReadonlyMap<string, AgentToolDef> {
  if (
    !registration ||
    !registration.skill ||
    typeof registration.skill.id !== 'string' ||
    !Array.isArray(registration.skill.tools) ||
    registration.skill.tools.length === 0 ||
    registration.skill.tools.length > MAX_TOOLS ||
    typeof registration.skill.executeTool !== 'function' ||
    !isPlainRecord(registration.policy) ||
    typeof registration.isOpen !== 'function' ||
    typeof registration.getRevision !== 'function' ||
    typeof registration.requestApproval !== 'function' ||
    typeof registration.captureSnapshot !== 'function' ||
    (registration.discardSnapshot !== undefined &&
      typeof registration.discardSnapshot !== 'function')
  ) {
    throw new ToolRouterError('invalid_tool_session')
  }
  if (
    jsonBytes(
      { tools: registration.skill.tools, policy: registration.policy },
      'tool_registration_limit',
    ) > MAX_REGISTRATION_BYTES
  ) {
    throw new ToolRouterError('tool_registration_limit')
  }
  const tools = new Map<string, AgentToolDef>()
  for (const tool of registration.skill.tools) {
    boundedString(tool?.name, MAX_NAME_BYTES, 'invalid_tool_definition')
    boundedString(tool?.description, MAX_DESCRIPTION_BYTES, 'invalid_tool_definition')
    if (
      !isPlainRecord(tool.inputSchema) ||
      jsonBytes(tool.inputSchema, 'invalid_tool_definition') > MAX_SCHEMA_BYTES
    ) {
      throw new ToolRouterError('invalid_tool_definition')
    }
    if (tools.has(tool.name)) throw new ToolRouterError('invalid_tool_definition')
    tools.set(tool.name, tool)
  }
  const policyNames = Object.keys(registration.policy)
  if (
    policyNames.length !== tools.size ||
    policyNames.some(
      (name) => !tools.has(name) || !['read', 'mutate'].includes(registration.policy[name] ?? ''),
    )
  ) {
    throw new ToolRouterError('invalid_tool_policy')
  }
  if (
    policyNames.some((name) => registration.policy[name] === 'mutate') &&
    (typeof registration.executeMutation !== 'function' ||
      typeof registration.validateMutation !== 'function')
  ) {
    throw new ToolRouterError('invalid_mutation_lifecycle')
  }
  return tools
}

function validCall(call: AgentToolCall): boolean {
  return (
    isPlainRecord(call) &&
    typeof call.id === 'string' &&
    call.id !== '' &&
    utf8Length(call.id) <= MAX_NAME_BYTES &&
    typeof call.name === 'string' &&
    call.name !== '' &&
    utf8Length(call.name) <= MAX_NAME_BYTES &&
    isPlainRecord(call.input) &&
    !call.inputError &&
    !call.truncated &&
    jsonBytes(call.input, 'invalid_tool_call') <= MAX_INPUT_BYTES
  )
}

function validateExecution(execution: ToolExecution, invalidMutated = false): ToolExecution {
  if (
    !isPlainRecord(execution) ||
    typeof execution.output !== 'string' ||
    utf8Length(execution.output) > MAX_OUTPUT_BYTES ||
    typeof execution.summary !== 'string' ||
    utf8Length(execution.summary) > MAX_SUMMARY_BYTES ||
    (execution.isError !== undefined && typeof execution.isError !== 'boolean') ||
    (execution.mutated !== undefined && typeof execution.mutated !== 'boolean')
  ) {
    return stableExecution('invalid_tool_result', 'Tool failed', true, invalidMutated)
  }
  return execution
}

function isOpen(session: ToolSession): boolean {
  if (session.closed) return false
  try {
    return session.registration.isOpen()
  } catch {
    return false
  }
}

function stopped(session: ToolSession, signal: AbortSignal): ToolExecution | undefined {
  if (signal.aborted) return stableExecution('tool_cancelled', 'Tool cancelled', true)
  if (!isOpen(session)) return stableExecution('tool_session_closed', 'Tool session closed', true)
  return undefined
}

function abortable<T>(value: T | PromiseLike<T>, signal: AbortSignal): Promise<T> {
  if (signal.aborted) return Promise.reject(new ToolRouterError('tool_cancelled'))
  return new Promise<T>((resolve, reject) => {
    const onAbort = (): void => reject(new ToolRouterError('tool_cancelled'))
    signal.addEventListener('abort', onAbort, { once: true })
    void Promise.resolve(value).then(
      (result) => {
        signal.removeEventListener('abort', onAbort)
        resolve(result)
      },
      (error: unknown) => {
        signal.removeEventListener('abort', onAbort)
        reject(error)
      },
    )
  })
}

export class DocumentToolRouter {
  readonly #randomBytes: (size: number) => Uint8Array
  readonly #sessions = new Map<string, ToolSession>()

  constructor(options: DocumentToolRouterOptions = {}) {
    this.#randomBytes = options.randomBytes ?? nodeRandomBytes
  }

  register(registration: ToolSessionRegistration): ToolSessionCredentials {
    const tools = validateRegistration(registration)
    let sessionId = ''
    for (
      let attempts = 0;
      attempts < 8 && (!sessionId || this.#sessions.has(sessionId));
      attempts++
    ) {
      sessionId = this.#newSecret()
    }
    if (!sessionId || this.#sessions.has(sessionId)) {
      throw new ToolRouterError('tool_session_entropy_failed')
    }
    const token = this.#newSecret()
    this.#sessions.set(sessionId, {
      token: Buffer.from(token, 'base64url'),
      registration,
      tools,
      pending: new Map(),
      closed: false,
    })
    return { sessionId, token }
  }

  listTools(credentials: ToolSessionCredentials): McpToolDefinition[] {
    const session = this.#authenticate(credentials)
    if (!isOpen(session)) throw new ToolRouterError('mcp_unauthorized')
    return [...session.tools.values()].map((tool) => {
      const readOnly = session.registration.policy[tool.name] === 'read'
      return {
        name: tool.name,
        description: tool.description,
        inputSchema: tool.inputSchema,
        annotations: { readOnlyHint: readOnly, destructiveHint: !readOnly },
      }
    })
  }

  async callTool(
    credentials: ToolSessionCredentials,
    call: AgentToolCall,
    signal?: AbortSignal,
  ): Promise<ToolExecution> {
    const session = this.#authenticate(credentials)
    if (!validCall(call)) return stableExecution('invalid_tool_call', 'Tool failed', true)
    if (!session.tools.has(call.name)) return stableExecution('unknown_tool', 'Tool failed', true)
    if (!isOpen(session)) return stableExecution('tool_session_closed', 'Tool session closed', true)
    if (session.pending.size > 0 || session.pending.has(call.id)) {
      return stableExecution('tool_call_in_progress', 'Tool busy', true)
    }

    const controller = new AbortController()
    const onAbort = (): void => controller.abort()
    signal?.addEventListener('abort', onAbort, { once: true })
    if (signal?.aborted) controller.abort()
    session.pending.set(call.id, controller)
    try {
      const early = stopped(session, controller.signal)
      if (early) return early
      const policy = session.registration.policy[call.name]
      if (policy === 'read') {
        let execution: ToolExecution
        try {
          execution = await abortable(
            session.registration.skill.executeTool(call, controller.signal),
            controller.signal,
          )
        } catch {
          return (
            stopped(session, controller.signal) ??
            stableExecution('tool_execution_failed', 'Tool failed', true)
          )
        }
        const after = stopped(session, controller.signal)
        if (after) return after
        const result = validateExecution(execution)
        if (result.mutated) return stableExecution('tool_policy_violation', 'Tool failed', true)
        return result
      }
      return await this.#callMutation(session, call, controller.signal)
    } finally {
      signal?.removeEventListener('abort', onAbort)
      session.pending.delete(call.id)
    }
  }

  cancel(credentials: ToolSessionCredentials, callId: string): boolean {
    const session = this.#authenticate(credentials)
    const pending = session.pending.get(callId)
    if (!pending) return false
    pending.abort()
    return true
  }

  close(credentials: ToolSessionCredentials): void {
    const session = this.#authenticate(credentials)
    session.closed = true
    for (const pending of session.pending.values()) pending.abort()
    this.#sessions.delete(credentials.sessionId)
  }

  closeAll(): void {
    for (const session of this.#sessions.values()) {
      session.closed = true
      for (const pending of session.pending.values()) pending.abort()
    }
    this.#sessions.clear()
  }

  async #callMutation(
    session: ToolSession,
    call: AgentToolCall,
    signal: AbortSignal,
  ): Promise<ToolExecution> {
    let capturedSnapshotId: string | undefined
    let guardedExecutionAcknowledged = false
    try {
      const revision = boundedString(
        await abortable(session.registration.getRevision(signal), signal),
        MAX_REVISION_BYTES,
        'invalid_document_revision',
      )
      let interrupted = stopped(session, signal)
      if (interrupted) return interrupted
      const approved = await abortable(
        session.registration.requestApproval(call, revision, signal),
        signal,
      )
      interrupted = stopped(session, signal)
      if (interrupted) return interrupted
      if (approved !== true) return stableExecution('tool_denied', 'Tool denied', false)

      const approvedRevision = boundedString(
        await abortable(session.registration.getRevision(signal), signal),
        MAX_REVISION_BYTES,
        'invalid_document_revision',
      )
      interrupted = stopped(session, signal)
      if (interrupted) return interrupted
      if (approvedRevision !== revision) {
        return stableExecution('document_changed', 'Document changed', true)
      }

      // Snapshot capture is safety-critical cleanup state. Await its definitive result even when
      // cancellation races it so a materialized snapshot can be explicitly discarded below.
      const snapshotId = boundedString(
        await session.registration.captureSnapshot(call, revision, signal),
        MAX_SNAPSHOT_ID_BYTES,
        'invalid_snapshot',
      )
      capturedSnapshotId = snapshotId
      interrupted = stopped(session, signal)
      if (interrupted) return interrupted
      const snapshotRevision = boundedString(
        await abortable(session.registration.getRevision(signal), signal),
        MAX_REVISION_BYTES,
        'invalid_document_revision',
      )
      interrupted = stopped(session, signal)
      if (interrupted) return interrupted
      if (snapshotRevision !== revision) {
        return stableExecution('document_changed', 'Document changed', true)
      }

      const guard = { expectedRevision: revision, snapshotId }
      // Once guarded commit begins, cancellation is advisory: retain the session lock until the
      // renderer definitively acknowledges cancellation or returns the commit result. This avoids
      // reporting a retry-safe cancellation while an abort-ignoring mutation is still running.
      const execution = validateExecution(
        await session.registration.executeMutation!(call, guard, signal),
        true,
      )
      // A successful renderer acknowledgement is the only point at which the router knows the
      // prepared snapshot was promoted into committed undo state. A rejection still belongs to
      // the pre-commit cleanup path; an acknowledgement followed by validation failure does not.
      guardedExecutionAcknowledged = true
      // A committed mutation still requires validation even if the turn was cancelled meanwhile.
      try {
        await session.registration.validateMutation!(call, execution, guard)
      } catch {
        return stableExecution(
          'tool_validation_failed',
          'Tool validation failed',
          true,
          execution.mutated === true,
        )
      }
      return execution
    } catch (error) {
      return (
        stopped(session, signal) ??
        stableExecution(
          error instanceof ToolRouterError ? error.code : 'tool_execution_failed',
          'Tool failed',
          true,
        )
      )
    } finally {
      if (
        capturedSnapshotId !== undefined &&
        !guardedExecutionAcknowledged &&
        session.registration.discardSnapshot
      ) {
        try {
          await session.registration.discardSnapshot(call, capturedSnapshotId)
        } catch {
          // Cleanup ownership is now uncertain. Fail closed for every later call in this session.
          session.closed = true
        }
      }
    }
  }

  #authenticate(credentials: ToolSessionCredentials): ToolSession {
    if (
      !credentials ||
      typeof credentials.sessionId !== 'string' ||
      typeof credentials.token !== 'string'
    ) {
      throw new ToolRouterError('mcp_unauthorized')
    }
    const session = this.#sessions.get(credentials.sessionId)
    let candidate: Buffer
    try {
      candidate = Buffer.from(credentials.token, 'base64url')
    } catch {
      throw new ToolRouterError('mcp_unauthorized')
    }
    if (
      !session ||
      candidate.length !== session.token.length ||
      !timingSafeEqual(candidate, session.token)
    ) {
      throw new ToolRouterError('mcp_unauthorized')
    }
    return session
  }

  authorize(credentials: ToolSessionCredentials): void {
    const session = this.#authenticate(credentials)
    if (!isOpen(session)) throw new ToolRouterError('mcp_unauthorized')
  }

  #newSecret(): string {
    const bytes = Buffer.from(this.#randomBytes(TOKEN_BYTES))
    if (bytes.length !== TOKEN_BYTES) throw new ToolRouterError('tool_session_entropy_failed')
    return bytes.toString('base64url')
  }
}
