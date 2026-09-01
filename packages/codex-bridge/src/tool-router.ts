import { randomBytes, timingSafeEqual } from 'node:crypto'
import type {
  AgentSkill,
  AgentToolCall,
  AgentToolDef,
  ToolExecution,
  ToolExecutionOutcome,
} from '@wiswork/agent-core'
import type {
  DocumentCarrierHandle,
  DocumentCarrierIssuer,
  DocumentCarrierTurnContext,
} from './types.js'

const SECRET_BYTES = 32
const MAX_TOOLS = 256
const MAX_NAME_BYTES = 128
const MAX_DESCRIPTION_BYTES = 256_000
const MAX_SCHEMA_BYTES = 1_000_000
const MAX_INPUT_BYTES = 1_000_000
const MAX_OUTPUT_BYTES = 1_000_000
const MAX_SUMMARY_BYTES = 4_096
const MAX_ID_BYTES = 256
const MAX_GRAPH_NODES = 20_000
const MAX_GRAPH_DEPTH = 48
const MAX_CALL_MS = 30_000

const HOSTS = new Set([
  'latex',
  'slides',
  'docs',
  'sheets',
  'office-word',
  'office-excel',
  'office-powerpoint',
])
const DENIED_TOOL_NAMES =
  /(?:^|_)(?:shell|exec_command|filesystem|file_system|git|browser|network|fetch|write_stdin)(?:_|$)/i

export type ToolMutability = 'read' | 'mutate'

export interface DocumentToolIdentity {
  readonly ownerId: string
  readonly host:
    'latex' | 'slides' | 'docs' | 'sheets' | 'office-word' | 'office-excel' | 'office-powerpoint'
  readonly documentId: string
  readonly sessionId: string
  readonly generation: number
}

export interface DocumentToolRegistration {
  readonly identity: DocumentToolIdentity
  readonly skill: AgentSkill
  readonly policy: Readonly<Record<string, ToolMutability>>
  readonly isOpen: () => boolean
  /** Optional Task 2 issuer. The capability stays in this host-owned closure. */
  readonly carrier?: Readonly<{ issuer: DocumentCarrierIssuer; capability: unknown }>
  readonly maxCallMs?: number
}

export interface ToolSessionCredentials {
  readonly sessionId: string
  readonly secret: string
}

export interface McpToolDefinition extends AgentToolDef {
  readonly annotations: { readonly readOnlyHint: boolean; readonly destructiveHint: boolean }
}

export interface DocumentToolSession {
  readonly identity: Readonly<DocumentToolIdentity>
  readonly credentials: ToolSessionCredentials
  listTools(credentials: ToolSessionCredentials): McpToolDefinition[]
  callTool(
    credentials: ToolSessionCredentials,
    call: AgentToolCall,
    signal?: AbortSignal,
  ): Promise<ToolExecution>
  issueCarrier(
    credentials: ToolSessionCredentials,
    turn: Omit<DocumentCarrierTurnContext, 'capability'>,
  ): DocumentCarrierHandle
  cancel(credentials: ToolSessionCredentials, callId: string): boolean
  close(): void
}

export class ToolRouterError extends Error {
  constructor(readonly code: string) {
    super(code)
    this.name = 'ToolRouterError'
  }
}

const utf8Length = (value: string): number => Buffer.byteLength(value, 'utf8')

function plainRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false
  try {
    const prototype = Object.getPrototypeOf(value)
    return (
      (prototype === Object.prototype || prototype === null) &&
      Object.getOwnPropertySymbols(value).length === 0 &&
      Object.values(Object.getOwnPropertyDescriptors(value)).every((item) => 'value' in item)
    )
  } catch {
    return false
  }
}

function inspectGraph(value: unknown, maxBytes: number, code: string): void {
  let encoded: string
  try {
    encoded = JSON.stringify(value)
  } catch {
    throw new ToolRouterError(code)
  }
  if (encoded === undefined || utf8Length(encoded) > maxBytes) throw new ToolRouterError(code)
  const pending: Array<{ value: unknown; depth: number }> = [{ value, depth: 0 }]
  const seen = new Set<object>()
  let nodes = 0
  while (pending.length > 0) {
    const item = pending.pop()!
    nodes += 1
    if (nodes > MAX_GRAPH_NODES || item.depth > MAX_GRAPH_DEPTH) throw new ToolRouterError(code)
    if (typeof item.value === 'string' && utf8Length(item.value) > maxBytes)
      throw new ToolRouterError(code)
    if (typeof item.value !== 'object' || item.value === null) continue
    if (seen.has(item.value)) throw new ToolRouterError(code)
    seen.add(item.value)
    if (Array.isArray(item.value)) {
      for (const child of item.value) pending.push({ value: child, depth: item.depth + 1 })
    } else {
      if (!plainRecord(item.value)) throw new ToolRouterError(code)
      for (const child of Object.values(item.value))
        pending.push({ value: child, depth: item.depth + 1 })
    }
  }
}

function boundedId(value: unknown): value is string {
  return typeof value === 'string' && value.length > 0 && utf8Length(value) <= MAX_ID_BYTES
}

function validateIdentity(value: unknown): Readonly<DocumentToolIdentity> {
  if (!plainRecord(value)) throw new ToolRouterError('invalid_tool_identity')
  if (
    Object.keys(value).length !== 5 ||
    !boundedId(value.ownerId) ||
    typeof value.host !== 'string' ||
    !HOSTS.has(value.host) ||
    !boundedId(value.documentId) ||
    !boundedId(value.sessionId) ||
    !Number.isSafeInteger(value.generation) ||
    (value.generation as number) < 0
  )
    throw new ToolRouterError('invalid_tool_identity')
  return Object.freeze({
    ownerId: value.ownerId,
    host: value.host as DocumentToolIdentity['host'],
    documentId: value.documentId,
    sessionId: value.sessionId,
    generation: value.generation as number,
  })
}

function validateRegistration(
  registration: DocumentToolRegistration,
): ReadonlyMap<string, AgentToolDef> {
  if (
    !plainRecord(registration) ||
    !registration.skill ||
    !plainRecord(registration.skill) ||
    !Array.isArray(registration.skill.tools) ||
    registration.skill.tools.length < 1 ||
    registration.skill.tools.length > MAX_TOOLS ||
    typeof registration.skill.executeTool !== 'function' ||
    !plainRecord(registration.policy) ||
    typeof registration.isOpen !== 'function'
  )
    throw new ToolRouterError('invalid_tool_session')
  const tools = new Map<string, AgentToolDef>()
  for (const tool of registration.skill.tools) {
    if (
      !plainRecord(tool) ||
      !boundedId(tool.name) ||
      utf8Length(tool.name) > MAX_NAME_BYTES ||
      DENIED_TOOL_NAMES.test(tool.name) ||
      typeof tool.description !== 'string' ||
      utf8Length(tool.description) > MAX_DESCRIPTION_BYTES ||
      !plainRecord(tool.inputSchema) ||
      tools.has(tool.name)
    )
      throw new ToolRouterError('invalid_tool_definition')
    inspectGraph(tool.inputSchema, MAX_SCHEMA_BYTES, 'invalid_tool_definition')
    tools.set(tool.name, tool)
  }
  const policies = Object.keys(registration.policy)
  if (
    policies.length !== tools.size ||
    policies.some(
      (name) =>
        !tools.has(name) ||
        (registration.policy[name] !== 'read' && registration.policy[name] !== 'mutate'),
    )
  ) {
    throw new ToolRouterError('invalid_tool_policy')
  }
  if (
    registration.carrier !== undefined &&
    (!plainRecord(registration.carrier) ||
      typeof registration.carrier.issuer?.issueForTurn !== 'function')
  ) {
    throw new ToolRouterError('invalid_carrier_issuer')
  }
  return tools
}

function stable(
  output: string,
  summary = 'Tool failed',
  isError = true,
  mutated = false,
): ToolExecution {
  return { output, summary, isError, mutated }
}

function validExecution(value: unknown): value is ToolExecution {
  if (
    !plainRecord(value) ||
    typeof value.output !== 'string' ||
    typeof value.summary !== 'string' ||
    utf8Length(value.output) > MAX_OUTPUT_BYTES ||
    utf8Length(value.summary) > MAX_SUMMARY_BYTES ||
    (value.isError !== undefined && typeof value.isError !== 'boolean') ||
    (value.mutated !== undefined && typeof value.mutated !== 'boolean')
  )
    return false
  if (value.modelContent !== undefined) {
    try {
      inspectGraph(value.modelContent, MAX_OUTPUT_BYTES, 'invalid_tool_result')
    } catch {
      return false
    }
  }
  return true
}

function isSuspension(value: ToolExecutionOutcome): value is ToolExecutionOutcome & {
  kind: 'tool-execution-suspension'
  result: Promise<ToolExecution>
} {
  return (
    plainRecord(value) &&
    value.kind === 'tool-execution-suspension' &&
    value.result instanceof Promise
  )
}

function awaitBounded<T>(
  pending: PromiseLike<T>,
  signal: AbortSignal,
  maxCallMs: number,
): Promise<T> {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(
      () => finish(() => reject(new ToolRouterError('tool_timeout'))),
      maxCallMs,
    )
    timer.unref()
    const onAbort = (): void => finish(() => reject(new ToolRouterError('tool_cancelled')))
    let settled = false
    const finish = (action: () => void): void => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      signal.removeEventListener('abort', onAbort)
      action()
    }
    signal.addEventListener('abort', onAbort, { once: true })
    Promise.resolve(pending).then(
      (value) => finish(() => resolve(value)),
      () => finish(() => reject(new ToolRouterError('tool_execution_failed'))),
    )
    if (signal.aborted) onAbort()
  })
}

export function createDocumentToolSession(
  registration: DocumentToolRegistration,
): DocumentToolSession {
  const identity = validateIdentity(registration?.identity)
  const tools = validateRegistration(registration)
  const maxCallMs = registration.maxCallMs ?? MAX_CALL_MS
  if (!Number.isSafeInteger(maxCallMs) || maxCallMs <= 0 || maxCallMs > MAX_CALL_MS)
    throw new ToolRouterError('invalid_tool_deadline')
  const sessionId = randomBytes(SECRET_BYTES).toString('base64url')
  const secretBytes = randomBytes(SECRET_BYTES)
  const credentials = Object.freeze({ sessionId, secret: secretBytes.toString('base64url') })
  const pending = new Map<string, AbortController>()
  let closed = false

  const authenticate = (candidate: ToolSessionCredentials): void => {
    let provided: Buffer
    try {
      provided = Buffer.from(candidate?.secret ?? '', 'base64url')
    } catch {
      throw new ToolRouterError('tool_unauthorized')
    }
    if (
      candidate?.sessionId !== sessionId ||
      provided.length !== secretBytes.length ||
      !timingSafeEqual(provided, secretBytes)
    ) {
      throw new ToolRouterError('tool_unauthorized')
    }
    if (closed || !safeOpen()) throw new ToolRouterError('tool_session_closed')
  }
  const safeOpen = (): boolean => {
    try {
      return registration.isOpen() === true
    } catch {
      return false
    }
  }

  const listTools = (candidate: ToolSessionCredentials): McpToolDefinition[] => {
    authenticate(candidate)
    return [...tools.values()].map((tool) =>
      Object.freeze({
        name: tool.name,
        description: tool.description,
        inputSchema: tool.inputSchema,
        annotations: Object.freeze({
          readOnlyHint: registration.policy[tool.name] === 'read',
          destructiveHint: registration.policy[tool.name] === 'mutate',
        }),
      }),
    )
  }

  const callTool = async (
    candidate: ToolSessionCredentials,
    call: AgentToolCall,
    outerSignal?: AbortSignal,
  ): Promise<ToolExecution> => {
    authenticate(candidate)
    if (
      !plainRecord(call) ||
      !boundedId(call.id) ||
      !boundedId(call.name) ||
      !plainRecord(call.input) ||
      call.inputError ||
      call.truncated
    )
      return stable('invalid_tool_call')
    try {
      inspectGraph(call.input, MAX_INPUT_BYTES, 'invalid_tool_call')
    } catch {
      return stable('invalid_tool_call')
    }
    if (!tools.has(call.name) || registration.policy[call.name] === undefined)
      return stable('unknown_tool')
    if (pending.size > 0 || pending.has(call.id)) return stable('tool_call_in_progress')
    const controller = new AbortController()
    const abort = (): void => controller.abort()
    outerSignal?.addEventListener('abort', abort, { once: true })
    if (outerSignal?.aborted) abort()
    pending.set(call.id, controller)
    try {
      if (controller.signal.aborted) return stable('tool_cancelled', 'Tool cancelled')
      let outcome: ToolExecutionOutcome
      try {
        outcome = await awaitBounded(
          Promise.resolve(registration.skill.executeTool(call, controller.signal)),
          controller.signal,
          maxCallMs,
        )
      } catch (error) {
        return stable(error instanceof ToolRouterError ? error.code : 'tool_execution_failed')
      }
      if (registration.policy[call.name] === 'read') {
        if (isSuspension(outcome) || !validExecution(outcome) || outcome.mutated === true)
          return stable('tool_policy_violation')
        return outcome
      }
      // The bridge owns no approval, snapshot, or write primitive. Only Agent Core's existing
      // suspended host-authority contract can complete a mutation.
      if (!isSuspension(outcome))
        return stable('mutation_authority_required', 'Mutation authority required', true, false)
      let execution: ToolExecution
      try {
        execution = await awaitBounded(outcome.result, controller.signal, maxCallMs)
      } catch (error) {
        return stable(error instanceof ToolRouterError ? error.code : 'tool_execution_failed')
      }
      if (!validExecution(execution))
        return stable('invalid_tool_result', 'Tool failed', true, true)
      return execution
    } finally {
      outerSignal?.removeEventListener('abort', abort)
      pending.delete(call.id)
    }
  }

  const session: DocumentToolSession = {
    identity,
    credentials,
    listTools,
    callTool,
    issueCarrier(
      candidate: ToolSessionCredentials,
      turn: Omit<DocumentCarrierTurnContext, 'capability'>,
    ) {
      authenticate(candidate)
      if (!registration.carrier) throw new ToolRouterError('carrier_not_available')
      return registration.carrier.issuer.issueForTurn({
        ...turn,
        capability: registration.carrier.capability,
      })
    },
    cancel(candidate: ToolSessionCredentials, callId: string) {
      authenticate(candidate)
      const controller = pending.get(callId)
      controller?.abort()
      return controller !== undefined
    },
    close() {
      if (closed) return
      closed = true
      for (const controller of pending.values()) controller.abort()
      pending.clear()
    },
  }
  return Object.freeze(session)
}
