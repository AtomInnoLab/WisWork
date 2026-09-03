import { randomBytes } from 'node:crypto'
import { mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs'
import { dirname } from 'node:path'
import { ENHANCED_HOSTS, type EnhancedHost } from '@wiswork/agent-runtime'

const MAX_TASKS = 10
const MAX_EVENTS = 256
const MAX_SYSTEM_EVENTS = 64
const MAX_FILE_BYTES = 512 * 1024
const DETAIL_DURATION_MS = 30 * 60_000
const MAX_PROBE_BYTES = 64 * 1024
const NORMAL_MILESTONE_CODES = new Set<DiagnosticSafeCode>([
  'upstream_started',
  'mcp_ready',
  'mcp_tool_started',
  'mcp_tool_completed',
  'turn_started',
  'turn_completed',
])

export type DiagnosticComponent = 'component' | 'auth' | 'runtime' | 'wisusage' | 'mcp' | 'host'
export type DiagnosticOutcome = 'started' | 'succeeded' | 'failed' | 'cancelled' | 'blocked'
export type DiagnosticTaskStatus = 'running' | 'succeeded' | 'failed' | 'cancelled'
export type DiagnosticSafeCode =
  | 'component_unavailable'
  | 'auth_required'
  | 'runtime_starting'
  | 'runtime_ready'
  | 'runtime_event'
  | 'runtime_unavailable'
  | 'runtime_crashed'
  | 'upstream_started'
  | 'upstream_timeout'
  | 'upstream_unavailable'
  | 'upstream_auth_failed'
  | 'upstream_rate_limited'
  | 'upstream_rejected'
  | 'upstream_invalid_content_type'
  | 'stream_protocol_rejected'
  | 'stream_ended_early'
  | 'stream_reasoning_unsupported'
  | 'stream_reasoning_limit_exceeded'
  | 'stream_usage_invalid'
  | 'stream_tool_input_invalid'
  | 'mcp_starting'
  | 'mcp_ready'
  | 'mcp_tool_started'
  | 'mcp_tool_completed'
  | 'mcp_tool_denied'
  | 'turn_started'
  | 'turn_completed'
  | 'turn_failed'
  | 'turn_timeout'
  | 'unknown_failure'

const SAFE_CODES = new Set<DiagnosticSafeCode>([
  'component_unavailable',
  'auth_required',
  'runtime_starting',
  'runtime_ready',
  'runtime_event',
  'runtime_unavailable',
  'runtime_crashed',
  'upstream_started',
  'upstream_timeout',
  'upstream_unavailable',
  'upstream_auth_failed',
  'upstream_rate_limited',
  'upstream_rejected',
  'upstream_invalid_content_type',
  'stream_protocol_rejected',
  'stream_ended_early',
  'stream_reasoning_unsupported',
  'stream_reasoning_limit_exceeded',
  'stream_usage_invalid',
  'stream_tool_input_invalid',
  'mcp_starting',
  'mcp_ready',
  'mcp_tool_started',
  'mcp_tool_completed',
  'mcp_tool_denied',
  'turn_started',
  'turn_completed',
  'turn_failed',
  'turn_timeout',
  'unknown_failure',
])
const COMPONENTS = new Set<DiagnosticComponent>([
  'component',
  'auth',
  'runtime',
  'wisusage',
  'mcp',
  'host',
])
const OUTCOMES = new Set<DiagnosticOutcome>([
  'started',
  'succeeded',
  'failed',
  'cancelled',
  'blocked',
])

export interface EnhancedDiagnosticEvent {
  readonly sequence: number
  readonly at: number
  readonly component: DiagnosticComponent
  readonly phase: string
  readonly outcome: DiagnosticOutcome
  readonly code: DiagnosticSafeCode
}

export interface EnhancedDiagnosticTask {
  readonly diagnosticId: string
  readonly host: EnhancedHost
  readonly startedAt: number
  readonly endedAt?: number
  readonly status: DiagnosticTaskStatus
  readonly failureCode?: DiagnosticSafeCode
  readonly events: readonly EnhancedDiagnosticEvent[]
}

export type SelfCheckLayer = 'component' | 'authentication' | 'runtime' | 'mcp' | 'wisusage'
export interface EnhancedSelfCheckItem {
  readonly layer: SelfCheckLayer
  readonly status: 'passed' | 'failed' | 'not_tested'
  readonly code?: DiagnosticSafeCode
}
export interface EnhancedSelfCheckResult {
  readonly diagnosticId: string
  readonly startedAt: number
  readonly endedAt: number
  readonly status: 'passed' | 'failed'
  readonly checks: readonly EnhancedSelfCheckItem[]
}

interface MutableTask {
  diagnosticId: string
  host: EnhancedHost
  startedAt: number
  endedAt?: number
  status: DiagnosticTaskStatus
  failureCode?: DiagnosticSafeCode
  events: EnhancedDiagnosticEvent[]
}

interface PersistedState {
  readonly version: 1
  readonly tasks: readonly EnhancedDiagnosticTask[]
}

export interface EnhancedDiagnosticsStoreOptions {
  readonly path: string
  readonly now?: () => number
  readonly id?: () => string
}

export interface EnhancedDiagnosticExportMetadata {
  readonly appVersion: string
  readonly componentVersion: string
  readonly platform: 'darwin' | 'win32' | 'linux'
  readonly arch: 'arm64' | 'x64'
}

const copy = <T>(value: T): T => structuredClone(value)
const boundedPhase = (value: string): string =>
  /^[a-z][a-z0-9_]{0,47}$/.test(value) ? value : 'unknown'

export async function probeWisUsageEventStream(response: Response): Promise<boolean> {
  if (!response.ok || !response.body) return false
  if (!/^text\/event-stream(?:;|$)/i.test(response.headers.get('content-type') ?? '')) return false
  const reader = response.body.getReader()
  const decoder = new TextDecoder('utf-8', { fatal: true })
  let bytes = 0
  let buffer = ''
  try {
    while (bytes <= MAX_PROBE_BYTES && !/(?:\r\n|\r|\n){2}/.test(buffer)) {
      const next = await reader.read()
      if (next.done) break
      bytes += next.value.byteLength
      if (bytes > MAX_PROBE_BYTES) return false
      buffer += decoder.decode(next.value, { stream: true })
    }
    buffer += decoder.decode()
    const frame = buffer.split(/(?:\r\n|\r|\n){2}/, 1)[0] ?? ''
    const data = frame
      .split(/\r\n|\r|\n/)
      .filter((line) => line.startsWith('data:'))
      .map((line) => line.slice(5).trimStart())
      .join('\n')
    if (!data || data === '[DONE]') return false
    const parsed: unknown = JSON.parse(data)
    return (
      typeof parsed === 'object' &&
      parsed !== null &&
      !Array.isArray(parsed) &&
      typeof (parsed as { type?: unknown }).type === 'string' &&
      /^[a-z][a-z0-9_]{0,47}$/.test((parsed as { type: string }).type)
    )
  } catch {
    return false
  } finally {
    await reader.cancel().catch(() => undefined)
  }
}

function safeDiagnostic(code: string): {
  component: DiagnosticComponent
  phase: string
  outcome: DiagnosticOutcome
  code: DiagnosticSafeCode
} {
  if (code === 'responses_upstream_started')
    return { component: 'wisusage', phase: 'request', outcome: 'started', code: 'upstream_started' }
  if (code === 'responses_upstream_timeout')
    return { component: 'wisusage', phase: 'request', outcome: 'failed', code: 'upstream_timeout' }
  if (code === 'responses_upstream_auth')
    return {
      component: 'wisusage',
      phase: 'request',
      outcome: 'failed',
      code: 'upstream_auth_failed',
    }
  if (code === 'responses_upstream_rate_limited')
    return {
      component: 'wisusage',
      phase: 'request',
      outcome: 'failed',
      code: 'upstream_rate_limited',
    }
  if (code === 'responses_upstream_rejected')
    return { component: 'wisusage', phase: 'request', outcome: 'failed', code: 'upstream_rejected' }
  if (code === 'responses_upstream_content_type')
    return {
      component: 'wisusage',
      phase: 'stream',
      outcome: 'failed',
      code: 'upstream_invalid_content_type',
    }
  if (code === 'responses_upstream_unavailable' || code === 'responses_upstream_empty')
    return {
      component: 'wisusage',
      phase: 'request',
      outcome: 'failed',
      code: 'upstream_unavailable',
    }
  if (code.startsWith('responses_stream_'))
    return {
      component: 'wisusage',
      phase: 'stream',
      outcome: 'failed',
      code: code.includes('premature_messages_eof')
        ? 'stream_ended_early'
        : code.includes('reasoning_content_limit_exceeded')
          ? 'stream_reasoning_limit_exceeded'
          : code.includes('invalid_messages_usage')
            ? 'stream_usage_invalid'
            : code.includes('unsupported_reasoning_block')
              ? 'stream_reasoning_unsupported'
              : code.includes('invalid_custom_tool_input')
                ? 'stream_tool_input_invalid'
                : 'stream_protocol_rejected',
    }
  if (code === 'enhanced_response_incompatible')
    return {
      component: 'wisusage',
      phase: 'stream',
      outcome: 'failed',
      code: 'stream_protocol_rejected',
    }
  if (code === 'enhanced_turn_timeout')
    return { component: 'host', phase: 'turn', outcome: 'failed', code: 'turn_timeout' }
  if (code === 'enhanced_turn_failed')
    return { component: 'host', phase: 'turn', outcome: 'failed', code: 'turn_failed' }
  if (code === 'codex_turn_started')
    return { component: 'runtime', phase: 'turn', outcome: 'started', code: 'turn_started' }
  if (code === 'codex_turn_completed')
    return { component: 'runtime', phase: 'turn', outcome: 'succeeded', code: 'turn_completed' }
  if (code === 'enhanced_runtime_crashed' || code.startsWith('codex_process_exit'))
    return { component: 'runtime', phase: 'process', outcome: 'failed', code: 'runtime_crashed' }
  if (code === 'gateway_tool_call_received' || code === 'mcp_tools_call_received')
    return { component: 'mcp', phase: 'tool', outcome: 'started', code: 'mcp_tool_started' }
  if (code === 'gateway_tool_call_completed')
    return { component: 'mcp', phase: 'tool', outcome: 'succeeded', code: 'mcp_tool_completed' }
  if (code === 'gateway_tool_call_denied' || code === 'mcp_request_failed')
    return { component: 'mcp', phase: 'tool', outcome: 'failed', code: 'mcp_tool_denied' }
  if (code === 'gateway_tools_list' || code === 'mcp_tools_list')
    return { component: 'mcp', phase: 'initialize', outcome: 'succeeded', code: 'mcp_ready' }
  if (code.startsWith('mcp_') || code.includes('mcpServer_startupStatus'))
    return { component: 'mcp', phase: 'initialize', outcome: 'started', code: 'mcp_starting' }
  if (code === 'auth_required' || code === 'enhanced_auth_required')
    return { component: 'auth', phase: 'session', outcome: 'failed', code: 'auth_required' }
  if (code === 'enhanced_runtime_unavailable' || code === 'enhanced_start_failed')
    return {
      component: 'runtime',
      phase: 'process',
      outcome: 'failed',
      code: 'runtime_unavailable',
    }
  if (
    code === 'app_server_error' ||
    code === 'codex_error' ||
    code === 'app_server_thread_status_systemError'
  )
    return {
      component: 'runtime',
      phase: 'protocol',
      outcome: 'failed',
      code: 'unknown_failure',
    }
  return {
    component: 'runtime',
    phase: boundedPhase(code),
    outcome: 'started',
    code: 'runtime_event',
  }
}

function isPersistedState(value: unknown): value is PersistedState {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const state = value as { version?: unknown; tasks?: unknown }
  if (state.version !== 1 || !Array.isArray(state.tasks) || state.tasks.length > MAX_TASKS)
    return false
  return state.tasks.every((task) => {
    if (!task || typeof task !== 'object' || Array.isArray(task)) return false
    const item = task as Partial<EnhancedDiagnosticTask>
    const keys = Object.keys(task)
    if (
      Object.getPrototypeOf(task) !== Object.prototype ||
      Object.getOwnPropertySymbols(task).length !== 0 ||
      keys.some(
        (key) =>
          ![
            'diagnosticId',
            'host',
            'startedAt',
            'endedAt',
            'status',
            'failureCode',
            'events',
          ].includes(key),
      )
    )
      return false
    return (
      typeof item.diagnosticId === 'string' &&
      /^diag_[A-Za-z0-9_-]{24}$/.test(item.diagnosticId) &&
      typeof item.host === 'string' &&
      ENHANCED_HOSTS.includes(item.host as EnhancedHost) &&
      typeof item.startedAt === 'number' &&
      Number.isSafeInteger(item.startedAt) &&
      (item.endedAt === undefined ||
        (Number.isSafeInteger(item.endedAt) && item.endedAt >= item.startedAt)) &&
      ['running', 'succeeded', 'failed', 'cancelled'].includes(item.status ?? '') &&
      (item.failureCode === undefined || SAFE_CODES.has(item.failureCode)) &&
      Array.isArray(item.events) &&
      item.events.length <= MAX_EVENTS &&
      item.events.every((event) => {
        if (!event || typeof event !== 'object' || Array.isArray(event)) return false
        const candidate = event as Partial<EnhancedDiagnosticEvent>
        return (
          Object.getPrototypeOf(event) === Object.prototype &&
          Object.getOwnPropertySymbols(event).length === 0 &&
          Object.keys(event).length === 6 &&
          ['sequence', 'at', 'component', 'phase', 'outcome', 'code'].every((key) =>
            Object.hasOwn(event, key),
          ) &&
          Number.isSafeInteger(candidate.sequence) &&
          Number.isSafeInteger(candidate.at) &&
          typeof candidate.component === 'string' &&
          COMPONENTS.has(candidate.component as DiagnosticComponent) &&
          typeof candidate.phase === 'string' &&
          boundedPhase(candidate.phase) === candidate.phase &&
          typeof candidate.outcome === 'string' &&
          OUTCOMES.has(candidate.outcome as DiagnosticOutcome) &&
          typeof candidate.code === 'string' &&
          SAFE_CODES.has(candidate.code as DiagnosticSafeCode)
        )
      })
    )
  })
}

export class EnhancedDiagnosticsStore {
  readonly #path: string
  readonly #now: () => number
  readonly #id: () => string
  readonly #tasks: MutableTask[]
  readonly #active = new Set<string>()
  readonly #systemEvents: EnhancedDiagnosticEvent[] = []
  #sequence = 0
  #detailedUntil = 0
  #lastSelfCheck: EnhancedSelfCheckResult | undefined

  constructor(options: EnhancedDiagnosticsStoreOptions) {
    this.#path = options.path
    this.#now = options.now ?? Date.now
    this.#id = options.id ?? (() => `diag_${randomBytes(18).toString('base64url')}`)
    this.#tasks = this.#load()
  }

  beginTask(host: EnhancedHost): string {
    const diagnosticId = this.#id()
    if (!/^diag_[A-Za-z0-9_-]{24}$/.test(diagnosticId)) throw new Error('invalid_diagnostic_id')
    const task: MutableTask = {
      diagnosticId,
      host,
      startedAt: this.#now(),
      status: 'running',
      events: [],
    }
    this.#tasks.unshift(task)
    this.#tasks.splice(MAX_TASKS)
    this.#active.add(diagnosticId)
    this.#append(task, {
      component: 'host',
      phase: 'turn',
      outcome: 'started',
      code: 'turn_started',
    })
    this.#persist()
    return diagnosticId
  }

  record(rawCode: string): void {
    const event = safeDiagnostic(rawCode)
    const targets = this.#tasks.filter((task) => this.#active.has(task.diagnosticId))
    if (targets.length !== 1) {
      this.#appendTo(this.#systemEvents, event, MAX_SYSTEM_EVENTS)
      return
    }
    for (const task of targets) {
      if (
        this.#now() < this.#detailedUntil ||
        event.outcome === 'failed' ||
        NORMAL_MILESTONE_CODES.has(event.code)
      ) {
        this.#append(task, event)
      }
    }
  }

  hasObserved(code: DiagnosticSafeCode): boolean {
    return (
      this.#systemEvents.some((event) => event.code === code) ||
      this.#tasks.some((task) => task.events.some((event) => event.code === code))
    )
  }

  finishTask(
    diagnosticId: string,
    status: Exclude<DiagnosticTaskStatus, 'running'>,
    rawFailureCode?: string,
  ): void {
    const task = this.#tasks.find((candidate) => candidate.diagnosticId === diagnosticId)
    if (!task || task.status !== 'running') return
    this.#active.delete(diagnosticId)
    task.status = status
    task.endedAt = this.#now()
    if (status === 'failed') {
      task.failureCode =
        task.events.find((event) => event.outcome === 'failed')?.code ??
        safeDiagnostic(rawFailureCode ?? '').code
    }
    this.#append(task, {
      component: 'host',
      phase: 'turn',
      outcome:
        status === 'succeeded' ? 'succeeded' : status === 'cancelled' ? 'cancelled' : 'failed',
      code:
        status === 'succeeded'
          ? 'turn_completed'
          : status === 'cancelled'
            ? 'turn_completed'
            : task.failureCode!,
    })
    this.#persist()
  }

  enableDetailed(): number {
    this.#detailedUntil = this.#now() + DETAIL_DURATION_MS
    return this.#detailedUntil
  }

  detailedUntil(): number | null {
    return this.#detailedUntil > this.#now() ? this.#detailedUntil : null
  }

  recent(): readonly EnhancedDiagnosticTask[] {
    return copy(this.#tasks)
  }

  exportReport(
    metadata: EnhancedDiagnosticExportMetadata,
    selfCheck: EnhancedSelfCheckResult | undefined = this.#lastSelfCheck,
  ) {
    const report = {
      schema: 'wiswork-enhanced-diagnostics/v1',
      generatedAt: this.#now(),
      metadata,
      detailedUntil: this.#detailedUntil > this.#now() ? this.#detailedUntil : null,
      tasks: this.recent(),
      systemEvents: copy(this.#systemEvents),
      ...(selfCheck ? { selfCheck: copy(selfCheck) } : {}),
    }
    const serialized = `${JSON.stringify(report, null, 2)}\n`
    if (Buffer.byteLength(serialized) > MAX_FILE_BYTES)
      throw new Error('diagnostic_report_too_large')
    return serialized
  }

  async runSelfCheck(
    probes: Readonly<Record<SelfCheckLayer, () => Promise<boolean | 'not_tested'>>>,
  ): Promise<EnhancedSelfCheckResult> {
    const diagnosticId = this.#id()
    const startedAt = this.#now()
    const checks: EnhancedSelfCheckItem[] = []
    for (const layer of ['component', 'authentication', 'runtime', 'mcp', 'wisusage'] as const) {
      try {
        const result = await probes[layer]()
        checks.push({
          layer,
          status: result === 'not_tested' ? 'not_tested' : result ? 'passed' : 'failed',
        })
      } catch {
        checks.push({ layer, status: 'failed', code: 'unknown_failure' })
      }
    }
    const result = Object.freeze({
      diagnosticId,
      startedAt,
      endedAt: this.#now(),
      status: checks.some((item) => item.status === 'failed') ? 'failed' : 'passed',
      checks: Object.freeze(checks),
    })
    this.#lastSelfCheck = result
    return result
  }

  #append(task: MutableTask, event: Omit<EnhancedDiagnosticEvent, 'sequence' | 'at'>): void {
    this.#appendTo(task.events, event, MAX_EVENTS)
  }

  #appendTo(
    target: EnhancedDiagnosticEvent[],
    event: Omit<EnhancedDiagnosticEvent, 'sequence' | 'at'>,
    maximum: number,
  ): void {
    target.push({ sequence: ++this.#sequence, at: this.#now(), ...event })
    if (target.length > maximum) target.splice(0, target.length - maximum)
  }

  #load(): MutableTask[] {
    try {
      const bytes = readFileSync(this.#path)
      if (bytes.byteLength > MAX_FILE_BYTES) return []
      const parsed: unknown = JSON.parse(bytes.toString('utf8'))
      if (!isPersistedState(parsed)) return []
      return parsed.tasks.map((task) => ({
        ...copy(task),
        status: task.status === 'running' ? 'failed' : task.status,
        ...(task.status === 'running' ? { failureCode: 'runtime_crashed' as const } : {}),
        events: [...task.events],
      }))
    } catch {
      return []
    }
  }

  #persist(): void {
    try {
      const serialized = `${JSON.stringify({ version: 1, tasks: this.#tasks } satisfies PersistedState)}\n`
      if (Buffer.byteLength(serialized) > MAX_FILE_BYTES) return
      mkdirSync(dirname(this.#path), { recursive: true })
      const temporary = `${this.#path}.${randomBytes(12).toString('hex')}.tmp`
      try {
        writeFileSync(temporary, serialized, { encoding: 'utf8', mode: 0o600, flag: 'wx' })
        renameSync(temporary, this.#path)
      } finally {
        rmSync(temporary, { force: true })
      }
    } catch {
      // Diagnostics are fail-open and must never affect a document task.
    }
  }
}
