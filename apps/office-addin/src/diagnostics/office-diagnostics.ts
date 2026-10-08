import {
  parsePresentationQaAttempt,
  type PresentationQaAttempt,
} from '../skills/powerpoint/presentation-qa-attempts.js'
import type { OfficeHost } from '../office-document.js'
import { acpPresentationStage, type PresentationStage } from '@wiswork/agent-harness'

export const MAX_LOCAL_DIAGNOSTIC_EVENTS = 200
export const MAX_DIAGNOSTIC_EVENT_BYTES = 4 * 1024
export const MAX_DIAGNOSTIC_EXPORT_BYTES = 256 * 1024

export type DiagnosticPhase =
  'run' | 'tool' | 'proposal' | 'validate' | 'write' | 'verify' | 'recovery' | 'transport'
export type DiagnosticOutcome = 'passed' | 'failed' | 'unsupported' | 'cancelled'
export type VerificationStage = 'text' | 'body_shape' | 'content' | 'boundary'
export interface PresentationDiagnosticContext {
  document_id?: string
  session_id?: string
  project_id?: string
  request_id?: string
  page_id?: string
  tool_call_id?: string
}

const DIAGNOSTIC_TOOL_ERRORS = new Set([
  'cancelled',
  'design_contract_review_required',
  'design_contract_prototype_required',
  'design_contract_production_incomplete',
  'design_contract_verification_failed',
  'design_contract_visual_review_failed',
  'design_contract_invalid_status',
  'design_contract_review_not_pending',
  'design_contract_acceptance_mismatch',
  'design_contract_screenshot_required',
  'image_fetch_unavailable',
  'image_limit',
  'image_mime_unsupported',
  'invalid_image',
  'invalid_tool_input',
  'office_api_unsupported',
  'office_read_failed',
  'office_screenshot_unavailable',
  'office_overwrite_required',
  'office_recovery_failed',
  'office_concurrent_change',
  'office_state_uncertain',
  'office_verify_failed',
  'office_write_failed',
  'office_write_pending',
  'proposal_missing',
  'proposal_stale',
  'presentation_screenshot_waiting',
  'presentation_qa_failed',
  'presentation_qa_capture_invalid',
  'presentation_qa_capture_required',
  'presentation_qa_stale',
  'presentation_qa_busy',
  'presentation_qa_state_invalid',
  'presentation_qa_page_not_imported',
  'presentation_qa_attempt_unresolved',
  'presentation_qa_attempt_history_full',
  'presentation_qa_attempt_state_invalid',
  'presentation_session_storage_full',
])
export function isDiagnosticToolError(value: string): boolean {
  return DIAGNOSTIC_TOOL_ERRORS.has(value) || /^office_recovery_failed:word_[a-z_]+$/.test(value)
}
const ERROR_CODES = new Set([
  ...DIAGNOSTIC_TOOL_ERRORS,
  'agent_run_completed',
  'agent_run_failed',
  'auth_required',
  'diagnostic_upload_failed',
  'network_error',
  'provider_unavailable',
  'request_timeout',
  'session_expired',
  'transport_stream_budget_exceeded',
])

export interface OfficeDiagnosticEvent {
  event_id: string
  trace_id: string
  timestamp_ms: number
  host: Exclude<OfficeHost, 'unknown'>
  platform: string
  build: string
  tool: string
  phase: DiagnosticPhase
  outcome: DiagnosticOutcome
  error_code: string
  office_error_code?: string
  office_error_name?: string
  office_error_location?: string
  verification_stage?: VerificationStage
  presentation_stage?: PresentationStage
  duration_ms: number
  requirement_sets: Readonly<Record<string, boolean>>
  presentation_context?: Readonly<PresentationDiagnosticContext>
}

export interface OfficeDiagnosticSnapshot {
  trace_id?: string
  events: readonly OfficeDiagnosticEvent[]
}

export interface OfficeDiagnostics {
  startTrace(): string
  setTool(name: string, context?: PresentationDiagnosticContext): void
  record(input: {
    phase: DiagnosticPhase
    errorCode: string
    error?: unknown
    durationMs?: number
  }): OfficeDiagnosticEvent
  snapshot(): OfficeDiagnosticSnapshot
  exportJson(options?: {
    includeLocalContext?: boolean
    screenshotAttempts?: () => PresentationQaAttempt[]
  }): string
  clear(): void
}

interface DiagnosticOptions {
  host: Exclude<OfficeHost, 'unknown'>
  platform?: string
  build: string
  localDocumentId?: string
  localSessionId?: () => string | undefined
  requirementSets?: Readonly<Record<string, boolean>>
  remoteEnabled?: boolean
  remoteSamplePercent?: number
  send?: (event: OfficeDiagnosticEvent) => void | Promise<void>
  randomUUID?: () => string
  now?: () => number
}

const PLATFORMS = new Set(['pc', 'mac', 'office_online', 'ios', 'android', 'universal', 'unknown'])
const REQUIREMENT_SETS = new Set(['OfficeApi', 'WordApi', 'ExcelApi', 'PowerPointApi'])
const VERIFICATION_STAGES = new Set<VerificationStage>([
  'text',
  'body_shape',
  'content',
  'boundary',
])

const PLATFORM_NAMES: Readonly<Record<string, string>> = Object.freeze({
  pc: 'pc',
  mac: 'mac',
  officeonline: 'office_online',
  office_online: 'office_online',
  ios: 'ios',
  android: 'android',
  universal: 'universal',
})

export function officeDiagnosticEnvironment(
  host: Exclude<OfficeHost, 'unknown'>,
  root: Record<string, any> = globalThis as unknown as Record<string, any>,
): { platform: string; requirementSets: Readonly<Record<string, boolean>> } {
  const context = root.Office?.context
  const platformValue = typeof context?.platform === 'string' ? context.platform.toLowerCase() : ''
  const platform = PLATFORM_NAMES[platformValue] ?? 'unknown'
  const requirement = {
    word: ['WordApi', '1.3'],
    excel: ['ExcelApi', '1.3'],
    powerpoint: ['PowerPointApi', '1.2'],
  }[host] as [string, string]
  const supports = context?.requirements?.isSetSupported
  const supported = (() => {
    try {
      return (
        typeof supports === 'function' &&
        supports.call(context.requirements, requirement[0], requirement[1]) === true
      )
    } catch {
      return false
    }
  })()
  return Object.freeze({
    platform,
    requirementSets: Object.freeze({ [requirement[0]]: supported }),
  })
}

const encoder = new TextEncoder()
const identifier = (value: unknown, fallback: string, maximum = 128): string => {
  if (typeof value !== 'string') return fallback
  const normalized = value.trim().slice(0, maximum)
  return normalized && /^[A-Za-z0-9_.:/()-]+$/.test(normalized) ? normalized : fallback
}
const stableError = (value: unknown): string =>
  typeof value === 'string' && (ERROR_CODES.has(value) || isDiagnosticToolError(value))
    ? value
    : 'office_write_failed'
const outcome = (code: string): DiagnosticOutcome =>
  code === 'agent_run_completed'
    ? 'passed'
    : code === 'office_api_unsupported' || code === 'presentation_screenshot_waiting'
      ? 'unsupported'
      : code === 'cancelled'
        ? 'cancelled'
        : 'failed'

type OfficeDiagnosticMetadata = Pick<
  OfficeDiagnosticEvent,
  'office_error_code' | 'office_error_name' | 'office_error_location' | 'verification_stage'
>

function safeProperty(value: Record<string, unknown>, property: string): unknown {
  try {
    return value[property]
  } catch {
    return undefined
  }
}

function officeIdentifier(value: unknown): string {
  const normalized = identifier(value, '')
  return /^[A-Za-z_][A-Za-z0-9_.()-]*$/.test(normalized) ? normalized : ''
}

export function officeIdentifiers(error: unknown): OfficeDiagnosticMetadata {
  const result: OfficeDiagnosticMetadata = {}
  const seen = new Set<unknown>()
  let current = error
  for (let depth = 0; depth < 3; depth += 1) {
    if (!current || (typeof current !== 'object' && typeof current !== 'function')) break
    if (seen.has(current)) break
    seen.add(current)
    const value = current as Record<string, unknown>
    const debugInfoValue = safeProperty(value, 'debugInfo')
    const debugInfo =
      debugInfoValue && typeof debugInfoValue === 'object'
        ? (debugInfoValue as Record<string, unknown>)
        : undefined
    const code = officeIdentifier(safeProperty(value, 'code'))
    const name = officeIdentifier(safeProperty(value, 'name'))
    const location = officeIdentifier(
      debugInfo ? safeProperty(debugInfo, 'errorLocation') : undefined,
    )
    if (code && !result.office_error_code) result.office_error_code = code
    if (
      name &&
      (!result.office_error_name || (result.office_error_name === 'Error' && name !== 'Error'))
    )
      result.office_error_name = name
    if (location && !result.office_error_location) result.office_error_location = location
    const stage = safeProperty(value, 'verificationStage')
    if (
      typeof stage === 'string' &&
      VERIFICATION_STAGES.has(stage as VerificationStage) &&
      !result.verification_stage
    )
      result.verification_stage = stage as VerificationStage
    current = safeProperty(value, 'cause')
  }
  return result
}

function requirementSets(value: Readonly<Record<string, boolean>> | undefined) {
  const entries = Object.entries(value ?? {})
    .filter(([name, supported]) => typeof supported === 'boolean' && REQUIREMENT_SETS.has(name))
    .slice(0, 16)
  return Object.freeze(Object.fromEntries(entries) as Record<string, boolean>)
}

function presentationContext(value: PresentationDiagnosticContext | undefined) {
  if (!value) return undefined
  const allowed = [
    'document_id',
    'session_id',
    'project_id',
    'request_id',
    'page_id',
    'tool_call_id',
  ] as const
  const safe = Object.fromEntries(
    allowed.flatMap((key) =>
      typeof value[key] === 'string' && /^[A-Za-z0-9_#-]{1,128}$/.test(value[key])
        ? [[key, value[key]]]
        : [],
    ),
  ) as PresentationDiagnosticContext
  return Object.keys(safe).length ? Object.freeze(safe) : undefined
}

function freezeEvent(event: OfficeDiagnosticEvent): OfficeDiagnosticEvent {
  return Object.freeze({
    ...event,
    requirement_sets: Object.freeze({ ...event.requirement_sets }),
    ...(event.presentation_context
      ? { presentation_context: Object.freeze({ ...event.presentation_context }) }
      : {}),
  })
}

function localScreenshotAttempts(provider: () => PresentationQaAttempt[]) {
  const scope = 'retained_visible_presentation_task' as const
  try {
    const supplied = provider()
    if (!Array.isArray(supplied) || supplied.length > 64) throw Error('invalid')
    const attempts = Array.from(supplied, parsePresentationQaAttempt)
    const ids = new Set(attempts.map((a) => a.id))
    const task = (a: PresentationQaAttempt) =>
      JSON.stringify([a.documentId, a.source, a.projectId, a.requestId, a.artifactDigest])
    if (
      ids.size !== attempts.length ||
      attempts.some((a) => task(a) !== task(attempts[0]!)) ||
      encoder.encode(JSON.stringify(attempts)).byteLength > 128 * 1024
    )
      throw Error('invalid')
    return {
      scope,
      status: 'available' as const,
      attempts,
      record_count: attempts.length,
      unresolved_count: attempts.filter((a) => a.status === 'started').length,
    }
  } catch {
    return { scope, status: 'unavailable' as const }
  }
}

export function createOfficeDiagnostics(options: DiagnosticOptions): OfficeDiagnostics {
  const samplePercent = options.remoteSamplePercent ?? 100
  if (!Number.isInteger(samplePercent) || samplePercent < 0 || samplePercent > 100)
    throw new Error('invalid_office_diagnostic_sample_percent')
  const randomUUID = options.randomUUID ?? (() => crypto.randomUUID())
  const now = options.now ?? (() => Date.now())
  const requirements = requirementSets(options.requirementSets)
  const candidatePlatform = identifier(options.platform, 'unknown', 32).toLowerCase()
  const platform = PLATFORMS.has(candidatePlatform) ? candidatePlatform : 'unknown'
  const build = identifier(options.build, 'unknown', 64)
  const documentContext = presentationContext({ document_id: options.localDocumentId })
  let events: OfficeDiagnosticEvent[] = []
  let traceId: string | undefined
  let traceGeneration = 0
  let tool = 'unknown'
  let context: Readonly<PresentationDiagnosticContext> | undefined
  let uploadFailureRecorded = false

  const local = (event: OfficeDiagnosticEvent) => {
    events = [...events, freezeEvent(event)].slice(-MAX_LOCAL_DIAGNOSTIC_EVENTS)
  }
  const uploadFailureEvent = (event: OfficeDiagnosticEvent): OfficeDiagnosticEvent => {
    const derived = {
      ...event,
      event_id: randomUUID(),
      timestamp_ms: Math.max(0, Math.trunc(now())),
      phase: 'transport' as const,
      outcome: 'failed' as const,
      error_code: 'diagnostic_upload_failed',
      duration_ms: 0,
    }
    delete derived.office_error_code
    delete derived.office_error_name
    delete derived.office_error_location
    delete derived.verification_stage
    delete derived.presentation_context
    return derived
  }
  const upload = (event: OfficeDiagnosticEvent) => {
    if (!options.remoteEnabled || !options.send || event.error_code === 'diagnostic_upload_failed')
      return
    if (samplePercent !== 100) {
      let hash = 2_166_136_261
      for (const char of event.trace_id) hash = Math.imul(hash ^ char.charCodeAt(0), 16_777_619)
      if ((hash >>> 0) % 100 >= samplePercent) return
    }
    const generation = traceGeneration
    try {
      const { presentation_context: _localContext, ...remoteEvent } = event
      const result = options.send(remoteEvent)
      void Promise.resolve(result).catch(() => {
        if (generation !== traceGeneration || uploadFailureRecorded) return
        uploadFailureRecorded = true
        local(uploadFailureEvent(event))
      })
    } catch {
      if (generation === traceGeneration && !uploadFailureRecorded) {
        uploadFailureRecorded = true
        local(uploadFailureEvent(event))
      }
    }
  }

  return {
    startTrace() {
      traceGeneration += 1
      traceId = randomUUID()
      tool = 'unknown'
      context = undefined
      uploadFailureRecorded = false
      return traceId
    },
    setTool(name, inputContext) {
      tool = identifier(name, 'unknown', 128)
      context = presentationContext({
        project_id: inputContext?.project_id,
        request_id: inputContext?.request_id,
        page_id: inputContext?.page_id,
        tool_call_id: inputContext?.tool_call_id,
      })
    },
    record(input) {
      if (!traceId) {
        traceGeneration += 1
        traceId = randomUUID()
      }
      const errorCode = stableError(input.errorCode)
      const presentationStage = acpPresentationStage(tool)
      const sessionContext = presentationContext({ session_id: options.localSessionId?.() })
      const event = freezeEvent({
        event_id: randomUUID(),
        trace_id: traceId,
        timestamp_ms: Math.max(0, Math.trunc(now())),
        host: options.host,
        platform,
        build,
        tool,
        phase: input.phase,
        outcome: outcome(errorCode),
        error_code: errorCode,
        ...officeIdentifiers(input.error),
        ...(presentationStage ? { presentation_stage: presentationStage } : {}),
        duration_ms:
          Number.isFinite(input.durationMs) && input.durationMs! >= 0
            ? // Match the existing Relay diagnostic bound, not a fictitious 10-minute run cap.
              Math.min(86_400_000, Math.trunc(input.durationMs!))
            : 0,
        requirement_sets: requirements,
        ...(documentContext || sessionContext || context
          ? { presentation_context: { ...context, ...sessionContext, ...documentContext } }
          : {}),
      })
      if (encoder.encode(JSON.stringify(event)).byteLength > MAX_DIAGNOSTIC_EVENT_BYTES) {
        throw new Error('invalid_diagnostic_event')
      }
      local(event)
      upload(event)
      return event
    },
    snapshot: () => Object.freeze({ trace_id: traceId, events: Object.freeze([...events]) }),
    exportJson(exportOptions) {
      const exportedEvents =
        exportOptions?.includeLocalContext === true
          ? events
          : events.map(({ presentation_context: _localContext, ...event }) => event)
      const attempts =
        exportOptions?.includeLocalContext === true && exportOptions.screenshotAttempts
          ? localScreenshotAttempts(exportOptions.screenshotAttempts)
          : undefined
      let omitted = 0
      const serialize = () =>
        JSON.stringify(
          {
            version: 1,
            trace_id: traceId,
            events: exportedEvents.slice(omitted),
            ...(attempts ? { local_presentation_qa_attempts: attempts } : {}),
            ...(omitted ? { omitted_event_count: omitted } : {}),
          },
          null,
          2,
        )
      let value = serialize()
      while (
        attempts &&
        encoder.encode(value).byteLength > MAX_DIAGNOSTIC_EXPORT_BYTES &&
        omitted < exportedEvents.length
      ) {
        omitted++
        value = serialize()
      }
      if (encoder.encode(value).byteLength > MAX_DIAGNOSTIC_EXPORT_BYTES)
        throw new Error('diagnostic_export_too_large')
      return value
    },
    clear() {
      traceGeneration += 1
      events = []
      traceId = undefined
      tool = 'unknown'
      context = undefined
      uploadFailureRecorded = false
    },
  }
}
