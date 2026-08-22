import type { OfficeHost } from '../office-document.js'

export const MAX_LOCAL_DIAGNOSTIC_EVENTS = 200
export const MAX_DIAGNOSTIC_EVENT_BYTES = 4 * 1024
export const MAX_DIAGNOSTIC_EXPORT_BYTES = 256 * 1024

export type DiagnosticPhase =
  'tool' | 'proposal' | 'validate' | 'write' | 'verify' | 'recovery' | 'transport'
export type DiagnosticOutcome = 'failed' | 'unsupported' | 'cancelled'

const ERROR_CODES = new Set([
  'agent_run_failed',
  'auth_required',
  'cancelled',
  'diagnostic_upload_failed',
  'network_error',
  'office_api_unsupported',
  'office_read_failed',
  'office_recovery_failed',
  'office_verify_failed',
  'office_write_failed',
  'proposal_missing',
  'proposal_stale',
  'provider_unavailable',
  'request_timeout',
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
  duration_ms: number
  requirement_sets: Readonly<Record<string, boolean>>
}

export interface OfficeDiagnosticSnapshot {
  trace_id?: string
  events: readonly OfficeDiagnosticEvent[]
}

export interface OfficeDiagnostics {
  startTrace(): string
  setTool(name: string): void
  record(input: {
    phase: DiagnosticPhase
    errorCode: string
    error?: unknown
    durationMs?: number
  }): OfficeDiagnosticEvent
  snapshot(): OfficeDiagnosticSnapshot
  exportJson(): string
  clear(): void
}

interface DiagnosticOptions {
  host: Exclude<OfficeHost, 'unknown'>
  platform?: string
  build: string
  requirementSets?: Readonly<Record<string, boolean>>
  remoteEnabled?: boolean
  send?: (event: OfficeDiagnosticEvent) => void | Promise<void>
  randomUUID?: () => string
  now?: () => number
}

const PLATFORMS = new Set(['pc', 'mac', 'office_online', 'ios', 'android', 'universal', 'unknown'])
const REQUIREMENT_SETS = new Set(['OfficeApi', 'WordApi', 'ExcelApi', 'PowerPointApi'])

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
  typeof value === 'string' && ERROR_CODES.has(value) ? value : 'office_write_failed'
const outcome = (code: string): DiagnosticOutcome =>
  code === 'office_api_unsupported' ? 'unsupported' : code === 'cancelled' ? 'cancelled' : 'failed'

function officeIdentifiers(
  error: unknown,
): Pick<
  OfficeDiagnosticEvent,
  'office_error_code' | 'office_error_name' | 'office_error_location'
> {
  if (!error || (typeof error !== 'object' && typeof error !== 'function')) return {}
  try {
    const value = error as Record<string, unknown>
    const debugInfo =
      value.debugInfo && typeof value.debugInfo === 'object'
        ? (value.debugInfo as Record<string, unknown>)
        : undefined
    const code = identifier(value.code, '')
    const name = identifier(value.name, '')
    const location = identifier(debugInfo?.errorLocation, '')
    return {
      ...(code ? { office_error_code: code } : {}),
      ...(name ? { office_error_name: name } : {}),
      ...(location ? { office_error_location: location } : {}),
    }
  } catch {
    return {}
  }
}

function requirementSets(value: Readonly<Record<string, boolean>> | undefined) {
  const entries = Object.entries(value ?? {})
    .filter(([name, supported]) => typeof supported === 'boolean' && REQUIREMENT_SETS.has(name))
    .slice(0, 16)
  return Object.freeze(Object.fromEntries(entries) as Record<string, boolean>)
}

function freezeEvent(event: OfficeDiagnosticEvent): OfficeDiagnosticEvent {
  return Object.freeze({ ...event, requirement_sets: Object.freeze({ ...event.requirement_sets }) })
}

export function createOfficeDiagnostics(options: DiagnosticOptions): OfficeDiagnostics {
  const randomUUID = options.randomUUID ?? (() => crypto.randomUUID())
  const now = options.now ?? (() => Date.now())
  const requirements = requirementSets(options.requirementSets)
  const candidatePlatform = identifier(options.platform, 'unknown', 32).toLowerCase()
  const platform = PLATFORMS.has(candidatePlatform) ? candidatePlatform : 'unknown'
  const build = identifier(options.build, 'unknown', 64)
  let events: OfficeDiagnosticEvent[] = []
  let traceId: string | undefined
  let tool = 'unknown'
  let uploadFailureRecorded = false

  const local = (event: OfficeDiagnosticEvent) => {
    events = [...events, freezeEvent(event)].slice(-MAX_LOCAL_DIAGNOSTIC_EVENTS)
  }
  const upload = (event: OfficeDiagnosticEvent) => {
    if (!options.remoteEnabled || !options.send || event.error_code === 'diagnostic_upload_failed')
      return
    try {
      const result = options.send(event)
      void Promise.resolve(result).catch(() => {
        if (uploadFailureRecorded) return
        uploadFailureRecorded = true
        local(
          freezeEvent({
            ...event,
            event_id: randomUUID(),
            timestamp_ms: Math.max(0, Math.trunc(now())),
            phase: 'transport',
            outcome: 'failed',
            error_code: 'diagnostic_upload_failed',
            office_error_code: undefined,
            office_error_name: undefined,
            office_error_location: undefined,
            duration_ms: 0,
          }),
        )
      })
    } catch {
      if (!uploadFailureRecorded) {
        uploadFailureRecorded = true
        local({
          ...event,
          event_id: randomUUID(),
          timestamp_ms: Math.max(0, Math.trunc(now())),
          phase: 'transport',
          outcome: 'failed',
          error_code: 'diagnostic_upload_failed',
          duration_ms: 0,
        })
      }
    }
  }

  return {
    startTrace() {
      traceId = randomUUID()
      tool = 'unknown'
      uploadFailureRecorded = false
      return traceId
    },
    setTool(name) {
      tool = identifier(name, 'unknown', 128)
    },
    record(input) {
      if (!traceId) traceId = randomUUID()
      const errorCode = stableError(input.errorCode)
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
        duration_ms:
          Number.isFinite(input.durationMs) && input.durationMs! >= 0
            ? Math.min(600_000, Math.trunc(input.durationMs!))
            : 0,
        requirement_sets: requirements,
      })
      if (encoder.encode(JSON.stringify(event)).byteLength > MAX_DIAGNOSTIC_EVENT_BYTES) {
        throw new Error('invalid_diagnostic_event')
      }
      local(event)
      upload(event)
      return event
    },
    snapshot: () => Object.freeze({ trace_id: traceId, events: Object.freeze([...events]) }),
    exportJson() {
      const value = JSON.stringify({ version: 1, trace_id: traceId, events }, null, 2)
      if (encoder.encode(value).byteLength > MAX_DIAGNOSTIC_EXPORT_BYTES)
        throw new Error('diagnostic_export_too_large')
      return value
    },
    clear() {
      events = []
      traceId = undefined
      tool = 'unknown'
      uploadFailureRecorded = false
    },
  }
}
