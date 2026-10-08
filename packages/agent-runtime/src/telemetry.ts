import { ENHANCED_HOSTS, type EnhancedHost } from './contracts'

export const ENHANCED_COMPONENT_PHASES = [
  'download',
  'verify',
  'install',
  'launch',
  'update',
  'remove',
] as const
export const ENHANCED_HOST_PHASES = [
  'plan',
  'dispatch',
  'verify',
  'complete',
  'correction',
  'pending',
] as const
export const ENHANCED_OUTCOMES = [
  'started',
  'succeeded',
  'failed',
  'cancelled',
  'blocked',
  'verified',
  'applied_unverified',
  'unchanged',
] as const

export type EnhancedComponentPhase = (typeof ENHANCED_COMPONENT_PHASES)[number]
export type EnhancedHostPhase = (typeof ENHANCED_HOST_PHASES)[number]
export type EnhancedTelemetryOutcome = (typeof ENHANCED_OUTCOMES)[number]
export type EnhancedTelemetryEvent =
  | Readonly<{
      kind: 'component'
      phase: EnhancedComponentPhase
      outcome: EnhancedTelemetryOutcome
    }>
  | Readonly<{
      kind: 'host'
      host: EnhancedHost
      phase: EnhancedHostPhase
      outcome: EnhancedTelemetryOutcome
    }>

export interface EnhancedTelemetry {
  component(phase: EnhancedComponentPhase, outcome: EnhancedTelemetryOutcome): void
  host(host: EnhancedHost, phase: EnhancedHostPhase, outcome: EnhancedTelemetryOutcome): void
}

const plainExact = (value: unknown, keys: readonly string[]): Record<string, unknown> => {
  if (
    typeof value !== 'object' ||
    value === null ||
    Object.getPrototypeOf(value) !== Object.prototype ||
    Object.getOwnPropertySymbols(value).length
  )
    throw new TypeError('invalid_enhanced_telemetry')
  const descriptors = Object.getOwnPropertyDescriptors(value)
  if (
    Object.keys(descriptors).sort().join('\0') !== [...keys].sort().join('\0') ||
    Object.values(descriptors).some((descriptor) => !('value' in descriptor))
  )
    throw new TypeError('invalid_enhanced_telemetry')
  return Object.fromEntries(
    Object.entries(descriptors).map(([key, descriptor]) => [key, descriptor.value]),
  )
}

export function parseEnhancedTelemetryEvent(value: unknown): EnhancedTelemetryEvent {
  const base = plainExact(
    value,
    (value as { kind?: unknown } | null)?.kind === 'component'
      ? ['kind', 'phase', 'outcome']
      : ['kind', 'host', 'phase', 'outcome'],
  )
  if (!ENHANCED_OUTCOMES.includes(base.outcome as EnhancedTelemetryOutcome))
    throw new TypeError('invalid_enhanced_telemetry')
  if (
    base.kind === 'component' &&
    ENHANCED_COMPONENT_PHASES.includes(base.phase as EnhancedComponentPhase)
  )
    return Object.freeze({
      kind: 'component',
      phase: base.phase,
      outcome: base.outcome,
    }) as EnhancedTelemetryEvent
  if (
    base.kind === 'host' &&
    ENHANCED_HOSTS.includes(base.host as EnhancedHost) &&
    ENHANCED_HOST_PHASES.includes(base.phase as EnhancedHostPhase)
  )
    return Object.freeze({
      kind: 'host',
      host: base.host,
      phase: base.phase,
      outcome: base.outcome,
    }) as EnhancedTelemetryEvent
  throw new TypeError('invalid_enhanced_telemetry')
}

/** Aggregate-only and deliberately fail-open: diagnostics must never alter task semantics. */
export function createEnhancedTelemetry(
  sink: (event: EnhancedTelemetryEvent) => void,
): EnhancedTelemetry {
  const emit = (event: EnhancedTelemetryEvent) => {
    try {
      sink(parseEnhancedTelemetryEvent(event))
    } catch {
      /* telemetry is non-authoritative */
    }
  }
  return Object.freeze({
    component: (phase: EnhancedComponentPhase, outcome: EnhancedTelemetryOutcome) =>
      emit({ kind: 'component', phase, outcome }),
    host: (host: EnhancedHost, phase: EnhancedHostPhase, outcome: EnhancedTelemetryOutcome) =>
      emit({ kind: 'host', host, phase, outcome }),
  })
}
