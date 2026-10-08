import { describe, expect, it, vi } from 'vitest'
import { createEnhancedTelemetry, parseEnhancedTelemetryEvent } from '../src/telemetry'

describe('Enhanced aggregate telemetry', () => {
  it('accepts only closed component and seven-host workflow dimensions', () => {
    const events: unknown[] = []
    const telemetry = createEnhancedTelemetry((event) => events.push(event))
    telemetry.component('download', 'succeeded')
    telemetry.host('office-powerpoint', 'verify', 'applied_unverified')
    expect(events).toEqual([
      { kind: 'component', phase: 'download', outcome: 'succeeded' },
      { kind: 'host', host: 'office-powerpoint', phase: 'verify', outcome: 'applied_unverified' },
    ])
  })

  it('strictly parses exact, bounded, plain closed-enum events', () => {
    expect(
      parseEnhancedTelemetryEvent({
        kind: 'host',
        host: 'docs',
        phase: 'correction',
        outcome: 'succeeded',
      }),
    ).toEqual({ kind: 'host', host: 'docs', phase: 'correction', outcome: 'succeeded' })
    for (const invalid of [
      { kind: 'host', host: 'docs', phase: 'complete', outcome: 'verified', documentId: 'secret' },
      { kind: 'host', host: 'unknown', phase: 'complete', outcome: 'verified' },
      { kind: 'component', phase: 'download', outcome: 'yes' },
      Object.assign(Object.create({ kind: 'component' }), {
        phase: 'download',
        outcome: 'succeeded',
      }),
      Object.defineProperty({ kind: 'component', outcome: 'succeeded' }, 'phase', {
        get: () => 'download',
      }),
      { kind: 'host', host: 'docs', phase: 'x'.repeat(100), outcome: 'failed' },
    ])
      expect(() => parseEnhancedTelemetryEvent(invalid)).toThrow('invalid_enhanced_telemetry')
  })

  it('is fail-open and never accepts identifiers or content', () => {
    const sink = vi.fn((_event: unknown) => {
      throw new Error('collector unavailable')
    })
    const telemetry = createEnhancedTelemetry(sink)
    expect(() => telemetry.host('latex', 'complete', 'verified')).not.toThrow()
    expect(Object.keys(sink.mock.calls[0]![0] as object).sort()).toEqual([
      'host',
      'kind',
      'outcome',
      'phase',
    ])
  })
})
