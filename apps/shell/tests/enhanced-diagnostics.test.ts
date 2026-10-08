import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { describe, expect, it } from 'vitest'
import recording from '../../../packages/codex-bridge/tests/fixtures/protocol-redacted-max-tokens.json'
import {
  EnhancedDiagnosticsStore,
  probeWisUsageEventStream,
} from '../src/main/enhanced-diagnostics'

const ids = Array.from({ length: 20 }, (_, index) => `diag_${String(index).padStart(24, '0')}`)

function fixture(now = 100) {
  let current = now
  let nextId = 0
  const root = mkdtempSync(join(tmpdir(), 'wiswork-diagnostics-'))
  const path = join(root, 'enhanced.json')
  return {
    path,
    tick: (amount = 1) => (current += amount),
    store: new EnhancedDiagnosticsStore({
      path,
      now: () => current,
      id: () => ids[nextId++]!,
    }),
  }
}

describe('EnhancedDiagnosticsStore', () => {
  it('exports strict preview build identity without unknown fields and labels capture ordering', () => {
    const { store, tick } = fixture()
    const metadata = {
      appVersion: '0.6.14-pr123.gabcdef0',
      componentVersion: '0.147.0',
      platform: 'darwin',
      arch: 'arm64',
      build: {
        commit: 'a'.repeat(40),
        mode: 'preview',
        pr: 123,
        builtAt: '2026-09-03T00:00:00.000Z',
        nested: { token: 'SECRET' },
      },
    } as const
    store.recordProtocol(recording)
    tick(5)
    store.recordProtocol(recording)
    const report = JSON.parse(store.exportReport(metadata))
    expect(report.metadata.appVersion).toBe(metadata.appVersion)
    expect(report.metadata.build).toEqual({
      commit: 'a'.repeat(40),
      mode: 'preview',
      pr: 123,
      builtAt: '2026-09-03T00:00:00.000Z',
    })
    expect(report.protocolRecordingInfo).toMatchObject([
      { index: 0, recordedAt: 100, association: 'unattributed' },
      { index: 1, recordedAt: 105, association: 'unattributed' },
    ])
    expect(report.protocolRecordingInfo[0].recordingId).toMatch(/^recording_[A-Za-z0-9_-]{24}$/)
    expect(report.protocolRecordingInfo[0].recordingId).not.toBe(
      report.protocolRecordingInfo[1].recordingId,
    )
    expect(report.protocolRecordingInfo[0].originalOutcome).toBe('not_observed')
    store.recordProtocol(recording, 'protocol_rejected')
    expect(JSON.parse(store.exportReport(metadata)).protocolRecordingInfo[2].originalOutcome).toBe(
      'protocol_rejected',
    )
    expect(JSON.stringify(report)).not.toContain('SECRET')
    for (const build of [
      { ...metadata.build, commit: '/private/SECRET' },
      { ...metadata.build, mode: 'SECRET' },
      { ...metadata.build, pr: 1_000_000 },
      { ...metadata.build, builtAt: '2026-02-31T00:00:00.000Z' },
    ]) {
      expect(JSON.parse(store.exportReport({ ...metadata, build })).metadata.build).toBeUndefined()
    }
    expect(
      JSON.parse(store.exportReport({ ...metadata, appVersion: '0.6.14-private.SECRET' })).metadata
        .appVersion,
    ).toBe('unknown')
  })
  it('exports only bounded validated recordings and projects unknown metadata away', () => {
    const { store } = fixture()
    for (let i = 0; i < 6; i++) store.recordProtocol(recording)
    store.recordProtocol({ ...recording, secret: 'PRIVATE JWT' })
    const report = store.exportReport({
      appVersion: '0.1.0',
      componentVersion: '0.2.0',
      platform: 'darwin',
      arch: 'arm64',
      secret: { path: '/private/data' },
    } as any)
    expect(JSON.parse(report).protocolRecordings).toHaveLength(4)
    expect(report).not.toMatch(/PRIVATE|JWT|private|secret/)
    const withCheck = store.exportReport(
      { appVersion: '0.1.0', componentVersion: '0.2.0', platform: 'darwin', arch: 'arm64' },
      {
        diagnosticId: ids[0],
        startedAt: 0,
        endedAt: 1,
        status: 'passed',
        checks: [{ layer: 'runtime', status: 'passed', secret: 'PRIVATE' }],
        secret: 'PRIVATE',
      } as any,
    )
    expect(withCheck).not.toContain('PRIVATE')
  })
  it('validates WisUsage data-only SSE framing without retaining response content', async () => {
    await expect(
      probeWisUsageEventStream(
        new Response('data: {"type":"message_start","private":"secret"}\n\n', {
          headers: { 'content-type': 'text/event-stream; charset=utf-8' },
        }),
      ),
    ).resolves.toBe(true)
    await expect(
      probeWisUsageEventStream(
        new Response('data: private response body\n\n', {
          headers: { 'content-type': 'text/event-stream' },
        }),
      ),
    ).resolves.toBe(false)
  })
  it('correlates a task and maps raw runtime codes to closed safe events', () => {
    const { store, tick } = fixture()
    const id = store.beginTask('slides')
    tick()
    store.record('app_server_thread_started')
    store.record('responses_stream_invalid_messages_sse')
    store.record('private prompt jwt_123 /Users/person/file.pptx')
    store.finishTask(id, 'failed', 'enhanced_turn_timeout')
    const task = store.recent()[0]!
    expect(task).toMatchObject({ diagnosticId: id, host: 'slides', status: 'failed' })
    expect(task.events.map((event) => event.code)).toContain('stream_protocol_rejected')
    expect(task.failureCode).toBe('turn_timeout')
    expect(task.events.some((event) => event.phase === 'app_server_thread_started')).toBe(false)
    expect(JSON.stringify(task)).not.toContain('private prompt')
    expect(JSON.stringify(task)).not.toContain('/Users/')
  })

  it('distinguishes safe stream failure categories without retaining model content', () => {
    const { store } = fixture()
    const id = store.beginTask('slides')
    store.record('responses_stream_unsupported_reasoning_block')
    store.record('responses_stream_reasoning_content_limit_exceeded')
    store.record('responses_stream_invalid_messages_usage')
    store.record('responses_stream_invalid_custom_tool_input')
    store.record('responses_stream_tool_input_retry_limit_exceeded')
    store.record('responses_stream_invalid_messages_event_order')
    store.finishTask(id, 'failed')

    expect(store.recent()[0]?.events.map((event) => event.code)).toEqual(
      expect.arrayContaining([
        'stream_reasoning_unsupported',
        'stream_reasoning_limit_exceeded',
        'stream_usage_invalid',
        'stream_tool_input_invalid',
        'stream_tool_input_recovery_exhausted',
        'stream_event_order_invalid',
      ]),
    )
    expect(JSON.stringify(store.recent()[0])).not.toContain('reasoning_block')
  })

  it('does not report an interrupted stream as a protocol rejection', () => {
    const { store, path } = fixture()
    const id = store.beginTask('office-powerpoint')
    store.record('responses_stream_invalid')
    store.record('responses_stream_upstream_interrupted')
    store.record('responses_stream_invalid_messages_sse')
    store.finishTask(id, 'failed')

    const restored = new EnhancedDiagnosticsStore({ path })
    const codes = restored
      .recent()[0]!
      .events.filter((event) => event.phase === 'stream')
      .map((event) => event.code)
    expect(codes.filter((code) => code === 'stream_interrupted')).toHaveLength(2)
    expect(codes.filter((code) => code === 'stream_protocol_rejected')).toHaveLength(1)
  })

  it('keeps ten tasks, bounds detailed events, and preserves the first safe state after restart', () => {
    const test = fixture()
    test.store.enableDetailed()
    for (let index = 0; index < 12; index += 1) {
      const id = test.store.beginTask('docs')
      for (let event = 0; event < 300; event += 1) test.store.record('mcp_request_received')
      test.store.finishTask(id, 'succeeded')
    }
    expect(test.store.recent()).toHaveLength(10)
    expect(test.store.recent()[0]!.events.length).toBeLessThanOrEqual(256)
    const restored = new EnhancedDiagnosticsStore({ path: test.path })
    expect(restored.recent()).toHaveLength(10)
    expect(readFileSync(test.path).byteLength).toBeLessThan(512 * 1024)
  })

  it('turns an interrupted running task into a crash-safe failed summary', () => {
    const test = fixture()
    const id = test.store.beginTask('sheets')
    const restored = new EnhancedDiagnosticsStore({ path: test.path })
    expect(restored.recent()[0]).toMatchObject({
      diagnosticId: id,
      status: 'failed',
      failureCode: 'runtime_crashed',
    })
  })

  it('maps a generic enhanced turn failure without treating normal runtime events as failures', () => {
    const { store } = fixture()
    const id = store.beginTask('slides')
    store.record('app_server_thread_started')
    store.finishTask(id, 'failed', 'enhanced_turn_failed')
    expect(store.recent()[0]).toMatchObject({ failureCode: 'turn_failed' })
  })

  it('records semantic document tool failure distinctly from transport completion', () => {
    const { store } = fixture()
    const id = store.beginTask('office-powerpoint')
    store.record('gateway_tool_call_failed')
    store.finishTask(id, 'succeeded')
    expect(store.recent()[0]!.events).toContainEqual(
      expect.objectContaining({
        component: 'mcp',
        phase: 'tool',
        outcome: 'failed',
        code: 'mcp_tool_failed',
      }),
    )
  })

  it('records confirmation expiry distinctly from runtime failure', () => {
    const { store } = fixture()
    const id = store.beginTask('slides')
    store.finishTask(id, 'failed', 'enhanced_proposal_expired')
    expect(store.recent()[0]).toMatchObject({ failureCode: 'proposal_expired' })
  })

  it('preserves questionnaire terminal failure instead of an earlier repair failure', () => {
    const { store } = fixture()
    const id = store.beginTask('slides')
    store.record('enhanced_proposal_execution_failed')
    store.finishTask(id, 'failed', 'enhanced_questionnaire_incomplete')
    expect(store.recent()[0]).toMatchObject({ failureCode: 'questionnaire_incomplete' })
    expect(store.recent()[0]!.events.at(-1)).toMatchObject({
      code: 'questionnaire_incomplete',
      outcome: 'failed',
    })
  })

  it.each([
    'carrier_invalid',
    'carrier_input_invalid',
    'capability_invalid',
    'tool_unavailable',
    'carrier_mismatch',
    'proposal_summary_invalid',
    'proposal_outcome_invalid',
    'proposal_handler_unavailable',
  ])('restores denial reason %s and unrelated task history after restart', (reason) => {
    const { store, path } = fixture()
    const previous = store.beginTask('docs')
    store.finishTask(previous, 'succeeded')
    const current = store.beginTask('slides')
    store.record(`gateway_tool_call_denied_${reason}`)
    store.finishTask(current, 'failed', 'enhanced_questionnaire_incomplete')

    const restored = new EnhancedDiagnosticsStore({ path })
    expect(restored.recent()).toEqual(store.recent())
    expect(restored.recent()).toHaveLength(2)
    expect(restored.recent().find((task) => task.diagnosticId === current)).toMatchObject({
      failureCode: 'questionnaire_incomplete',
      events: expect.arrayContaining([
        expect.objectContaining({ code: 'mcp_tool_denied', phase: reason }),
      ]),
    })
  })

  it('keeps allowlisted per-call denial reasons without persisting arbitrary details', () => {
    const { store } = fixture()
    store.beginTask('slides')
    store.record('gateway_tool_call_denied_carrier_mismatch')
    store.record('gateway_tool_call_denied_capability_invalid')
    store.record('gateway_tool_call_denied_PRIVATE_CONTENT')
    const task = store.recent()[0]!
    expect(task.events).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ code: 'mcp_tool_denied', phase: 'carrier_mismatch' }),
        expect.objectContaining({ code: 'mcp_tool_denied', phase: 'capability_invalid' }),
      ]),
    )
    expect(JSON.stringify(task)).not.toContain('PRIVATE_CONTENT')
  })

  it('reports missing carrier tools and binding rejection without arbitrary details', () => {
    const { store } = fixture()
    store.beginTask('office-powerpoint')
    store.record('resolver_method_missing')
    store.record('resolver_protocol_carrier_authorization_mismatch')
    store.record('resolver_protocol_PRIVATE_CONTENT')
    expect(store.recent()[0]?.events).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          component: 'mcp',
          phase: 'initialize',
          outcome: 'failed',
          code: 'mcp_tools_missing',
        }),
        expect.objectContaining({
          component: 'runtime',
          phase: 'protocol',
          outcome: 'failed',
          code: 'carrier_authorization_mismatch',
        }),
      ]),
    )
    expect(JSON.stringify(store.recent())).not.toContain('PRIVATE_CONTENT')
  })

  it('keeps local start and app-server failure boundaries distinct', () => {
    const { store } = fixture()
    const id = store.beginTask('slides')
    for (const code of [
      'enhanced_thread_starting',
      'enhanced_thread_start_failed',
      'enhanced_turn_starting',
      'enhanced_turn_start_failed',
      'app_server_error',
      'codex_error',
      'app_server_thread_status_systemError',
    ]) {
      store.record(code)
    }
    store.finishTask(id, 'failed')

    expect(store.recent()[0]?.events.map((event) => event.code)).toEqual(
      expect.arrayContaining([
        'thread_starting',
        'thread_start_failed',
        'turn_starting',
        'turn_start_failed',
        'app_server_error',
        'codex_error',
        'thread_system_error',
      ]),
    )
  })

  it('records only closed proposal lifecycle states', () => {
    const { store } = fixture()
    const id = store.beginTask('slides')
    for (const code of [
      'enhanced_proposal_created',
      'enhanced_proposal_applied',
      'enhanced_proposal_cancelled',
      'enhanced_proposal_execution_failed',
    ]) {
      store.record(code)
    }
    store.finishTask(id, 'failed')

    expect(store.recent()[0]?.events.map((event) => event.code)).toEqual(
      expect.arrayContaining([
        'proposal_created',
        'proposal_applied',
        'proposal_cancelled',
        'proposal_execution_failed',
      ]),
    )
  })

  it('rejects a tampered persisted report instead of re-exporting injected content', () => {
    const test = fixture()
    writeFileSync(
      test.path,
      JSON.stringify({
        version: 1,
        tasks: [
          {
            diagnosticId: ids[0],
            host: 'slides',
            startedAt: 1,
            status: 'failed',
            events: [{ prompt: 'private document content' }],
          },
        ],
      }),
    )
    const restored = new EnhancedDiagnosticsStore({ path: test.path })
    expect(restored.recent()).toEqual([])
    expect(
      restored.exportReport({
        appVersion: '0.6.10',
        componentVersion: '0.147.0',
        platform: 'linux',
        arch: 'x64',
      }),
    ).not.toContain('private document content')
  })

  it('runs all self-check layers without exposing probe errors', async () => {
    const { store } = fixture()
    const result = await store.runSelfCheck({
      component: async () => true,
      authentication: async () => true,
      runtime: async () => true,
      mcp: async () => true,
      wisusage: async () => {
        throw new Error('Bearer private')
      },
    })
    expect(result.status).toBe('failed')
    expect(result.checks.at(-1)).toEqual({
      layer: 'wisusage',
      status: 'failed',
      code: 'unknown_failure',
    })
    expect(JSON.stringify(result)).not.toContain('Bearer private')
  })

  it('preserves safe failure codes returned by self-check probes', async () => {
    const { store } = fixture()
    const result = await store.runSelfCheck({
      component: async () => ({ status: 'failed', code: 'component_unavailable' }),
      authentication: async () => true,
      runtime: async () => ({ status: 'failed', code: 'runtime_unavailable' }),
      mcp: async () => 'not_tested',
      wisusage: async () => true,
    })
    expect(result.checks[0]).toEqual({
      layer: 'component',
      status: 'failed',
      code: 'component_unavailable',
    })
    expect(result.checks[2]).toEqual({
      layer: 'runtime',
      status: 'failed',
      code: 'runtime_unavailable',
    })
  })

  it('keeps ambiguous concurrent runtime events system-scoped instead of misattributing them', () => {
    const { store } = fixture()
    const slides = store.beginTask('slides')
    const docs = store.beginTask('docs')
    store.record('responses_upstream_timeout')
    store.finishTask(slides, 'failed', 'enhanced_turn_timeout')
    store.finishTask(docs, 'failed', 'enhanced_turn_timeout')
    expect(
      store
        .recent()
        .every((task) => task.events.every((event) => event.code !== 'upstream_timeout')),
    ).toBe(true)
    expect(
      store.exportReport({
        appVersion: '0.6.10',
        componentVersion: '0.147.0',
        platform: 'linux',
        arch: 'x64',
      }),
    ).toContain('upstream_timeout')
  })

  it('exports only bounded schema-approved metadata and events', () => {
    const { store } = fixture()
    const id = store.beginTask('latex')
    store.finishTask(id, 'cancelled')
    const report = store.exportReport({
      appVersion: '0.6.10',
      componentVersion: '0.147.0',
      platform: 'darwin',
      arch: 'arm64',
    })
    expect(JSON.parse(report)).toMatchObject({ schema: 'wiswork-enhanced-diagnostics/v1' })
    expect(report).not.toContain('prompt')
    expect(Buffer.byteLength(report)).toBeLessThan(512 * 1024)
  })
})
