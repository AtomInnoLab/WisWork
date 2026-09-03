import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { describe, expect, it } from 'vitest'
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
    store.record('responses_stream_invalid_messages_sse')
    store.record('private prompt jwt_123 /Users/person/file.pptx')
    store.finishTask(id, 'failed', 'enhanced_turn_timeout')
    const task = store.recent()[0]!
    expect(task).toMatchObject({ diagnosticId: id, host: 'slides', status: 'failed' })
    expect(task.events.map((event) => event.code)).toContain('stream_protocol_rejected')
    expect(task.failureCode).toBe('stream_protocol_rejected')
    expect(JSON.stringify(task)).not.toContain('private prompt')
    expect(JSON.stringify(task)).not.toContain('/Users/')
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
