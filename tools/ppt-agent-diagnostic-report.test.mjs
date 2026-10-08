import assert from 'node:assert/strict'
import { spawnSync } from 'node:child_process'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { test } from 'node:test'
import { summarizePresentationDiagnostics } from './ppt-agent-diagnostic-report.mjs'

const event = {
  event_id: 'event-1',
  trace_id: 'trace-1',
  phase: 'tool',
  outcome: 'failed',
  error_code: 'office_write_failed',
  presentation_stage: 'production',
  presentation_context: {
    session_id: 'session-1',
    document_id: 'document-1',
    project_id: 'project-1',
    request_id: 'request-1',
    page_id: 'page-3',
    tool_call_id: 'call-1',
  },
}

test('joins a page failure with the exact Relay session, trace and event', () => {
  const relay = JSON.stringify({
    event: 'office_diagnostic',
    session_id: 'session-1',
    trace_id: 'trace-1',
    event_id: 'event-1',
  })
  const report = summarizePresentationDiagnostics(
    { version: 1, events: [{ ...event, private_body: 'sensitive slide text' }] },
    relay,
  )
  assert.equal(report.events, 1)
  assert.deepEqual(report.counts, { tool: { total: 1, failed: 1 } })
  assert.deepEqual(report.stages, { production: { total: 1, failed: 1 } })
  assert.deepEqual(report.failures, [
    {
      trace_id: 'trace-1',
      event_id: 'event-1',
      phase: 'tool',
      stage: 'production',
      error_code: 'office_write_failed',
      ...event.presentation_context,
      relay_observed: true,
    },
  ])
  assert.equal(JSON.stringify(report).includes('sensitive slide text'), false)
})

test('does not mistake a sampled or absent Relay event for a proven upload failure', () => {
  const wrongSession = JSON.stringify({
    event: 'office_diagnostic',
    session_id: 'session-2',
    trace_id: 'trace-1',
    event_id: 'event-1',
  })
  const report = summarizePresentationDiagnostics({ version: 1, events: [event] }, wrongSession)
  assert.equal(report.failures[0].relay_observed, false)
  assert.equal(report.failures[0].page_id, 'page-3')
})

test('counts completed events but reports only failures and unsupported operations', () => {
  const report = summarizePresentationDiagnostics({
    version: 1,
    events: [
      event,
      {
        ...event,
        event_id: 'event-2',
        phase: 'run',
        outcome: 'passed',
        error_code: 'agent_run_completed',
        presentation_context: { session_id: 'session-1' },
      },
      {
        ...event,
        event_id: 'event-3',
        phase: 'verify',
        outcome: 'unsupported',
        error_code: 'office_api_unsupported',
      },
    ],
  })
  assert.equal(report.failures.length, 2)
  assert.deepEqual(report.counts, {
    tool: { total: 1, failed: 1 },
    run: { total: 1, failed: 0 },
    verify: { total: 1, failed: 1 },
  })
  assert.deepEqual(report.stages, { production: { total: 3, failed: 2 } })
})

test('rejects malformed identities and duplicate event IDs', () => {
  assert.throws(
    () =>
      summarizePresentationDiagnostics({
        version: 1,
        events: [{ ...event, presentation_context: { page_id: 'private slide title' } }],
      }),
    /diagnostic_export_invalid/,
  )
  assert.throws(
    () => summarizePresentationDiagnostics({ version: 1, events: [event, event] }),
    /diagnostic_export_invalid/,
  )
  assert.throws(
    () => summarizePresentationDiagnostics({ version: 1, events: [event] }, '{'),
    /relay_log_invalid/,
  )
})

test('CLI reports malformed private input with a stable error only', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'ppt-diagnostic-'))
  try {
    const path = join(directory, 'private.json')
    await writeFile(path, '{private slide text')
    const result = spawnSync(process.execPath, ['tools/ppt-agent-diagnostic-report.mjs', path], {
      encoding: 'utf8',
    })
    assert.equal(result.status, 1)
    assert.equal(result.stderr, 'diagnostic_export_invalid\n')
    assert.equal(result.stdout, '')
  } finally {
    await rm(directory, { recursive: true, force: true })
  }
})
