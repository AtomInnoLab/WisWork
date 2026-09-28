import { readFile, stat } from 'node:fs/promises'
import { resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

const id = (value) => typeof value === 'string' && /^[A-Za-z0-9_#-]{1,128}$/.test(value)
const code = (value) => typeof value === 'string' && /^[a-z][a-z0-9_:]{0,127}$/.test(value)
const phases = new Set([
  'run',
  'tool',
  'proposal',
  'validate',
  'write',
  'verify',
  'recovery',
  'transport',
])
const stages = new Set([
  'project_recovery',
  'planning',
  'production',
  'import',
  'review',
  'evidence',
  'sources',
  'editing',
])

/** Produces a local-only failure index. Relay absence can mean sampling or unavailable logs. */
export function summarizePresentationDiagnostics(local, relayJsonl = '') {
  if (local?.version !== 1 || !Array.isArray(local.events) || local.events.length > 200)
    throw new Error('diagnostic_export_invalid')
  const relay = new Set()
  for (const line of relayJsonl.split(/\r?\n/)) {
    if (!line.trim()) continue
    if (!line.trimStart().startsWith('{')) continue
    let entry
    try {
      entry = JSON.parse(line)
    } catch {
      throw new Error('relay_log_invalid')
    }
    if (entry?.event !== 'office_diagnostic') continue
    if (![entry.session_id, entry.trace_id, entry.event_id].every(id))
      throw new Error('relay_log_invalid')
    relay.add(`${entry.session_id}/${entry.trace_id}/${entry.event_id}`)
  }
  const seen = new Set()
  const failures = []
  const counts = {}
  const stageCounts = {}
  for (const event of local.events) {
    if (
      !id(event?.event_id) ||
      !id(event.trace_id) ||
      !phases.has(event.phase) ||
      !['passed', 'failed', 'unsupported', 'cancelled'].includes(event.outcome) ||
      !code(event.error_code) ||
      (event.presentation_stage !== undefined && !stages.has(event.presentation_stage))
    )
      throw new Error('diagnostic_export_invalid')
    if (seen.has(event.event_id)) throw new Error('diagnostic_export_invalid')
    seen.add(event.event_id)
    const context = event.presentation_context ?? {}
    if (
      Object.entries(context).some(
        ([key, value]) =>
          ![
            'session_id',
            'document_id',
            'project_id',
            'request_id',
            'page_id',
            'tool_call_id',
          ].includes(key) || !id(value),
      )
    )
      throw new Error('diagnostic_export_invalid')
    counts[event.phase] ??= { total: 0, failed: 0 }
    counts[event.phase].total += 1
    const stage = event.presentation_stage ?? 'unclassified'
    stageCounts[stage] ??= { total: 0, failed: 0 }
    stageCounts[stage].total += 1
    if (event.outcome !== 'failed' && event.outcome !== 'unsupported') continue
    counts[event.phase].failed += 1
    stageCounts[stage].failed += 1
    failures.push({
      trace_id: event.trace_id,
      event_id: event.event_id,
      phase: event.phase,
      ...(event.presentation_stage ? { stage: event.presentation_stage } : {}),
      error_code: event.error_code,
      ...Object.fromEntries(
        Object.entries(context).filter(([key]) =>
          [
            'session_id',
            'document_id',
            'project_id',
            'request_id',
            'page_id',
            'tool_call_id',
          ].includes(key),
        ),
      ),
      relay_observed: context.session_id
        ? relay.has(`${context.session_id}/${event.trace_id}/${event.event_id}`)
        : false,
    })
  }
  return { events: local.events.length, failures, counts, stages: stageCounts }
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const [localPath, relayPath] = process.argv.slice(2)
  if (!localPath || process.argv.length > 4) {
    process.stderr.write(
      'Usage: node tools/ppt-agent-diagnostic-report.mjs <office-diagnostics.json> [relay.jsonl]\n',
    )
    process.exitCode = 2
  } else {
    try {
      const localFile = resolve(localPath),
        relayFile = relayPath ? resolve(relayPath) : undefined
      if (
        (await stat(localFile)).size > 256 * 1024 ||
        (relayFile && (await stat(relayFile)).size > 16 * 1024 * 1024)
      )
        throw new Error('diagnostic_input_too_large')
      const localText = await readFile(localFile, 'utf8')
      const relayText = relayFile ? await readFile(relayFile, 'utf8') : ''
      let local
      try {
        local = JSON.parse(localText)
      } catch {
        throw new Error('diagnostic_export_invalid')
      }
      const report = summarizePresentationDiagnostics(local, relayText)
      process.stdout.write(`${JSON.stringify(report, null, 2)}\n`)
    } catch (error) {
      const safe = new Set([
        'diagnostic_export_invalid',
        'relay_log_invalid',
        'diagnostic_input_too_large',
      ])
      process.stderr.write(
        `${error instanceof Error && safe.has(error.message) ? error.message : 'diagnostic_report_failed'}\n`,
      )
      process.exitCode = 1
    }
  }
}
