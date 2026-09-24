import { readdir, readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

export const CASE_IDS = Array.from({ length: 20 }, (_, index) =>
  `PPT-P0-${String(index + 1).padStart(2, '0')}`,
)
const digest = (value) => typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
const nonempty = (value) => typeof value === 'string' && value.trim().length > 0

export function summarizePresentationAcceptance(records) {
  if (!Array.isArray(records)) throw new Error('acceptance_records_invalid')
  const byCase = new Map(CASE_IDS.map((id) => [id, []]))
  const attempts = new Set()
  for (const record of records) {
    if (!record || typeof record !== 'object' || Array.isArray(record))
      throw new Error('acceptance_record_invalid')
    const entries = byCase.get(record.case_id)
    if (!entries || !nonempty(record.attempt_id) || !Number.isSafeInteger(record.attempt_no) || record.attempt_no < 1)
      throw new Error('acceptance_record_invalid')
    const key = `${record.case_id}/${record.attempt_id}`
    if (attempts.has(key) || entries.some((entry) => entry.attempt_no === record.attempt_no))
      throw new Error('acceptance_duplicate_attempt')
    attempts.add(key)
    if (!['passed', 'failed', 'blocked'].includes(record.outcome))
      throw new Error('acceptance_outcome_invalid')
    if (record.outcome === 'passed') {
      if (
        record.material_status !== 'ready' ||
        !nonempty(record.material_manifest) ||
        !nonempty(record.commit_and_versions) ||
        !nonempty(record.identity) ||
        !nonempty(record.reviewer_and_date) ||
        !digest(record.artifacts?.pptx_sha256) ||
        record.artifacts?.powerpoint_reopened !== true ||
        record.artifacts?.editable_after_reopen !== true ||
        record.restart_required !== false ||
        record.p0_defects !== 0
      ) throw new Error('acceptance_pass_evidence_missing')
    }
    entries.push(record)
  }
  let executed = 0
  let passed = 0
  let blocked = 0
  const cases = CASE_IDS.map((id) => {
    const entries = byCase.get(id).sort((a, b) => a.attempt_no - b.attempt_no)
    const latest = entries.at(-1)
    if (latest) executed += 1
    if (latest?.outcome === 'passed') passed += 1
    if (latest?.outcome === 'blocked') blocked += 1
    return { id, attempts: entries.length, status: latest?.outcome ?? 'not_run' }
  })
  return {
    denominator: 20,
    attempted: executed,
    passed,
    failed: executed - passed - blocked,
    blocked,
    notRun: 20 - executed,
    attempts: records.length,
    completionRate: executed === 20 ? `${Math.round((passed / 20) * 100)}%` : 'not_measured',
    rateThresholdMet: executed === 20 && passed >= 16,
    cases,
  }
}

export async function readPresentationAcceptance(directory) {
  const names = (await readdir(directory)).filter((name) => name.endsWith('.json')).sort()
  const records = []
  for (const name of names) {
    const parsed = JSON.parse(await readFile(resolve(directory, name), 'utf8'))
    if (!Array.isArray(parsed)) throw new Error(`acceptance_file_invalid:${name}`)
    records.push(...parsed)
  }
  return summarizePresentationAcceptance(records)
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const directory = process.argv[2]
  if (!directory) {
    process.stderr.write('Usage: node tools/ppt-agent-acceptance.mjs <records-directory>\n')
    process.exitCode = 2
  } else {
    readPresentationAcceptance(resolve(directory))
      .then((report) => process.stdout.write(`${JSON.stringify(report, null, 2)}\n`))
      .catch((error) => {
        process.stderr.write(`${error instanceof Error ? error.message : 'acceptance_failed'}\n`)
        process.exitCode = 1
      })
  }
}
