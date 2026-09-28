import { createHash } from 'node:crypto'
import { createReadStream } from 'node:fs'
import { readdir, readFile, realpath, stat } from 'node:fs/promises'
import { isAbsolute, relative, resolve, sep } from 'node:path'
import { pathToFileURL } from 'node:url'

export const CASE_IDS = Array.from(
  { length: 20 },
  (_, index) => `PPT-P0-${String(index + 1).padStart(2, '0')}`,
)
const digest = (value) => typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
const nonempty = (value) => typeof value === 'string' && value.trim().length > 0
const nonnegative = (value) => Number.isSafeInteger(value) && value >= 0
const RATIO_KEYS = [
  'first_two_pages_style_revisions',
  'first_round_visual_passes',
  'native_editable_objects',
  'critical_facts_sourced',
  'critical_claims_traced',
  'citations_accurate',
  'unsupported_factual_claims',
  'timely_numeric_claims',
  'reproducible_calculations',
  'successful_recoveries',
  'prepared_images',
  'user_interruptions',
  'taskpane_recoveries',
  'pairing_first_try',
  'manual_changes_preserved',
]
const COUNTER_KEYS = [
  'manual_correction_pages',
  'duplicate_writes',
  'screenshot_failures',
  'image_failures',
  'confidentiality_violations',
  'cross_document_writes',
]
const timestamp = (value) =>
  typeof value === 'string' &&
  /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(value) &&
  !Number.isNaN(Date.parse(value)) &&
  new Date(value).toISOString() === value

function validateMeasurements(value) {
  if (value === undefined) return
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new Error('acceptance_measurements_invalid')
  const keys = ['started_at', 'first_real_page_at', 'finished_at', ...COUNTER_KEYS, 'ratios']
  if (Object.keys(value).some((key) => !keys.includes(key)))
    throw new Error('acceptance_measurements_invalid')
  for (const key of keys.slice(0, 3))
    if (value[key] !== undefined && !timestamp(value[key]))
      throw new Error('acceptance_measurements_invalid')
  for (const key of COUNTER_KEYS)
    if (value[key] !== undefined && (!nonnegative(value[key]) || value[key] > 1_000_000))
      throw new Error('acceptance_measurements_invalid')
  if (value.ratios !== undefined) {
    if (!value.ratios || typeof value.ratios !== 'object' || Array.isArray(value.ratios))
      throw new Error('acceptance_measurements_invalid')
    for (const [key, ratio] of Object.entries(value.ratios))
      if (
        !RATIO_KEYS.includes(key) ||
        !ratio ||
        typeof ratio !== 'object' ||
        Array.isArray(ratio) ||
        Object.keys(ratio).sort().join(',') !== 'denominator,numerator' ||
        !nonnegative(ratio.numerator) ||
        !nonnegative(ratio.denominator) ||
        ratio.numerator > ratio.denominator ||
        ratio.denominator > 1_000_000
      )
        throw new Error('acceptance_measurements_invalid')
  }
  const { started_at: start, first_real_page_at: first, finished_at: finish } = value
  if ((first || finish) && !start) throw new Error('acceptance_measurements_invalid')
  if (first && first < start) throw new Error('acceptance_measurements_invalid')
  if (finish && (finish < start || (first && finish < first)))
    throw new Error('acceptance_measurements_invalid')
}

function measurementSummary(latest) {
  const completed = latest.filter(Boolean)
  const ready = completed.length === CASE_IDS.length
  const counters = {}
  for (const key of COUNTER_KEYS) {
    const values = completed.map((record) => record.measurements?.[key])
    counters[key] = {
      observed: values.filter((value) => value !== undefined).length,
      total:
        ready && values.every((value) => value !== undefined)
          ? values.reduce((sum, value) => sum + value, 0)
          : 'not_measured',
    }
  }
  const latency = (end) => {
    const values = completed
      .map((record) => record.measurements)
      .filter((value) => value?.started_at && value?.[end])
      .map((value) => Date.parse(value[end]) - Date.parse(value.started_at))
      .sort((a, b) => a - b)
    return {
      observed: values.length,
      p95_ms:
        ready && values.length === CASE_IDS.length
          ? values[Math.ceil(0.95 * values.length) - 1]
          : 'not_measured',
    }
  }
  const ratios = {}
  for (const key of RATIO_KEYS) {
    const values = completed.map((record) => record.measurements?.ratios?.[key])
    const observed = values.filter(Boolean).length
    const complete = ready && observed === CASE_IDS.length
    const numerator = complete
      ? values.reduce((sum, value) => sum + value.numerator, 0)
      : 'not_measured'
    const denominator = complete
      ? values.reduce((sum, value) => sum + value.denominator, 0)
      : 'not_measured'
    ratios[key] = {
      observed,
      numerator,
      denominator,
      rate:
        complete && denominator > 0
          ? `${Math.round((numerator / denominator) * 10_000) / 100}%`
          : 'not_measured',
    }
  }
  return {
    durations: latency('finished_at'),
    first_page_latency: latency('first_real_page_at'),
    ratios,
    ...counters,
  }
}

export function summarizePresentationAcceptance(records) {
  if (!Array.isArray(records)) throw new Error('acceptance_records_invalid')
  const byCase = new Map(CASE_IDS.map((id) => [id, []]))
  const attempts = new Set()
  for (const record of records) {
    if (!record || typeof record !== 'object' || Array.isArray(record))
      throw new Error('acceptance_record_invalid')
    const entries = byCase.get(record.case_id)
    if (
      !entries ||
      !nonempty(record.attempt_id) ||
      !Number.isSafeInteger(record.attempt_no) ||
      record.attempt_no < 1
    )
      throw new Error('acceptance_record_invalid')
    const key = `${record.case_id}/${record.attempt_id}`
    if (attempts.has(key) || entries.some((entry) => entry.attempt_no === record.attempt_no))
      throw new Error('acceptance_duplicate_attempt')
    attempts.add(key)
    if (!['passed', 'failed', 'blocked'].includes(record.outcome))
      throw new Error('acceptance_outcome_invalid')
    validateMeasurements(record.measurements)
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
        record.p0_defects !== 0 ||
        record.measurements?.confidentiality_violations !== 0 ||
        record.measurements?.cross_document_writes !== 0
      )
        throw new Error('acceptance_pass_evidence_missing')
    }
    entries.push(record)
  }
  let executed = 0
  let passed = 0
  let firstAttemptPassed = 0
  let blocked = 0
  const latestAttempts = []
  const cases = CASE_IDS.map((id) => {
    const entries = byCase.get(id).sort((a, b) => a.attempt_no - b.attempt_no)
    if (entries.some((entry, index) => entry.attempt_no !== index + 1))
      throw new Error('acceptance_attempt_gap')
    const latest = entries.at(-1)
    if (entries[0]?.outcome === 'passed') firstAttemptPassed += 1
    latestAttempts.push(latest)
    if (latest) executed += 1
    if (latest?.outcome === 'passed') passed += 1
    if (latest?.outcome === 'blocked') blocked += 1
    return { id, attempts: entries.length, status: latest?.outcome ?? 'not_run' }
  })
  return {
    denominator: 20,
    attempted: executed,
    passed,
    firstAttemptPassed,
    failed: executed - passed - blocked,
    blocked,
    notRun: 20 - executed,
    attempts: records.length,
    completionRate: executed === 20 ? `${Math.round((passed / 20) * 100)}%` : 'not_measured',
    firstAttemptDeliveryRate:
      executed === 20 ? `${Math.round((firstAttemptPassed / 20) * 100)}%` : 'not_measured',
    rateThresholdMet: executed === 20 && passed >= 16,
    measurements: measurementSummary(latestAttempts),
    cases,
  }
}

export async function readPresentationAcceptance(directory) {
  const root = await realpath(directory)
  const names = (await readdir(directory)).filter((name) => name.endsWith('.json')).sort()
  const records = []
  for (const name of names) {
    const path = await boundedFile(root, name, 1024 * 1024)
    const parsed = JSON.parse(await readFile(path, 'utf8'))
    if (!Array.isArray(parsed)) throw new Error(`acceptance_file_invalid:${name}`)
    records.push(...parsed)
  }
  const report = summarizePresentationAcceptance(records)
  for (const record of records) {
    if (record.outcome !== 'passed') continue
    await verifyArtifact(root, record.material_manifest, record.material_manifest_sha256)
    if (
      typeof record.artifacts?.pptx_file !== 'string' ||
      !record.artifacts.pptx_file.endsWith('.pptx')
    )
      throw new Error('acceptance_artifact_invalid:pptx_file')
    const pptxPath = await verifyArtifact(
      root,
      record.artifacts.pptx_file,
      record.artifacts.pptx_sha256,
    )
    const reopenPath = await verifyArtifact(
      root,
      record.artifacts.reopen_evidence_file,
      record.artifacts.reopen_evidence_sha256,
    )
    if (pptxPath === reopenPath) throw new Error('acceptance_artifact_invalid:reopen_evidence_file')
  }
  return report
}

async function boundedFile(root, name, maxBytes) {
  if (typeof name !== 'string' || !name.trim() || isAbsolute(name))
    throw new Error('acceptance_artifact_invalid:path')
  const path = resolve(root, name)
  let actual, details
  try {
    actual = await realpath(path)
    details = await stat(actual)
  } catch {
    throw new Error('acceptance_artifact_invalid:path')
  }
  const fromRoot = relative(root, actual)
  if (
    !fromRoot ||
    fromRoot === '..' ||
    fromRoot.startsWith(`..${sep}`) ||
    isAbsolute(fromRoot) ||
    !details.isFile() ||
    details.size < 1 ||
    details.size > maxBytes
  )
    throw new Error('acceptance_artifact_invalid:path')
  return actual
}

async function verifyArtifact(root, name, expected) {
  if (!digest(expected)) throw new Error('acceptance_artifact_invalid:digest')
  const path = await boundedFile(root, name, 100 * 1024 * 1024)
  const hash = createHash('sha256')
  let size = 0
  for await (const chunk of createReadStream(path)) {
    size += chunk.length
    if (size > 100 * 1024 * 1024) throw new Error('acceptance_artifact_invalid:size')
    hash.update(chunk)
  }
  if (hash.digest('hex') !== expected) throw new Error('acceptance_artifact_digest_mismatch')
  return path
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
