import assert from 'node:assert/strict'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { test } from 'node:test'
import {
  CASE_IDS,
  readPresentationAcceptance,
  summarizePresentationAcceptance,
} from './ppt-agent-acceptance.mjs'

const passed = (caseId, attemptNo = 1) => ({
  case_id: caseId,
  attempt_id: `attempt-${attemptNo}`,
  attempt_no: attemptNo,
  outcome: 'passed',
  material_status: 'ready',
  material_manifest: 'Authorized source and SHA256 recorded',
  commit_and_versions: 'PC, Add-in and PowerPoint versions recorded',
  identity: 'Project, request, document and session IDs recorded',
  reviewer_and_date: 'Reviewer 2026-09-25',
  artifacts: {
    pptx_sha256: 'a'.repeat(64),
    powerpoint_reopened: true,
    editable_after_reopen: true,
  },
  restart_required: false,
  p0_defects: 0,
  measurements: { confidentiality_violations: 0, cross_document_writes: 0 },
})

test('empty records remain unmeasured against the fixed 20-case denominator', () => {
  const report = summarizePresentationAcceptance([])
  assert.deepEqual(
    [report.denominator, report.attempted, report.passed, report.notRun],
    [20, 0, 0, 20],
  )
  assert.equal(report.completionRate, 'not_measured')
  assert.equal(report.rateThresholdMet, false)
})

test('keeps every attempt and counts only the latest outcome for each fixed case', () => {
  const initialFailure = {
    case_id: CASE_IDS[0],
    attempt_id: 'attempt-1',
    attempt_no: 1,
    outcome: 'failed',
  }
  const records = [
    initialFailure,
    passed(CASE_IDS[0], 2),
    ...CASE_IDS.slice(1).map((id) => passed(id)),
  ]
  const report = summarizePresentationAcceptance(records)
  assert.equal(report.attempts, 21)
  assert.equal(report.passed, 20)
  assert.equal(report.completionRate, '100%')
  assert.equal(report.rateThresholdMet, true)
  assert.deepEqual(report.cases[0], { id: CASE_IDS[0], attempts: 2, status: 'passed' })
})

test('requires all 20 cases to be attempted before the 16-case P0 gate can pass', () => {
  const incomplete = summarizePresentationAcceptance(CASE_IDS.slice(0, 16).map((id) => passed(id)))
  assert.equal(incomplete.passed, 16)
  assert.equal(incomplete.completionRate, 'not_measured')
  assert.equal(incomplete.rateThresholdMet, false)
  const complete = summarizePresentationAcceptance([
    ...CASE_IDS.slice(0, 16).map((id) => passed(id)),
    ...CASE_IDS.slice(16).map((id) => ({
      case_id: id,
      attempt_id: 'attempt-1',
      attempt_no: 1,
      outcome: 'failed',
    })),
  ])
  assert.equal(complete.completionRate, '80%')
  assert.equal(complete.rateThresholdMet, true)
})

test('refuses unsupported cases, duplicate attempts and unsupported success claims', () => {
  assert.throws(
    () => summarizePresentationAcceptance([{ ...passed('PPT-P0-21') }]),
    /record_invalid/,
  )
  assert.throws(
    () => summarizePresentationAcceptance([passed(CASE_IDS[0]), passed(CASE_IDS[0])]),
    /duplicate_attempt/,
  )
  assert.throws(
    () => summarizePresentationAcceptance([{ ...passed(CASE_IDS[0]), restart_required: true }]),
    /pass_evidence_missing/,
  )
  assert.throws(
    () =>
      summarizePresentationAcceptance([
        { ...passed(CASE_IDS[0]), artifacts: { pptx_sha256: 'a'.repeat(64) } },
      ]),
    /pass_evidence_missing/,
  )
})

test('rejects a missing earlier attempt rather than silently replacing its outcome', () => {
  assert.throws(
    () => summarizePresentationAcceptance([passed(CASE_IDS[0], 2)]),
    /acceptance_attempt_gap/,
  )
  assert.throws(
    () =>
      summarizePresentationAcceptance([
        { case_id: CASE_IDS[0], attempt_id: 'attempt-1', attempt_no: 1, outcome: 'failed' },
        passed(CASE_IDS[0], 3),
      ]),
    /acceptance_attempt_gap/,
  )
})

test('reads JSON attempt arrays from a records directory', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'ppt-acceptance-'))
  try {
    await writeFile(
      join(directory, '01.json'),
      JSON.stringify([{ case_id: CASE_IDS[0], attempt_id: 'a', attempt_no: 1, outcome: 'failed' }]),
    )
    const report = await readPresentationAcceptance(directory)
    assert.equal(report.attempted, 1)
    assert.equal(report.failed, 1)
    assert.equal(report.completionRate, 'not_measured')
  } finally {
    await rm(directory, { recursive: true, force: true })
  }
})

test('reports latest-attempt quality measurements only when all 20 cases have evidence', () => {
  const measurements = (minutes, corrections = 0) => ({
    started_at: '2026-09-25T00:00:00.000Z',
    first_real_page_at: '2026-09-25T00:01:00.000Z',
    finished_at: `2026-09-25T00:${String(minutes).padStart(2, '0')}:00.000Z`,
    manual_correction_pages: corrections,
    duplicate_writes: 0,
    screenshot_failures: 0,
    image_failures: 0,
    confidentiality_violations: 0,
    cross_document_writes: 0,
  })
  const partial = summarizePresentationAcceptance([
    { ...passed(CASE_IDS[0]), measurements: measurements(3, 2) },
  ])
  assert.deepEqual(partial.measurements.manual_correction_pages, {
    observed: 1,
    total: 'not_measured',
  })
  assert.equal(partial.measurements.durations.p95_ms, 'not_measured')
  assert.equal(partial.measurements.first_page_latency.p95_ms, 'not_measured')
  const records = CASE_IDS.map((id, index) => ({
    ...passed(id),
    measurements: measurements(index + 2, index === 0 ? 2 : 0),
  }))
  records.push({ ...passed(CASE_IDS[0], 2), measurements: measurements(22, 1) })
  const complete = summarizePresentationAcceptance(records)
  assert.equal(complete.measurements.manual_correction_pages.total, 1)
  assert.equal(complete.measurements.duplicate_writes.total, 0)
  assert.equal(complete.measurements.durations.observed, 20)
  assert.equal(complete.measurements.durations.p95_ms, 21 * 60_000)
  assert.deepEqual(complete.measurements.first_page_latency, { observed: 20, p95_ms: 60_000 })
  const missing = summarizePresentationAcceptance(
    records.map((record) =>
      record.case_id === CASE_IDS[1]
        ? { ...record, measurements: { confidentiality_violations: 0, cross_document_writes: 0 } }
        : record,
    ),
  )
  assert.equal(missing.measurements.manual_correction_pages.total, 'not_measured')
  assert.equal(missing.measurements.durations.p95_ms, 'not_measured')
  assert.equal(missing.measurements.first_page_latency.p95_ms, 'not_measured')
})

test('rejects impossible or unbounded measurement records', () => {
  for (const measurements of [
    { started_at: '2026-09-25T00:00:00.000Z', finished_at: '2026-09-24T00:00:00.000Z' },
    { first_real_page_at: '2026-09-25T00:00:00.000Z' },
    { started_at: '2026-09-25', manual_correction_pages: 0 },
    { duplicate_writes: -1 },
    { screenshot_failures: 0.5 },
    { screenshot_failures: Number.MAX_SAFE_INTEGER },
    { image_failures: 0, extra: true },
  ])
    assert.throws(
      () => summarizePresentationAcceptance([{ ...passed(CASE_IDS[0]), measurements }]),
      /measurements_invalid/,
    )
})

test('aggregates professional quality numerators and denominators only from latest complete attempts', () => {
  const ratio = (numerator, denominator) => ({ numerator, denominator })
  const records = CASE_IDS.map((id) => ({
    ...passed(id),
    measurements: {
      ratios: {
        native_editable_objects: ratio(9, 10),
        critical_claims_traced: ratio(2, 2),
        citations_accurate: ratio(0, 0),
      },
      confidentiality_violations: 0,
      cross_document_writes: 0,
    },
  }))
  records.push({
    ...passed(CASE_IDS[0], 2),
    measurements: {
      ratios: {
        native_editable_objects: ratio(8, 10),
        critical_claims_traced: ratio(1, 2),
        citations_accurate: ratio(0, 0),
      },
      confidentiality_violations: 0,
      cross_document_writes: 0,
    },
  })
  const report = summarizePresentationAcceptance(records)
  assert.deepEqual(report.measurements.ratios.native_editable_objects, {
    observed: 20,
    numerator: 179,
    denominator: 200,
    rate: '89.5%',
  })
  assert.deepEqual(report.measurements.ratios.critical_claims_traced, {
    observed: 20,
    numerator: 39,
    denominator: 40,
    rate: '97.5%',
  })
  assert.deepEqual(report.measurements.ratios.citations_accurate, {
    observed: 20,
    numerator: 0,
    denominator: 0,
    rate: 'not_measured',
  })
  assert.deepEqual(report.measurements.ratios.reproducible_calculations, {
    observed: 0,
    numerator: 'not_measured',
    denominator: 'not_measured',
    rate: 'not_measured',
  })
  assert.equal(report.measurements.confidentiality_violations.total, 0)
  assert.equal(report.measurements.cross_document_writes.total, 0)
  const missing = summarizePresentationAcceptance([
    ...records,
    {
      ...passed(CASE_IDS[1], 2),
      measurements: { confidentiality_violations: 0, cross_document_writes: 0 },
    },
  ])
  assert.equal(missing.measurements.ratios.native_editable_objects.observed, 19)
  assert.equal(missing.measurements.ratios.native_editable_objects.rate, 'not_measured')
})

test('rejects invalid quality counts and forbids passed tasks with boundary violations', () => {
  const invalid = [
    { native_editable_objects: { numerator: 2, denominator: 1 } },
    { native_editable_objects: { numerator: -1, denominator: 1 } },
    { native_editable_objects: { numerator: 0, denominator: 1.5 } },
    { native_editable_objects: { numerator: 0, denominator: 1, extra: 1 } },
    { unknown_metric: { numerator: 0, denominator: 1 } },
  ]
  for (const ratios of invalid)
    assert.throws(
      () => summarizePresentationAcceptance([{ ...passed(CASE_IDS[0]), measurements: { ratios } }]),
      /measurements_invalid/,
    )
  for (const key of ['confidentiality_violations', 'cross_document_writes'])
    assert.throws(
      () =>
        summarizePresentationAcceptance([{ ...passed(CASE_IDS[0]), measurements: { [key]: 1 } }]),
      /pass_evidence_missing/,
    )
  assert.throws(
    () => summarizePresentationAcceptance([{ ...passed(CASE_IDS[0]), measurements: undefined }]),
    /pass_evidence_missing/,
  )
})
