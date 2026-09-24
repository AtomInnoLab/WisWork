import assert from 'node:assert/strict'
import { test } from 'node:test'
import { CASE_IDS, summarizePresentationAcceptance } from './ppt-agent-acceptance.mjs'

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
})

test('empty records remain unmeasured against the fixed 20-case denominator', () => {
  const report = summarizePresentationAcceptance([])
  assert.deepEqual([report.denominator, report.attempted, report.passed, report.notRun], [20, 0, 0, 20])
  assert.equal(report.completionRate, 'not_measured')
  assert.equal(report.rateThresholdMet, false)
})

test('keeps every attempt and counts only the latest outcome for each fixed case', () => {
  const initialFailure = {
    case_id: CASE_IDS[0], attempt_id: 'attempt-1', attempt_no: 1, outcome: 'failed',
  }
  const records = [initialFailure, passed(CASE_IDS[0], 2), ...CASE_IDS.slice(1).map((id) => passed(id))]
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
      case_id: id, attempt_id: 'attempt-1', attempt_no: 1, outcome: 'failed',
    })),
  ])
  assert.equal(complete.completionRate, '80%')
  assert.equal(complete.rateThresholdMet, true)
})

test('refuses unsupported cases, duplicate attempts and unsupported success claims', () => {
  assert.throws(() => summarizePresentationAcceptance([{ ...passed('PPT-P0-21') }]), /record_invalid/)
  assert.throws(() => summarizePresentationAcceptance([passed(CASE_IDS[0]), passed(CASE_IDS[0])]), /duplicate_attempt/)
  assert.throws(() => summarizePresentationAcceptance([{ ...passed(CASE_IDS[0]), restart_required: true }]), /pass_evidence_missing/)
  assert.throws(() => summarizePresentationAcceptance([{ ...passed(CASE_IDS[0]), artifacts: { pptx_sha256: 'a'.repeat(64) } }]), /pass_evidence_missing/)
})
