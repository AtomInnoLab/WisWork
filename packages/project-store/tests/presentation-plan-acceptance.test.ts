import { mkdtempSync, rmSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { expect, it } from 'vitest'
import { PresentationStore } from '../src/presentation-store.js'
import { benchmarkPlan } from '../../pptx-engine/tests/fixtures/presentation-plan.js'
import { parsePresentationPlanAcceptances } from '../src/presentation-plan-acceptance.js'

it('accepts an exact plan on explicit decision, preserves identity after reopen and rejects stale or conflicting decisions', () => {
  const root = mkdtempSync(join(tmpdir(), 'plan-acceptance-'))
  try {
    const store = new PresentationStore(root),
      plan = benchmarkPlan()
    const saved = store.savePlan(plan.projectId, 'doc', 0, plan)
    expect(store.planAcceptances(plan.projectId, 'doc').records).toEqual([])
    const accepted = store.acceptPlan(plan.projectId, 'doc', 'decision', 1, saved.inputDigest)
    expect(accepted).toMatchObject({
      decisionId: 'decision',
      planRevision: 1,
      planDigest: saved.inputDigest,
    })
    expect(new PresentationStore(root).planAcceptances(plan.projectId, 'doc').records).toEqual([
      accepted,
    ])
    store.savePlan(plan.projectId, 'doc', 1, { ...plan, title: 'revised' })
    expect(store.acceptPlan(plan.projectId, 'doc', 'decision', 1, saved.inputDigest)).toEqual(
      accepted,
    )
    expect(() => store.acceptPlan(plan.projectId, 'doc', 'new', 1, saved.inputDigest)).toThrow(
      'revision_conflict',
    )
    expect(() => store.acceptPlan(plan.projectId, 'doc', 'decision', 2, saved.inputDigest)).toThrow(
      'request_conflict',
    )
    expect(() => store.planAcceptances(plan.projectId, 'foreign')).toThrow('document_mismatch')
    const path = join(
      root,
      'projects',
      'presentations',
      createHash('sha256').update(plan.projectId).digest('hex'),
      'plan-acceptances.json',
    )
    const raw = JSON.parse(readFileSync(path, 'utf8'))
    raw.records[0].planDigest = 'f'.repeat(64)
    writeFileSync(path, JSON.stringify(raw))
    expect(() => store.planAcceptances(plan.projectId, 'doc')).toThrow('invalid_state')
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})
it('keeps bounded decisions, rejects malformed public metadata and replays identity even at capacity', () => {
  const root = mkdtempSync(join(tmpdir(), 'plan-acceptance-capacity-'))
  try {
    const store = new PresentationStore(root),
      plan = benchmarkPlan(),
      saved = store.savePlan(plan.projectId, 'doc', 0, plan)
    for (let index = 0; index < 64; index++)
      store.acceptPlan(plan.projectId, 'doc', `decision-${index}`, 1, saved.inputDigest)
    expect(() => store.acceptPlan(plan.projectId, 'doc', 'overflow', 1, saved.inputDigest)).toThrow(
      'acceptance_capacity',
    )
    expect(
      store.acceptPlan(plan.projectId, 'doc', 'decision-0', 1, saved.inputDigest).decisionId,
    ).toBe('decision-0')
    const ledger = store.planAcceptances(plan.projectId, 'doc')
    expect(() =>
      parsePresentationPlanAcceptances({
        ...ledger,
        records: [ledger.records[0], ledger.records[0]],
      }),
    ).toThrow('invalid_state')
    expect(() =>
      parsePresentationPlanAcceptances({
        ...ledger,
        records: [{ ...ledger.records[0], acceptedAt: 'private error' }],
      }),
    ).toThrow('invalid_state')
    expect(() =>
      parsePresentationPlanAcceptances({
        ...ledger,
        records: [{ ...ledger.records[0], rawPlan: plan }],
      }),
    ).toThrow('invalid_state')
    ledger.records[0]!.planDigest = 'f'.repeat(64)
    expect(store.planAcceptances(plan.projectId, 'doc').records[0]!.planDigest).toBe(
      saved.inputDigest,
    )
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})
