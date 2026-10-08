import { mkdtempSync, rmSync, readFileSync, writeFileSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, expect, it } from 'vitest'
import { PresentationStore } from '../src/presentation-store.js'
import { benchmarkPlan } from '../../pptx-engine/tests/fixtures/presentation-plan.js'
import { parsePresentationPlan } from '../../pptx-engine/src/presentation-plan.js'

const roots: string[] = []
afterEach(() => roots.splice(0).forEach((root) => rmSync(root, { recursive: true, force: true })))
it('keeps opaque legacy JSON readable when it has no structured page lock contract', () => {
  const root = mkdtempSync(join(tmpdir(), 'ppt-opaque-plan-'))
  roots.push(root)
  const store = new PresentationStore(root)
  store.savePlan('opaque', 'doc', 0, { slides: 'legacy text', title: 'A' })
  expect(store.plan('opaque', 'doc')?.revision).toBe(1)
  expect(store.savePlan('opaque', 'doc', 1, { slides: 'legacy text', title: 'B' }).revision).toBe(2)
})
it('refuses missing or damaged lock anchors rather than silently selecting another input', () => {
  const root = mkdtempSync(join(tmpdir(), 'ppt-lock-anchor-'))
  roots.push(root)
  const store = new PresentationStore(root)
  const plan = benchmarkPlan()
  store.savePlan(plan.projectId, 'doc', 0, plan)
  store.setPlanPageLock(plan.projectId, 'doc', 1, plan.slides[1]!.id, true)
  const path = join(
    root,
    'projects',
    'presentations',
    createHash('sha256').update(plan.projectId).digest('hex'),
    'plan.json',
  )
  const saved = JSON.parse(readFileSync(path, 'utf8'))
  const { pageLocks: _locks, pageLocksDigest: _digest, ...withoutLocks } = saved
  writeFileSync(path, JSON.stringify(withoutLocks))
  expect(() => store.plan(plan.projectId, 'doc')).toThrow('invalid_state')
  writeFileSync(
    path,
    JSON.stringify({
      ...saved,
      pageLocks: [{ pageId: plan.slides[1]!.id, productionInputDigest: 'f'.repeat(64) }],
    }),
  )
  expect(() => store.plan(plan.projectId, 'doc')).toThrow('invalid_state')
})
it('persists locks, rejects indirect changes and requires an explicit revision-bound unlock', () => {
  const root = mkdtempSync(join(tmpdir(), 'ppt-plan-lock-'))
  roots.push(root)
  let store = new PresentationStore(root)
  const plan = benchmarkPlan()
  store.savePlan(plan.projectId, 'doc', 0, plan)
  const pageId = plan.slides[1]!.id
  const locked = store.setPlanPageLock(plan.projectId, 'doc', 1, pageId, true)
  expect(parsePresentationPlan(locked.plan).slides[1]).toMatchObject({ locked: true })
  store = new PresentationStore(root)
  expect(store.plan(plan.projectId, 'doc')?.revision).toBe(2)
  for (const change of ['unlock', 'delete', 'purpose', 'style', 'claim', 'source', 'move']) {
    const next = structuredClone(locked.plan) as typeof plan
    if (change === 'unlock') delete (next.slides[1] as { locked?: boolean }).locked
    if (change === 'delete') next.slides.splice(1, 1)
    if (change === 'purpose') next.slides[1]!.purpose += ' changed'
    if (change === 'style') next.style.accentColor = 'FFFFFF'
    if (change === 'claim') next.claims[0]!.statement += ' changed'
    if (change === 'source') next.sources[0]!.excerpt += ' changed'
    if (change === 'move') [next.slides[1], next.slides[2]] = [next.slides[2]!, next.slides[1]!]
    expect(() => store.savePlan(plan.projectId, 'doc', 2, next), change).toThrow('page_locked')
    expect(store.plan(plan.projectId, 'doc')?.revision).toBe(2)
  }
  const unrelated = structuredClone(locked.plan) as typeof plan
  unrelated.slides[7]!.purpose += ' unrelated'
  expect(store.savePlan(plan.projectId, 'doc', 2, unrelated).revision).toBe(3)
  expect(() => store.setPlanPageLock(plan.projectId, 'doc', 2, pageId, false)).toThrow(
    'revision_conflict',
  )
  expect(() => store.setPlanPageLock(plan.projectId, 'foreign', 3, pageId, false)).toThrow(
    'document_mismatch',
  )
  expect(store.setPlanPageLock(plan.projectId, 'doc', 3, pageId, false).revision).toBe(4)
  expect(store.setPlanPageLock(plan.projectId, 'doc', 4, pageId, false).revision).toBe(4)
  expect(() => store.setPlanPageLock(plan.projectId, 'doc', 4, 'missing', true)).toThrow(
    'invalid_request',
  )
})
