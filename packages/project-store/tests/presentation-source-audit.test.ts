import { mkdtempSync, rmSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { afterEach, expect, it } from 'vitest'
import { PresentationStore } from '../src/presentation-store.js'
import { benchmarkPlan } from '../../pptx-engine/tests/fixtures/presentation-plan.js'
import {
  parsePresentationSourceAuditHistory,
  presentationSourceAuditHistory,
} from '../src/presentation-source-audit.js'
const roots: string[] = []
afterEach(() => roots.splice(0).forEach((root) => rmSync(root, { recursive: true, force: true })))
function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'source-audit-ledger-'))
  roots.push(root)
  const store = new PresentationStore(root),
    plan = benchmarkPlan()
  plan.sources[0]!.uri = `attachment:${'a'.repeat(64)}`
  store.savePlan(plan.projectId, 'doc', 0, plan)
  return { root, store, plan }
}
it('persists exact research scope and results, reopens detached copies and preserves completed request identity across plan changes', () => {
  const { root, store, plan } = fixture()
  const started = store.beginSourceAudit(plan.projectId, 'doc', 'audit')
  expect(started).toMatchObject({
    scope: 'source_excerpt_audit',
    planRevision: 1,
    state: 'running',
    sourceRefs: [{ sourceId: plan.sources[0]!.id, attachmentId: 'a'.repeat(64) }],
  })
  expect(store.beginSourceAudit(plan.projectId, 'doc', 'audit')).toEqual(started)
  const completed = store.finishSourceAudit(plan.projectId, 'doc', 'audit', {
    sources: [{ ...started.sourceRefs[0]!, status: 'found', offset: 0, locator: '第 1 段' }],
  })
  store.savePlan(plan.projectId, 'doc', 1, { ...plan, title: 'new plan' })
  expect(new PresentationStore(root).sourceAudit(plan.projectId, 'doc', 'audit')).toEqual(completed)
  expect(store.beginSourceAudit(plan.projectId, 'doc', 'audit')).toEqual(completed)
  expect(() =>
    store.finishSourceAudit(plan.projectId, 'doc', 'audit', {
      sources: [{ ...started.sourceRefs[0]!, status: 'missing' }],
    }),
  ).toThrow('request_conflict')
  const ledger = store.sourceAudits(plan.projectId, 'doc')
  const summary = parsePresentationSourceAuditHistory(presentationSourceAuditHistory(ledger))
  expect(summary.runs[0]).toMatchObject({ state: 'completed', sourceCount: 1, foundCount: 1 })
  expect(summary.runs[0]).not.toHaveProperty('sources')
  completed.sources![0]!.offset = 44
  expect(store.sourceAudit(plan.projectId, 'doc', 'audit')!.sources![0]!.offset).toBe(0)
  expect(() => store.sourceAudits(plan.projectId, 'foreign')).toThrow('document_mismatch')
})
it('rejects corrupt ledger data and mismatched completion scope; unfinished audit remains explicit and bounded history keeps its identity', () => {
  const { root, store, plan } = fixture()
  store.beginSourceAudit(plan.projectId, 'doc', 'pending')
  expect(() => store.finishSourceAudit(plan.projectId, 'doc', 'pending', { sources: [] })).toThrow(
    'invalid_state',
  )
  for (let index = 0; index < 36; index++) {
    store.beginSourceAudit(plan.projectId, 'doc', `audit-${index}`)
    store.finishSourceAudit(plan.projectId, 'doc', `audit-${index}`, { error: 'aborted' })
  }
  const ledger = store.sourceAudits(plan.projectId, 'doc')
  expect(ledger.runs).toHaveLength(32)
  expect(ledger.runs.some((run) => run.id === 'pending' && run.state === 'running')).toBe(true)
  const archived = new PresentationStore(root).sourceAudit(plan.projectId, 'doc', 'audit-0')
  expect(archived).toMatchObject({ id: 'audit-0', state: 'failed', error: 'aborted' })
  expect(store.beginSourceAudit(plan.projectId, 'doc', 'audit-0')).toEqual(archived)
  expect(store.finishSourceAudit(plan.projectId, 'doc', 'audit-0', { error: 'aborted' })).toEqual(
    archived,
  )
  const path = join(
    root,
    'projects',
    'presentations',
    createHash('sha256').update(plan.projectId).digest('hex'),
    'source-audits.json',
  )
  const value = JSON.parse(readFileSync(path, 'utf8'))
  value.runs[1].error = '/private/path secret'
  writeFileSync(path, JSON.stringify(value))
  expect(() => store.sourceAudits(plan.projectId, 'doc')).toThrow('invalid_state')
})
it('rejects public history with result payloads, duplicate identities, unsafe error text or impossible counts', () => {
  const { store, plan } = fixture()
  store.beginSourceAudit(plan.projectId, 'doc', 'audit')
  store.finishSourceAudit(plan.projectId, 'doc', 'audit', { error: 'aborted' })
  const history = presentationSourceAuditHistory(store.sourceAudits(plan.projectId, 'doc'))
  for (const patch of [
    { sources: [] },
    { error: '/private/path' },
    { sourceCount: 257 },
    { sequence: history.revision + 1 },
  ]) {
    expect(() =>
      parsePresentationSourceAuditHistory({ ...history, runs: [{ ...history.runs[0], ...patch }] }),
    ).toThrow('invalid_state')
  }
  expect(() =>
    parsePresentationSourceAuditHistory({ ...history, runs: [history.runs[0], history.runs[0]] }),
  ).toThrow('invalid_state')
})
