import { createHash } from 'node:crypto'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, expect, it } from 'vitest'
import { PresentationStore } from '../src/index.js'
import {
  parsePresentationIssueActionInput,
  parsePresentationIssueLedger,
} from '../src/presentation-issue.js'
let root: string
let store: PresentationStore
const action = {
  actionId: 'a',
  issueId: 'issue',
  issueDigest: 'a'.repeat(64),
  state: 'explained' as const,
  note: 'Explanation only',
}
beforeEach(() => {
  root = mkdtempSync(join(tmpdir(), 'issue-actions-'))
  store = new PresentationStore(root)
  store.beginProduction(
    'project',
    'doc',
    'run',
    { slides: [{ id: 'page' }] },
    { revision: 1, plan: { title: 'Frozen' } },
  )
})
afterEach(() => rmSync(root, { recursive: true, force: true }))
const append = (revision = 0, value = action) =>
  store.appendIssueAction('project', 'doc', 'run', revision, value)
it('preserves immutable history, CAS, idempotency before CAS, and restart', () => {
  expect(store.issueActions('project', 'doc', 'run').revision).toBe(0)
  const first = append()
  expect(append(100)).toEqual(first)
  expect(() => append(1, { ...action, note: 'different' })).toThrow('request_conflict')
  expect(() => append(0, { ...action, actionId: 'b' })).toThrow('revision_conflict')
  const reopened = append(1, { ...action, actionId: 'b', state: 'open' as typeof action.state })
  expect(reopened.actions.map((item) => item.state)).toEqual(['explained', 'open'])
  expect(new PresentationStore(root).issueActions('project', 'doc', 'run')).toEqual(reopened)
  first.actions[0]!.note = 'mutated'
  expect(store.issueActions('project', 'doc', 'run').actions[0]!.note).toBe(action.note)
})
it('bounds history, and permits identical retries at capacity', () => {
  for (let index = 0; index < 128; index++) append(index, { ...action, actionId: `a${index}` })
  expect(append(0, { ...action, actionId: 'a0' }).revision).toBe(128)
  const before = store.issueActions('project', 'doc', 'run')
  expect(() => append(128)).toThrow('quota_exceeded')
  expect(store.issueActions('project', 'doc', 'run')).toEqual(before)
})
it('rejects malformed inputs and forged sequence/history', () => {
  for (const patch of [
    { actionId: '../x' },
    { issueDigest: 'x' },
    { note: ' ' },
    { note: '\u0000' },
    { state: 'resolved' },
    { extra: true },
  ])
    expect(() => parsePresentationIssueActionInput({ ...action, ...patch })).toThrow()
  const ledger = append()
  expect(() => parsePresentationIssueLedger({ ...ledger, revision: 0 })).toThrow()
  expect(() =>
    parsePresentationIssueLedger({ ...ledger, actions: [{ ...ledger.actions[0], sequence: 2 }] }),
  ).toThrow()
})
it('rejects cross-document/request use and corrupted persisted records', () => {
  expect(() => store.issueActions('project', 'other', 'run')).toThrow('document_mismatch')
  expect(() => store.issueActions('project', 'doc', 'missing')).toThrow('not_found')
  append()
  const hash = (s: string) => createHash('sha256').update(s).digest('hex')
  const path = join(
    root,
    'projects',
    'presentations',
    hash('project'),
    `issue-actions-${hash('run')}.json`,
  )
  const record = JSON.parse(readFileSync(path, 'utf8'))
  record.ledger.actions[0].note = 'tampered'
  writeFileSync(path, JSON.stringify(record))
  expect(() => store.issueActions('project', 'doc', 'run')).toThrow('invalid_state')
})
