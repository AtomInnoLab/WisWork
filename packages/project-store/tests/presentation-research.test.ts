import { expect, it } from 'vitest'
import { parsePresentationResearchDraft } from '../src/presentation-research'
import { PresentationResearchStore } from '../src/presentation-research-store'
it('preserves both conflicting conclusions and rejects unknown fields', () => {
  const draft = {
    scope: '研究范围',
    sources: [],
    facts: [
      {
        claimId: 'a',
        statement: '原文判断',
        type: 'judgment',
        sourceRefs: [],
        sourceTier: 'unverified',
        slideRefs: [],
        confidence: 'low',
        reviewStatus: 'needs_review',
        conflictsWith: ['b'],
      },
      {
        claimId: 'b',
        statement: '相反假设',
        type: 'assumption',
        sourceRefs: [],
        sourceTier: 'unverified',
        slideRefs: [],
        confidence: 'low',
        reviewStatus: 'needs_review',
        conflictsWith: ['a'],
      },
    ],
  }
  expect(parsePresentationResearchDraft(draft).facts).toHaveLength(2)
  expect(() => parsePresentationResearchDraft({ ...draft, extra: true })).toThrow('invalid_state')
  expect(PresentationResearchStore).toBeTypeOf('function')
})
import { mkdtempSync, rmSync, writeFileSync, symlinkSync, readdirSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach } from 'vitest'
import {
  parsePresentationResearchRecord,
  parsePresentationResearchHistory,
  parsePresentationResearchSummary,
} from '../src/presentation-research'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const root = () => {
  const p = mkdtempSync(join(tmpdir(), 'research-'))
  roots.push(p)
  return p
}
const draft = () => ({ scope: '原文研究未核验', sources: [], facts: [] })
it('uses CAS and idempotent failed/running attempts and restores archived records beyond its 32 window', async () => {
  const path = root(),
    store = new PresentationResearchStore(path),
    d = draft()
  const first = await store.begin('doc', 'project', 0, 'first', d)
  expect(first.created).toBe(true)
  const attempts = await Promise.allSettled([
    store.begin('doc', 'project', 1, 'second', d),
    store.begin('doc', 'project', 1, 'third', d),
  ])
  expect(attempts.filter((r) => r.status === 'fulfilled')).toHaveLength(1)
  expect(attempts.filter((r) => r.status === 'rejected')).toHaveLength(1)
  expect((await store.begin('doc', 'project', 0, 'first', d)).created).toBe(false)
  await expect(
    store.begin('doc', 'project', 0, 'first', { ...d, scope: 'changed' }),
  ).rejects.toThrow('request_conflict')
  const failed = await store.finish('doc', 'project', 'first', {
    state: 'failed',
    error: 'aborted',
  })
  expect(
    await store.finish('doc', 'project', 'first', { error: 'aborted', state: 'failed' }),
  ).toEqual(failed)
  expect((await store.begin('doc', 'project', 0, 'first', d)).record).toEqual(failed)
  for (let i = 2; i < 34; i++)
    await store.begin(
      'doc',
      'project',
      (await store.summary('doc', 'project')).revision,
      'ledger-' + i,
      d,
    )
  const restarted = new PresentationResearchStore(path),
    summary = await restarted.summary('doc', 'project')
  expect(parsePresentationResearchSummary(summary).records).toHaveLength(32)
  expect(summary.records[0]!.sequence).toBe(3)
  expect(summary.totalRecords).toBe(34)
  expect(await restarted.read('doc', 'project', 'first')).toEqual(failed)
  expect(
    parsePresentationResearchHistory(await restarted.history('doc', 'project')).totalRecords,
  ).toBe(34)
  expect((await restarted.summary('other', 'project')).totalRecords).toBe(0)
})
it('enforces 128 archived records and keeps existing records readable at quota', async () => {
  const path = root(),
    store = new PresentationResearchStore(path)
  for (let i = 0; i < 128; i++) await store.begin('doc', 'project', i, 'ledger-' + i, draft())
  await expect(store.begin('doc', 'project', 128, 'overflow', draft())).rejects.toThrow(
    'quota_exceeded',
  )
  expect(
    (await new PresentationResearchStore(path).read('doc', 'project', 'ledger-0')).sequence,
  ).toBe(1)
})
it('rejects corrupted checksum and symlink roots without leaking disk details', async () => {
  const path = root(),
    store = new PresentationResearchStore(path)
  await store.begin('doc', 'project', 0, 'ledger', draft())
  const docs = join(path, 'presentation-research')
  const doc = join(docs, readdirSync(docs)[0]!)
  const project = join(doc, readdirSync(doc)[0]!)
  writeFileSync(join(project, 'state.json'), 'private path sensitive error')
  await expect(store.summary('doc', 'project')).rejects.toThrow(/^invalid_state$/)
  const other = root()
  symlinkSync(docs, join(other, 'presentation-research'))
  await expect(new PresentationResearchStore(other).summary('doc', 'project')).rejects.toThrow(
    'invalid_state',
  )
})
it('rejects broken references, spoofed evidence, invalid dates and oversized UTF8 drafts', async () => {
  const d = {
    scope: 'scope',
    sources: [
      { id: 'source', title: 'original', uri: 'attachment:' + 'a'.repeat(64), excerpt: 'literal' },
    ],
    facts: [
      {
        claimId: 'claim',
        statement: 'fact',
        type: 'fact',
        sourceRefs: ['source'],
        sourceTier: 'primary',
        slideRefs: [],
        confidence: 'high',
        reviewStatus: 'needs_review',
        conflictsWith: [],
      },
    ],
  }
  for (const invalid of [
    { ...d, facts: [{ ...d.facts[0], sourceRefs: ['missing'] }] },
    { ...d, facts: [{ ...d.facts[0], conflictsWith: ['claim'] }] },
    { ...d, sources: [{ ...d.sources[0], asOf: '2026-02-30' }] },
    { ...d, sources: [{ ...d.sources[0], uri: 'https://user:password@example.com/' }] },
    {
      ...d,
      sources: Array.from({ length: 64 }, (_, i) => ({
        ...d.sources[0],
        id: 'source-' + i,
        excerpt: '字'.repeat(12000),
      })),
      facts: [],
    },
  ])
    expect(() => parsePresentationResearchDraft(invalid)).toThrow('invalid_state')
  const store = new PresentationResearchStore(root()),
    begun = await store.begin('doc', 'project', 0, 'ledger', parsePresentationResearchDraft(d))
  expect(() =>
    parsePresentationResearchRecord({
      ...begun.record,
      state: 'completed',
      finishedAt: begun.record.startedAt,
      sources: [
        {
          sourceId: 'source',
          attachmentId: 'a'.repeat(64),
          status: 'found',
          offset: 0,
          provenance: 'fetched_url_matched',
          sha256: 'a'.repeat(64),
          retrievedAt: begun.record.startedAt,
        },
      ],
    }),
  ).toThrow('invalid_state')
})
it('finds the latest completed archive when all 32 recent records are unresolved', async () => {
  const path = root(),
    store = new PresentationResearchStore(path)
  await store.begin('doc', 'project', 0, 'completed', draft())
  const completed = await store.finish('doc', 'project', 'completed', {
    state: 'completed',
    sources: [],
  })
  for (let i = 0; i < 32; i++) await store.begin('doc', 'project', i + 2, 'running-' + i, draft())
  const reopened = new PresentationResearchStore(path)
  expect(
    (await reopened.summary('doc', 'project')).records.every((r) => r.state === 'running'),
  ).toBe(true)
  expect(await reopened.latestCompleted('doc', 'project')).toEqual(completed)
  expect((await reopened.summary('doc', 'project')).revision).toBe(34)
  expect(await reopened.latestCompleted('other', 'project')).toBeNull()
})
