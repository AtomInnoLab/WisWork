import { afterEach, expect, it } from 'vitest'
import { mkdtempSync, rmSync, readFileSync, writeFileSync, watch } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, dirname } from 'node:path'
import { createHash } from 'node:crypto'
import { PresentationResearchStore } from '../src/presentation-research-store'
const roots: string[] = []
afterEach(() => {
  for (const p of roots.splice(0)) rmSync(p, { recursive: true, force: true })
})
const draft = { scope: '合成中断研究', sources: [], facts: [] }
const hash = (v: string) => createHash('sha256').update(v).digest('hex')
async function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'abandon-research-'))
  roots.push(root)
  const store = new PresentationResearchStore(root),
    begun = await store.begin('doc', 'project', 0, 'A', draft)
  return {
    root,
    store,
    record: begun.record,
    path: join(root, 'presentation-research', hash('doc'), hash('project'), 'state.json'),
  }
}
it('explicitly ends a running original after restart without changing its shape or replaying', async () => {
  const f = await fixture(),
    store = new PresentationResearchStore(f.root),
    ended = await store.abandon('doc', 'project', 1, 'A', f.record.draftDigest)
  expect(ended).toEqual({
    ...f.record,
    state: 'failed',
    error: 'aborted',
    finishedAt: expect.any(String),
  })
  expect(await store.abandon('doc', 'project', 1, 'A', f.record.draftDigest)).toEqual(ended)
  expect(await store.history('doc', 'project')).toMatchObject({
    version: 1,
    revision: 2,
    totalRecords: 1,
  })
  expect((await store.begin('doc', 'project', 0, 'A', draft)).record).toEqual(ended)
})
it('retains V2 and never changes sequence or tombstones when ending a running record', async () => {
  const f = await fixture()
  await f.store.finish('doc', 'project', 'A', { state: 'completed', sources: [] })
  await f.store.deleteRecord('doc', 'project', 2, 'delete-A', 'A', f.record.draftDigest)
  const b = await f.store.begin('doc', 'project', 3, 'B', draft),
    ended = await f.store.abandon('doc', 'project', 4, 'B', b.record.draftDigest)
  expect(ended.sequence).toBe(2)
  expect(await f.store.history('doc', 'project')).toMatchObject({
    version: 2,
    lastSequence: 2,
    totalRecords: 1,
    revision: 5,
  })
  expect(await f.store.deletedReceipt('doc', 'project', 'delete-A')).toHaveProperty('sequence', 1)
  await expect(f.store.abandon('doc', 'project', 5, 'A', f.record.draftDigest)).rejects.toThrow(
    'record_deleted',
  )
})
it('checks CAS and exact digest; rejects completed and other failed outcomes without writing', async () => {
  const f = await fixture(),
    before = readFileSync(f.path, 'utf8')
  await expect(f.store.abandon('doc', 'project', 0, 'A', f.record.draftDigest)).rejects.toThrow(
    'revision_conflict',
  )
  await expect(f.store.abandon('doc', 'project', 1, 'A', 'f'.repeat(64))).rejects.toThrow(
    'request_conflict',
  )
  await expect(f.store.abandon('other', 'project', 1, 'A', f.record.draftDigest)).rejects.toThrow(
    'not_found',
  )
  expect(readFileSync(f.path, 'utf8')).toBe(before)
  await f.store.finish('doc', 'project', 'A', { state: 'completed', sources: [] })
  const completed = readFileSync(f.path, 'utf8')
  await expect(f.store.abandon('doc', 'project', 2, 'A', f.record.draftDigest)).rejects.toThrow(
    'record_not_running',
  )
  expect(readFileSync(f.path, 'utf8')).toBe(completed)
  const b = await f.store.begin('doc', 'project', 2, 'B', draft)
  await f.store.finish('doc', 'project', 'B', { state: 'failed', error: 'source_unavailable' })
  await expect(f.store.abandon('doc', 'project', 4, 'B', b.record.draftDigest)).rejects.toThrow(
    'record_not_running',
  )
})
it('preserves an already aborted terminal record including partial sources and finishedAt', async () => {
  const f = await fixture()
  const withSource = {
    scope: '合成带证据研究',
    sources: [{ id: 's', title: '合成来源', uri: 'https://example.com/s', excerpt: '原文' }],
    facts: [],
  }
  const b = await f.store.begin('doc', 'project', 1, 'B', withSource),
    ended = await f.store.finish('doc', 'project', 'B', {
      state: 'failed',
      error: 'aborted',
      sources: [{ sourceId: 's', status: 'missing', provenance: 'unavailable' }],
    })
  const before = await f.store.history('doc', 'project')
  expect(await f.store.abandon('doc', 'project', 0, 'B', b.record.draftDigest)).toEqual(ended)
  expect(await f.store.history('doc', 'project')).toEqual(before)
})
it('honors real precommit aborts and fails closed on checksum or draft corruption', async () => {
  const f = await fixture(),
    before = readFileSync(f.path, 'utf8'),
    cancelled = new AbortController()
  cancelled.abort()
  await expect(
    f.store.abandon('doc', 'project', 1, 'A', f.record.draftDigest, cancelled.signal),
  ).rejects.toThrow('aborted')
  const controller = new AbortController(),
    observer = watch(dirname(f.path), (_event, name) => {
      if (String(name).endsWith('.tmp')) controller.abort()
    })
  try {
    await expect(
      f.store.abandon('doc', 'project', 1, 'A', f.record.draftDigest, controller.signal),
    ).rejects.toThrow('aborted')
  } finally {
    observer.close()
  }
  expect(readFileSync(f.path, 'utf8')).toBe(before)
  const state = JSON.parse(before).state
  state.records[0].draft.scope = '篡改'
  writeFileSync(f.path, JSON.stringify({ state, checksum: hash(JSON.stringify(state)) }))
  await expect(f.store.abandon('doc', 'project', 1, 'A', f.record.draftDigest)).rejects.toThrow(
    'invalid_state',
  )
  writeFileSync(f.path, JSON.stringify({ state, checksum: 'a'.repeat(64) }))
  await expect(f.store.abandon('doc', 'project', 1, 'A', f.record.draftDigest)).rejects.toThrow(
    'invalid_state',
  )
})
it('clamps completion against all state times and serializes concurrent endings once', async () => {
  const f = await fixture(),
    state = JSON.parse(readFileSync(f.path, 'utf8')).state
  state.records[0].startedAt = '2099-01-01T00:00:00.000Z'
  writeFileSync(f.path, JSON.stringify({ state, checksum: hash(JSON.stringify(state)) }))
  const [a, b] = await Promise.all([
    f.store.abandon('doc', 'project', 1, 'A', f.record.draftDigest),
    f.store.abandon('doc', 'project', 1, 'A', f.record.draftDigest),
  ])
  expect(a).toEqual(b)
  expect(a.finishedAt).toBe('2099-01-01T00:00:00.000Z')
  expect((await f.store.history('doc', 'project')).revision).toBe(2)
})
