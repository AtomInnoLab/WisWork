import { afterEach, expect, it } from 'vitest'
import {
  mkdtempSync,
  rmSync,
  readFileSync,
  writeFileSync,
  watch,
  symlinkSync,
  unlinkSync,
} from 'node:fs'
import { createHash } from 'node:crypto'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PresentationResearchStore } from '../src/presentation-research-store'
import {
  parsePresentationResearchHistory,
  parsePresentationResearchSummary,
  parsePresentationResearchDeleteReceipt,
} from '../src/presentation-research'
const roots: string[] = []
afterEach(() => {
  for (const p of roots.splice(0)) rmSync(p, { recursive: true, force: true })
})
function fixture() {
  const p = mkdtempSync(join(tmpdir(), 'research-clean-'))
  roots.push(p)
  return { p, store: new PresentationResearchStore(p) }
}
const draft = { scope: '合成整理', sources: [], facts: [] }
it('migrates only after finished deletion, preserves sequence, restores tombstone and same lost-ACK receipt', async () => {
  const { p, store } = fixture()
  const a = await store.begin('doc', 'project', 0, 'A', draft)
  await store.finish('doc', 'project', 'A', { state: 'completed', sources: [] })
  await store.begin('doc', 'project', 2, 'B', draft)
  expect((await store.history('doc', 'project')).version).toBe(1)
  const receipt = await store.deleteRecord(
    'doc',
    'project',
    3,
    'delete-A',
    'A',
    a.record.draftDigest,
  )
  expect(parsePresentationResearchDeleteReceipt(receipt)).toEqual(receipt)
  const restarted = new PresentationResearchStore(p)
  expect(await restarted.deletedReceipt('doc', 'project', 'delete-A')).toEqual(receipt)
  expect(
    await restarted.deleteRecord('doc', 'project', 3, 'delete-A', 'A', a.record.draftDigest),
  ).toEqual(receipt)
  const h = await restarted.history('doc', 'project')
  expect(h).toMatchObject({ version: 2, lastSequence: 2, totalRecords: 1, revision: 4 })
  expect(h.records.map((r) => r.sequence)).toEqual([2])
  expect(parsePresentationResearchSummary(await restarted.summary('doc', 'project')).version).toBe(
    2,
  )
  await expect(restarted.read('doc', 'project', 'A')).rejects.toThrow('record_deleted')
  await expect(restarted.begin('doc', 'project', 4, 'A', draft)).rejects.toThrow('record_deleted')
  await expect(
    restarted.finish('doc', 'project', 'A', { state: 'completed', sources: [] }),
  ).rejects.toThrow('record_deleted')
  const c = await restarted.begin('doc', 'project', 4, 'C', draft)
  expect(c.record.sequence).toBe(3)
})
it('strictly parses V2 gaps and receipt shape without relaxing V1', () => {
  const empty = {
    version: 2,
    documentId: 'doc',
    projectId: 'project',
    revision: 3,
    lastSequence: 1,
    totalRecords: 0,
    records: [],
  }
  expect(parsePresentationResearchHistory(empty)).toEqual(empty)
  expect(() => parsePresentationResearchHistory({ ...empty, version: 1 })).toThrow()
  for (const r of [
    {
      version: 1,
      documentId: 'doc',
      projectId: 'project',
      ledgerId: 'A',
      sequence: 1,
      draftDigest: 'a'.repeat(64),
      deleteId: 'del',
      deletedAt: '2026-09-29T00:00:00.000Z',
      revision: 3,
    },
  ]) {
    expect(parsePresentationResearchDeleteReceipt(r)).toEqual(r)
    expect(() => parsePresentationResearchDeleteReceipt({ ...r, extra: true })).toThrow()
    expect(() =>
      parsePresentationResearchDeleteReceipt({ ...r, deletedAt: '2026-02-30T00:00:00.000Z' }),
    ).toThrow()
  }
})

const hash = (v: string) => createHash('sha256').update(v).digest('hex')
const statePath = (p: string) =>
  join(p, 'presentation-research', hash('doc'), hash('project'), 'state.json')
const publishFixture = (p: string, state: unknown) =>
  writeFileSync(statePath(p), JSON.stringify({ state, checksum: hash(JSON.stringify(state)) }))
it('rejects running, digest/CAS conflicts and pre-commit cancellation without migrating', async () => {
  const { p, store } = fixture()
  const a = await store.begin('doc', 'project', 0, 'A', draft)
  await expect(
    store.deleteRecord('doc', 'project', 1, 'delete-A', 'A', a.record.draftDigest),
  ).rejects.toThrow('record_running')
  await store.finish('doc', 'project', 'A', { state: 'failed', error: 'aborted' })
  const before = readFileSync(statePath(p), 'utf8')
  await expect(
    store.deleteRecord('doc', 'project', 1, 'delete-A', 'A', a.record.draftDigest),
  ).rejects.toThrow('revision_conflict')
  await expect(
    store.deleteRecord('doc', 'project', 2, 'delete-A', 'A', 'f'.repeat(64)),
  ).rejects.toThrow('request_conflict')
  const controller = new AbortController()
  controller.abort()
  await expect(
    store.deleteRecord(
      'doc',
      'project',
      2,
      'delete-A',
      'A',
      a.record.draftDigest,
      controller.signal,
    ),
  ).rejects.toThrow('aborted')
  expect(readFileSync(statePath(p), 'utf8')).toBe(before)
  const live = new AbortController(),
    observer = watch(join(statePath(p), '..'), (_event, name) => {
      if (String(name).endsWith('.tmp')) live.abort()
    })
  try {
    await expect(
      store.deleteRecord('doc', 'project', 2, 'delete-A', 'A', a.record.draftDigest, live.signal),
    ).rejects.toThrow('aborted')
  } finally {
    observer.close()
  }
  expect(readFileSync(statePath(p), 'utf8')).toBe(before)
  await expect(store.deletedReceipt('doc', 'project', 'delete-A')).rejects.toThrow('not_found')
  const receipt = await store.deleteRecord(
    'doc',
    'project',
    2,
    'delete-A',
    'A',
    a.record.draftDigest,
  )
  expect(receipt.revision).toBe(3)
  await expect(
    store.deleteRecord('doc', 'project', 2, 'delete-A', 'A', 'f'.repeat(64)),
  ).rejects.toThrow('request_conflict')
  await expect(
    store.deleteRecord('doc', 'project', 2, 'delete-A', 'B', a.record.draftDigest),
  ).rejects.toThrow('request_conflict')
})
it('deletes an ended archive outside the 32 window and releases full 128 capacity without resetting sequence', async () => {
  const { p, store } = fixture()
  const a = await store.begin('doc', 'project', 0, 'A', draft)
  await store.finish('doc', 'project', 'A', { state: 'completed', sources: [] })
  for (let index = 2; index <= 128; index++)
    await store.begin('doc', 'project', index, 'R' + index, draft)
  expect((await store.summary('doc', 'project')).records.some((r) => r.id === 'A')).toBe(false)
  await expect(store.begin('doc', 'project', 129, 'overflow', draft)).rejects.toThrow(
    'quota_exceeded',
  )
  await store.deleteRecord('doc', 'project', 129, 'delete-A', 'A', a.record.draftDigest)
  expect(await store.latestCompleted('doc', 'project')).toBeNull()
  const next = await new PresentationResearchStore(p).begin('doc', 'project', 130, 'new', draft)
  expect(next.record.sequence).toBe(129)
  const history = await store.history('doc', 'project')
  expect(history).toMatchObject({ version: 2, lastSequence: 129, totalRecords: 128, revision: 131 })
  expect(history.records).toHaveLength(32)
  await expect(store.read('doc', 'project', 'A')).rejects.toThrow('record_deleted')
})
it('enforces 4096 tombstone quota and recovers existing receipts at capacity', async () => {
  const { p, store } = fixture()
  await store.begin('doc', 'project', 0, 'active', draft)
  const record = await store.finish('doc', 'project', 'active', { state: 'completed', sources: [] })
  const tombstones = Array.from({ length: 4096 }, (_, i) => ({
    version: 1,
    documentId: 'doc',
    projectId: 'project',
    ledgerId: 'deleted' + i,
    sequence: i + 1,
    draftDigest: record.draftDigest,
    deleteId: 'delete' + i,
    deletedAt: record.finishedAt!,
    revision: 3 * (i + 1),
  }))
  const state = {
    version: 2,
    documentId: 'doc',
    projectId: 'project',
    revision: 12290,
    totalRecords: 1,
    lastSequence: 4097,
    records: [{ ...record, sequence: 4097 }],
    tombstones,
  }
  publishFixture(p, state)
  const before = readFileSync(statePath(p), 'utf8')
  expect(await store.deletedReceipt('doc', 'project', 'delete0')).toEqual(tombstones[0])
  expect(
    await store.deleteRecord('doc', 'project', 0, 'delete0', 'deleted0', record.draftDigest),
  ).toEqual(tombstones[0])
  await expect(
    store.deleteRecord('doc', 'project', 12290, 'newdelete', 'active', record.draftDigest),
  ).rejects.toThrow('cleanup_quota_exceeded')
  expect(readFileSync(statePath(p), 'utf8')).toBe(before)
  expect((await store.read('doc', 'project', 'active')).sequence).toBe(4097)
})
it('fails closed on V2 corruption including recomputed checksum, draft hash, tombstone collisions and unsafe file', async () => {
  const { p, store } = fixture()
  const a = await store.begin('doc', 'project', 0, 'A', draft)
  await store.finish('doc', 'project', 'A', { state: 'completed', sources: [] })
  await store.begin('doc', 'project', 2, 'B', draft)
  await store.deleteRecord('doc', 'project', 3, 'delete-A', 'A', a.record.draftDigest)
  const good = JSON.parse(readFileSync(statePath(p), 'utf8')).state
  for (const change of [
    (s: typeof good) => {
      s.records[0].sequence = 1
    },
    (s: typeof good) => {
      s.tombstones[0].ledgerId = 'B'
    },
    (s: typeof good) => {
      s.tombstones[0].revision = 1
    },
    (s: typeof good) => {
      s.tombstones[0].sequence = 3
    },
    (s: typeof good) => {
      s.records[0].draft.scope = 'changed'
    },
    (s: typeof good) => {
      s.revision++
    },
    (s: typeof good) => {
      s.tombstones.push(s.tombstones[0])
    },
  ]) {
    const bad = structuredClone(good)
    change(bad)
    publishFixture(p, bad)
    await expect(store.deletedReceipt('doc', 'project', 'delete-A')).rejects.toThrow(
      'invalid_state',
    )
  }
  publishFixture(p, good)
  const target = join(p, 'outside.json')
  writeFileSync(target, readFileSync(statePath(p)))
  unlinkSync(statePath(p))
  symlinkSync(target, statePath(p))
  await expect(store.history('doc', 'project')).rejects.toThrow('invalid_state')
})
it('keeps empty cleaned projects V2 and starts new sequences after every active record is removed', async () => {
  const { p, store } = fixture()
  const a = await store.begin('doc', 'project', 0, 'only', draft)
  await store.finish('doc', 'project', 'only', { state: 'completed', sources: [] })
  await store.deleteRecord('doc', 'project', 2, 'delete-only', 'only', a.record.draftDigest)
  expect(await store.history('doc', 'project')).toEqual({
    version: 2,
    documentId: 'doc',
    projectId: 'project',
    revision: 3,
    totalRecords: 0,
    lastSequence: 1,
    records: [],
  })
  const restarted = new PresentationResearchStore(p),
    next = await restarted.begin('doc', 'project', 3, 'next', draft)
  expect(next.record.sequence).toBe(2)
  expect((await restarted.history('doc', 'project')).version).toBe(2)
})
it('rejects V2 revision impossibilities and keeps old V1 summary exact', async () => {
  const { store } = fixture()
  await store.begin('doc', 'project', 0, 'A', draft)
  const old = await store.summary('doc', 'project')
  expect(Object.keys(old).sort()).toEqual([
    'documentId',
    'projectId',
    'records',
    'revision',
    'totalRecords',
    'version',
  ])
  const empty = {
    version: 2,
    documentId: 'doc',
    projectId: 'project',
    revision: 3,
    lastSequence: 1,
    totalRecords: 0,
    records: [],
  }
  for (const changed of [
    { ...empty, revision: 2 },
    { ...empty, revision: 4 },
    { ...empty, lastSequence: 2 },
    { ...empty, tombstones: [] },
    { ...empty, totalRecords: 1 },
    { ...empty, lastSequence: 1.5 },
  ])
    expect(() => parsePresentationResearchHistory(changed)).toThrow('invalid_state')
  const receipt = {
    version: 1,
    documentId: 'doc',
    projectId: 'project',
    ledgerId: 'A',
    sequence: 1,
    draftDigest: 'a'.repeat(64),
    deleteId: 'del',
    deletedAt: '2026-09-29T00:00:00.000Z',
    revision: 3,
  }
  for (const changed of [
    { ...receipt, sequence: 0 },
    { ...receipt, revision: 0 },
    { ...receipt, draftDigest: 'bad' },
    { ...receipt, deleteId: '../x' },
    { ...receipt, deletedAt: '2026-09-29T08:00:00+08:00' },
  ])
    expect(() => parsePresentationResearchDeleteReceipt(changed)).toThrow('invalid_state')
})
