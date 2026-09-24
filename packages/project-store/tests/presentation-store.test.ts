import {
  mkdtempSync,
  readFileSync,
  readdirSync,
  renameSync,
  symlinkSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import { PresentationStore } from '../src/presentation-store.js'

vi.mock('node:fs', async (original) => {
  const fs = await original<typeof import('node:fs')>()
  return { ...fs, renameSync: vi.fn(fs.renameSync) }
})

describe('presentation receipts', () => {
  it('persists complete inputs/results, reserves idempotency keys and keeps prior output on failure', () => {
    const root = mkdtempSync(join(tmpdir(), 'presentation-store-'))
    const store = new PresentationStore(root)
    const first = store.begin('project', 'opaque:/document', 'one', { title: 'A' })
    store.complete(first, { status: 'compiled', pptxBase64: 'eA==' })
    const reload = new PresentationStore(root)
    expect(reload.begin('project', 'opaque:/document', 'one', { title: 'A' }).result).toEqual({
      status: 'compiled',
      pptxBase64: 'eA==',
    })
    expect(() => reload.begin('project', 'opaque:/document', 'one', { title: 'B' })).toThrow(
      'request_conflict',
    )
    reload.begin('project', 'opaque:/document', 'two', { title: 'B' })
    expect(reload.latest('project', 'opaque:/document')?.requestId).toBe('one')
    expect(() => reload.latest('project', 'another')).toThrow('document_mismatch')
  })
  it('denies unsafe IDs and unknown persisted versions without modifying them', () => {
    const root = mkdtempSync(join(tmpdir(), 'presentation-store-'))
    const store = new PresentationStore(root)
    for (const id of ['../x', '/tmp/x', '..', 'a/b', 'a\\b']) {
      expect(() => store.begin(id, 'doc', 'one', {})).toThrow('invalid_request')
      expect(() => store.begin('ok', 'doc', id, {})).toThrow('invalid_request')
    }
    store.begin('ok', 'doc', 'one', {})
    const directory = join(
      root,
      'projects',
      'presentations',
      readdirSync(join(root, 'projects', 'presentations'))[0]!,
    )
    const path = join(directory, 'project.json')
    const bad = JSON.stringify({ ...JSON.parse(readFileSync(path, 'utf8')), version: 42 })
    writeFileSync(path, bad)
    expect(() => store.begin('ok', 'doc', 'two', {})).toThrow('invalid_state')
    expect(readFileSync(path, 'utf8')).toBe(bad)
  })
  it('a late retry of an older request never supersedes a newer completed request', () => {
    const store = new PresentationStore(mkdtempSync(join(tmpdir(), 'presentation-store-')))
    const old = store.begin('p', 'd', 'old', {})
    const newer = store.begin('p', 'd', 'new', {})
    store.complete(newer, { value: 2 })
    store.complete(old, { value: 1 })
    expect(store.latest('p', 'd')?.requestId).toBe('new')
  })
  it('surfaces authoritative write failures and leaves the previous completed receipt readable', () => {
    const store = new PresentationStore(mkdtempSync(join(tmpdir(), 'presentation-store-')))
    store.complete(store.begin('p', 'd', 'first', {}), { value: 1 })
    const next = store.begin('p', 'd', 'next', {})
    vi.mocked(renameSync).mockImplementationOnce(() => {
      throw new Error('disk failure')
    })
    expect(() => store.complete(next, { value: 2 })).toThrow('disk failure')
    expect(store.latest('p', 'd')?.requestId).toBe('first')
    store.complete(next, { value: 2 })
    expect(store.latest('p', 'd')?.requestId).toBe('next')
  })
  it('rejects symlinked storage directories', () => {
    const root = mkdtempSync(join(tmpdir(), 'presentation-store-'))
    const outside = mkdtempSync(join(tmpdir(), 'presentation-outside-'))
    symlinkSync(outside, join(root, 'projects'), 'dir')
    expect(() => new PresentationStore(root).begin('p', 'd', 'r', {})).toThrow('invalid_state')
    expect(readdirSync(outside)).toEqual([])
  })
})

it('retrieves explicit requests and bounded newest-first history with document binding', () => {
  const store = new PresentationStore(mkdtempSync(join(tmpdir(), 'presentation-store-')))
  for (let i = 0; i < 25; i++) store.begin('p', 'd', `r${i}`, { title: `${i}` })
  expect(store.history('p', 'd').map((r) => r.sequence)).toEqual(
    Array.from({ length: 20 }, (_, i) => 25 - i),
  )
  expect(store.request('p', 'd', 'r0')?.deck).toEqual({ title: '0' })
  expect(store.request('p', 'd', 'missing')).toBeUndefined()
  expect(store.history('missing', 'd')).toEqual([])
  expect(() => store.request('p', 'other', 'r0')).toThrow('document_mismatch')
  expect(() => store.history('p', 'other')).toThrow('document_mismatch')
  expect(() => store.request('p', 'd', '../x')).toThrow('invalid_request')
})

function planFixture() {
  const root = mkdtempSync(join(tmpdir(), 'presentation-plan-'))
  const store = new PresentationStore(root)
  const directory = () =>
    join(
      root,
      'projects',
      'presentations',
      readdirSync(join(root, 'projects', 'presentations'))[0]!,
    )
  return { root, store, directory }
}

describe('durable plans', () => {
  it('creates and reloads document-bound plans with compare-and-swap and idempotent retries', () => {
    const { root, store } = planFixture()
    expect(store.plan('p', 'd')).toBeUndefined()
    expect(() => store.savePlan('p', 'd', 1, { title: 'A' })).toThrow('revision_conflict')
    const first = store.savePlan('p', 'd', 0, { title: 'A' })
    expect(first.revision).toBe(1)
    expect(first.revisions).toEqual([{ revision: 1, inputDigest: first.inputDigest,
      createdAt: expect.any(String) }])
    expect(new PresentationStore(root).plan('p', 'd')).toEqual(first)
    expect(store.savePlan('p', 'd', 0, { title: 'A' })).toEqual(first)
    expect(store.savePlan('p', 'd', 1, { title: 'A' })).toEqual(first)
    expect(() => store.savePlan('p', 'd', 0, { title: 'B' })).toThrow('revision_conflict')
    const second = store.savePlan('p', 'd', 1, { title: 'B' })
    expect(second.revision).toBe(2)
    expect(second.revisions?.map((entry) => entry.revision)).toEqual([1, 2])
    expect(new PresentationStore(root).plan('p', 'd')?.revisions).toEqual(second.revisions)
    expect(() => store.savePlan('p', 'd', 0, { title: 'B' })).toThrow('revision_conflict')
    expect(() => store.plan('p', 'foreign')).toThrow('document_mismatch')
    expect(() => store.savePlan('p', 'foreign', 2, {})).toThrow('document_mismatch')
    vi.mocked(renameSync).mockImplementationOnce(() => {
      throw new Error('disk failure')
    })
    expect(() => store.savePlan('p', 'd', 2, { title: 'C' })).toThrow('disk failure')
    expect(store.plan('p', 'd')).toEqual(second)
  })
  it('rejects corrupted and symlinked plan files, including dangling links', () => {
    for (const mode of ['corrupt', 'symlink', 'dangling']) {
      const { store, directory } = planFixture()
      store.savePlan('p', 'd', 0, { title: 'A' })
      const path = join(directory(), 'plan.json')
      const saved = JSON.parse(readFileSync(path, 'utf8'))
      if (mode === 'corrupt')
        writeFileSync(path, JSON.stringify({ ...saved, plan: { title: 'tampered' } }))
      else {
        renameSync(path, `${path}.original`)
        symlinkSync(mode === 'symlink' ? `${path}.original` : `${path}.missing`, path)
      }
      expect(() => store.plan('p', 'd')).toThrow('invalid_state')
      expect(() => store.savePlan('p', 'd', 1, { title: 'B' })).toThrow('invalid_state')
    }
  })
  it('bounds revision history and rejects tampered revision metadata', () => {
    const { root, store, directory } = planFixture()
    for (let revision = 0; revision < 34; revision++)
      store.savePlan('p', 'd', revision, { title: String(revision) })
    const saved = new PresentationStore(root).plan('p', 'd')!
    expect(saved.revisions).toHaveLength(32)
    expect(saved.revisions?.[0]?.revision).toBe(3)
    expect(saved.revisions?.at(-1)?.revision).toBe(34)
    const path = join(directory(), 'plan.json')
    writeFileSync(path, JSON.stringify({ ...saved, revisions: saved.revisions?.map((entry, index) =>
      index === 0 ? { ...entry, revision: 2 } : entry) }))
    expect(() => new PresentationStore(root).plan('p', 'd')).toThrow('invalid_state')
  })
  it('loads pre-history plans and starts revision events on the next save', () => {
    const { store, directory } = planFixture()
    const first = store.savePlan('p', 'd', 0, { title: 'A' })
    const path = join(directory(), 'plan.json')
    const { revisions: _revisions, ...legacy } = first
    writeFileSync(path, JSON.stringify(legacy))
    expect(store.plan('p', 'd')?.revisions).toBeUndefined()
    expect(store.savePlan('p', 'd', 1, { title: 'B' }).revisions?.map((entry) => entry.revision)).toEqual([2])
  })
  it('bounds and validates persisted JSON input', () => {
    const { store } = planFixture()
    for (const plan of [
      null,
      [],
      { a: undefined },
      { a: NaN },
      { a: Array(2) },
      { a: new Date() },
      { title: '界'.repeat(192 * 1024) },
    ])
      expect(() => store.savePlan('p', 'd', 0, plan)).toThrow('invalid_plan')
    expect(() => store.savePlan('p', 'd', -1, {})).toThrow('invalid_request')
  })
  it('snapshots plan revisions in receipts, preserves them on completion and detects tampering', () => {
    const { store, directory } = planFixture()
    const plan = { title: 'A' }
    store.savePlan('p', 'd', 0, plan)
    const binding = { revision: 1, plan }
    const receipt = store.begin('p', 'd', 'r', {}, binding)
    expect(receipt.plan).toEqual(binding)
    expect(store.begin('p', 'd', 'r', {}, binding)).toEqual(receipt)
    expect(() => store.begin('p', 'd', 'r', {})).toThrow('request_conflict')
    expect(() => store.begin('p', 'd', 'r', {}, { ...binding, revision: 2 })).toThrow(
      'request_conflict',
    )
    expect(() => store.begin('p', 'd', 'r', {}, { revision: 1, plan: { title: 'B' } })).toThrow(
      'request_conflict',
    )
    store.savePlan('p', 'd', 1, { title: 'B' })
    store.complete(receipt, { ok: true })
    expect(store.request('p', 'd', 'r')?.plan).toEqual(binding)
    const path = join(
      directory(),
      readdirSync(directory()).find((name) => /^[a-f0-9]{64}\.json$/.test(name))!,
    )
    const saved = JSON.parse(readFileSync(path, 'utf8'))
    writeFileSync(path, JSON.stringify({ ...saved, plan: { ...binding, revision: 2 } }))
    expect(() => store.request('p', 'd', 'r')).toThrow('invalid_state')
  })
})
