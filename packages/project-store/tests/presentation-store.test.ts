import { mkdtempSync, readFileSync, readdirSync, renameSync, symlinkSync, writeFileSync } from 'node:fs'
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
    expect(reload.begin('project', 'opaque:/document', 'one', { title: 'A' }).result).toEqual({ status: 'compiled', pptxBase64: 'eA==' })
    expect(() => reload.begin('project', 'opaque:/document', 'one', { title: 'B' })).toThrow('request_conflict')
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
    const directory = join(root, 'projects', 'presentations', readdirSync(join(root, 'projects', 'presentations'))[0]!)
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
    vi.mocked(renameSync).mockImplementationOnce(() => { throw new Error('disk failure') })
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
