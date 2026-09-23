import { afterEach, describe, expect, it, vi } from 'vitest'
import { createHash } from 'node:crypto'
import {
  mkdtempSync,
  rmSync,
  readFileSync,
  writeFileSync,
  symlinkSync,
  unlinkSync,
  renameSync,
  readdirSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PresentationStore } from '../src/presentation-store'
vi.mock('node:fs', async (original) => {
  const fs = await original<typeof import('node:fs')>()
  return { ...fs, renameSync: vi.fn(fs.renameSync) }
})
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const digest = (value: string) => createHash('sha256').update(value).digest('hex')
const deck = () => ({
  id: 'project',
  slides: [
    { id: 'first', title: 'First' },
    { id: 'second', title: 'Second' },
  ],
})
const plan = () => ({ revision: 1, plan: { projectId: 'project', slides: ['first', 'second'] } })
const result = () => ({
  pptxBase64: 'UEsDBAAAAAA=',
  sourceSlideId: '256#',
  report: { slideCount: 1 },
})
function setup() {
  const root = mkdtempSync(join(tmpdir(), 'ppt-production-'))
  roots.push(root)
  return {
    root,
    store: new PresentationStore(root),
    path: (request = 'request') =>
      join(
        root,
        'projects',
        'presentations',
        digest('project'),
        `production-${digest(request)}.json`,
      ),
  }
}
describe('durable page production', () => {
  it('freezes inputs idempotently and keeps production separate from legacy receipts', () => {
    const { store, root } = setup()
    const input = deck(),
      binding = plan()
    const created = store.beginProduction('project', 'doc', 'request', input, binding)
    expect(created.pages).toEqual([
      { pageId: 'first', state: 'pending', attempt: 0 },
      { pageId: 'second', state: 'pending', attempt: 0 },
    ])
    input.slides[0]!.title = 'changed'
    binding.plan.slides.push('other')
    const reopened = new PresentationStore(root)
    expect(reopened.production('project', 'doc', 'request')?.deck).toEqual(deck())
    expect(reopened.beginProduction('project', 'doc', 'request', deck(), plan())).toEqual(created)
    expect(() => reopened.beginProduction('project', 'doc', 'request', input, plan())).toThrow(
      'request_conflict',
    )
    expect(() => reopened.production('project', 'foreign')).toThrow('document_mismatch')
    expect(store.history('project', 'doc')).toEqual([])
    store.begin('project', 'doc', 'legacy', {})
    store.beginProduction('project', 'doc', 'next', deck(), plan())
    expect(store.production('project', 'doc')?.requestId).toBe('next')
    expect(store.history('project', 'doc')).toHaveLength(1)
  })
  it('updates one page with CAS while preserving concurrent updates to other pages', () => {
    const { store } = setup()
    const initial = store.beginProduction('project', 'doc', 'request', deck(), plan())
    const first = store.updateProductionPage(initial, 'first', { state: 'building', attempt: 1 })
    const both = store.updateProductionPage(initial, 'second', { state: 'building', attempt: 1 })
    expect(both.pages.map((p) => p.state)).toEqual(['building', 'building'])
    expect(() =>
      store.updateProductionPage(initial, 'first', { state: 'building', attempt: 1 }),
    ).toThrow('revision_conflict')
    const finished = store.updateProductionPage(first, 'first', {
      state: 'compiled',
      attempt: 1,
      result: result(),
    })
    expect(finished.pages[1]!.state).toBe('building')
    expect(finished.pages[0]!.resultDigest).toMatch(/^[a-f0-9]{64}$/)
    expect(() =>
      store.updateProductionPage(finished, 'first', { state: 'building', attempt: 2 }),
    ).toThrow('invalid_state')
  })
  it('retries failed and interrupted building pages with monotonically increasing attempts', () => {
    const { store } = setup()
    let record = store.beginProduction('project', 'doc', 'request', deck(), plan())
    expect(() =>
      store.updateProductionPage(record, 'first', {
        state: 'compiled',
        attempt: 0,
        result: result(),
      }),
    ).toThrow('invalid_state')
    record = store.updateProductionPage(record, 'first', { state: 'building', attempt: 1 })
    record = store.updateProductionPage(record, 'first', {
      state: 'failed',
      attempt: 1,
      error: 'compile_failed',
    })
    record = store.updateProductionPage(record, 'first', { state: 'building', attempt: 2 })
    record = store.updateProductionPage(record, 'first', { state: 'building', attempt: 3 })
    expect(record.pages[0]).toEqual({ pageId: 'first', state: 'building', attempt: 3 })
    expect(() =>
      store.updateProductionPage(record, 'first', {
        state: 'failed',
        attempt: 2,
        error: 'compile_failed',
      }),
    ).toThrow('invalid_state')
    expect(() =>
      store.updateProductionPage(record, 'unknown', { state: 'building', attempt: 1 }),
    ).toThrow('invalid_state')
  })
  it('rejects corrupted digests, frozen input tampering and symlinked records', () => {
    const { store, path } = setup()
    let record = store.beginProduction('project', 'doc', 'request', deck(), plan())
    record = store.updateProductionPage(record, 'first', { state: 'building', attempt: 1 })
    record = store.updateProductionPage(record, 'first', {
      state: 'compiled',
      attempt: 1,
      result: result(),
    })
    const saved = readFileSync(path(), 'utf8'),
      corrupt = JSON.parse(saved)
    corrupt.pages[0].result.pptxBase64 = 'YWJj'
    writeFileSync(path(), JSON.stringify(corrupt))
    expect(() => store.production('project', 'doc')).toThrow('invalid_state')
    writeFileSync(path(), saved)
    ;(record.deck as ReturnType<typeof deck>).slides[0]!.title = 'forged'
    expect(() =>
      store.updateProductionPage(record, 'second', { state: 'building', attempt: 1 }),
    ).toThrow('invalid_state')
    unlinkSync(path())
    symlinkSync('/tmp/nonexistent-production-record', path())
    expect(() => store.production('project', 'doc')).toThrow('invalid_state')
  })
  it('rejects invalid page IDs, result metadata and excessive production requests', () => {
    const { store } = setup()
    for (const input of [
      { slides: [] },
      { slides: [{ id: 'duplicate' }, { id: 'duplicate' }] },
      { slides: [{ id: '../unsafe' }] },
      { slides: Array.from({ length: 33 }, (_, i) => ({ id: String(i) })) },
    ])
      expect(() => store.beginProduction('project', 'doc', 'bad', input, plan())).toThrow()
    let record = store.beginProduction('project', 'doc', 'request', deck(), plan())
    record = store.updateProductionPage(record, 'first', { state: 'building', attempt: 1 })
    for (const invalid of [
      { ...result(), sourceSlideId: '255#' },
      { ...result(), sourceSlideId: '0256#' },
      { ...result(), pptxBase64: 'not-base64' },
      { ...result(), report: undefined },
      { ...result(), report: { bad: NaN } },
    ])
      expect(() =>
        store.updateProductionPage(record, 'first', {
          state: 'compiled',
          attempt: 1,
          result: invalid,
        }),
      ).toThrow('invalid_state')
    for (let i = 1; i < 32; i++) store.beginProduction('project', 'doc', `req${i}`, deck(), plan())
    expect(() => store.beginProduction('project', 'doc', 'overflow', deck(), plan())).toThrow(
      'output_too_large',
    )
  })
  it('enforces the aggregate decoded PPTX budget without corrupting a completed page', () => {
    const { store } = setup()
    let record = store.beginProduction('project', 'doc', 'request', deck(), plan())
    record = store.updateProductionPage(record, 'first', { state: 'building', attempt: 1 })
    record = store.updateProductionPage(record, 'first', {
      state: 'compiled',
      attempt: 1,
      result: { ...result(), pptxBase64: Buffer.alloc(6 * 1024 * 1024).toString('base64') },
    })
    record = store.updateProductionPage(record, 'second', { state: 'building', attempt: 1 })
    expect(() =>
      store.updateProductionPage(record, 'second', {
        state: 'compiled',
        attempt: 1,
        result: { ...result(), pptxBase64: Buffer.alloc(5 * 1024 * 1024).toString('base64') },
      }),
    ).toThrow('output_too_large')
    expect(store.production('project', 'doc')?.pages.map((page) => page.state)).toEqual([
      'compiled',
      'building',
    ])
  })
  it('preserves the prior record and cleans temporary files when atomic replacement fails', () => {
    const { store, path } = setup()
    const record = store.beginProduction('project', 'doc', 'request', deck(), plan())
    vi.mocked(renameSync).mockImplementationOnce(() => {
      throw new Error('disk failure')
    })
    expect(() =>
      store.updateProductionPage(record, 'first', { state: 'building', attempt: 1 }),
    ).toThrow('disk failure')
    expect(store.production('project', 'doc', 'request')).toEqual(record)
    expect(readdirSync(join(path(), '..')).some((name) => name.endsWith('.tmp'))).toBe(false)
  })
  it('rejects repeated sequences, duplicate page identities and plan digest corruption on read', () => {
    const { store, path } = setup()
    store.beginProduction('project', 'doc', 'request', deck(), plan())
    store.beginProduction('project', 'doc', 'next', deck(), plan())
    const saved = readFileSync(path('next'), 'utf8')
    for (const corrupt of [
      { ...JSON.parse(saved), sequence: 1 },
      { ...JSON.parse(saved), planDigest: '0'.repeat(64) },
      {
        ...JSON.parse(saved),
        pages: [
          { pageId: 'first', state: 'pending', attempt: 0 },
          { pageId: 'first', state: 'pending', attempt: 0 },
        ],
      },
    ]) {
      writeFileSync(path('next'), JSON.stringify(corrupt))
      expect(() => store.production('project', 'doc')).toThrow('invalid_state')
      writeFileSync(path('next'), saved)
    }
  })
  it('rejects incompatible state fields and oversized or non-JSON reports', () => {
    const { store } = setup()
    let record = store.beginProduction('project', 'doc', 'request', deck(), plan())
    expect(() =>
      store.updateProductionPage(record, 'first', {
        state: 'building',
        attempt: 1,
        error: 'compile_failed',
      }),
    ).toThrow('invalid_state')
    record = store.updateProductionPage(record, 'first', { state: 'building', attempt: 1 })
    expect(() =>
      store.updateProductionPage(record, 'first', { state: 'failed', attempt: 1 }),
    ).toThrow('invalid_state')
    expect(() =>
      store.updateProductionPage(record, 'first', {
        state: 'failed',
        attempt: 1,
        error: '/private/error',
      }),
    ).toThrow('invalid_state')
    for (const report of ['x'.repeat(192 * 1024), { x: undefined }, { x: BigInt(1) }, new Date()])
      expect(() =>
        store.updateProductionPage(record, 'first', {
          state: 'compiled',
          attempt: 1,
          result: { ...result(), report },
        }),
      ).toThrow('invalid_state')
  })
})
