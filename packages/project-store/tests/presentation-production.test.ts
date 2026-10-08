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
  it('lists validated production requests newest first with isolated returned data', () => {
    const { store, path } = setup()
    expect(store.productionHistory('project', 'doc')).toEqual([])
    const first = store.beginProduction('project', 'doc', 'first', deck(), plan())
    const second = store.beginProduction('project', 'doc', 'second', deck(), plan())
    const history = store.productionHistory('project', 'doc')
    expect(history).toEqual([second, first])
    history[0]!.pages[0]!.attempt = 99
    ;(history[0]!.deck as ReturnType<typeof deck>).slides[0]!.title = 'mutated'
    expect(store.productionHistory('project', 'doc')).toEqual([second, first])
    expect(() => store.productionHistory('project', 'foreign')).toThrow('document_mismatch')
    const corrupted = JSON.parse(readFileSync(path('first'), 'utf8'))
    corrupted.inputDigest = 'a'.repeat(64)
    writeFileSync(path('first'), JSON.stringify(corrupted))
    expect(() => store.productionHistory('project', 'doc')).toThrow('invalid_state')
  })

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

function compiledParent(f: ReturnType<typeof setup>) {
  let parent = f.store.beginProduction('project', 'doc', 'request', deck(), plan())
  for (const page of parent.pages) {
    parent = f.store.updateProductionPage(parent, page.pageId, { state: 'building', attempt: 1 })
    parent = f.store.updateProductionPage(parent, page.pageId, {
      state: 'compiled',
      attempt: 1,
      result: result(),
    })
  }
  return parent
}
it('atomically derives only the requested page and preserves parent and copied results', () => {
  const f = setup(),
    parent = compiledParent(f),
    before = readFileSync(f.path(), 'utf8'),
    input = deck()
  input.slides[1]!.title = 'Rebuilt'
  const child = f.store.deriveProduction('project', 'doc', 'request', 'child', 'second', input)
  expect(child.revision).toEqual({
    parentRequestId: 'request',
    pageId: 'second',
    parentInputDigest: parent.inputDigest,
  })
  expect(child.pages).toEqual([parent.pages[0], { pageId: 'second', state: 'pending', attempt: 0 }])
  expect(child.plan).toEqual(parent.plan)
  expect(readFileSync(f.path(), 'utf8')).toBe(before)
  expect(new PresentationStore(f.root).production('project', 'doc', 'child')).toEqual(child)
  expect(f.store.deriveProduction('project', 'doc', 'request', 'child', 'second', input)).toEqual(
    child,
  )
  expect(() => f.store.beginProduction('project', 'doc', 'child', input, plan())).toThrow(
    'request_conflict',
  )
  expect(() =>
    f.store.deriveProduction('project', 'doc', 'request', 'child', 'first', deck()),
  ).toThrow('request_conflict')
  input.slides[1]!.title = 'Mutated'
  expect(f.store.production('project', 'doc', 'child')?.deck).toEqual({
    ...deck(),
    slides: [deck().slides[0], { id: 'second', title: 'Rebuilt' }],
  })
})
it('rejects incomplete parents, shared resource changes and non-target modifications', () => {
  const f = setup()
  f.store.beginProduction('project', 'doc', 'request', deck(), plan())
  expect(() =>
    f.store.deriveProduction('project', 'doc', 'request', 'child', 'second', deck()),
  ).toThrow('page_not_ready')
  compiledParent(f)
  for (const input of [
    { ...deck(), theme: 'different' },
    { ...deck(), slides: [{ id: 'first', title: 'changed' }, deck().slides[1]] },
    { ...deck(), slides: deck().slides.reverse() },
  ])
    expect(() =>
      f.store.deriveProduction('project', 'doc', 'request', 'child', 'second', input),
    ).toThrow('invalid_request')
  expect(() =>
    f.store.deriveProduction('project', 'doc', 'request', 'request', 'second', deck()),
  ).toThrow('invalid_request')
  expect(() =>
    f.store.deriveProduction('project', 'doc', 'request', 'child', 'missing', deck()),
  ).toThrow('invalid_request')
  expect(f.store.production('project', 'doc', 'child')).toBeUndefined()
})
it('keeps revision immutable during page CAS and validates parent digest on reload', () => {
  const f = setup()
  compiledParent(f)
  const child = f.store.deriveProduction('project', 'doc', 'request', 'child', 'second', deck())
  expect(() =>
    f.store.updateProductionPage(
      { ...child, revision: { ...child.revision!, parentInputDigest: 'a'.repeat(64) } },
      'second',
      { state: 'building', attempt: 1 },
    ),
  ).toThrow('invalid_state')
  const tampered = JSON.parse(readFileSync(f.path('child'), 'utf8'))
  tampered.revision.parentInputDigest = 'a'.repeat(64)
  writeFileSync(f.path('child'), JSON.stringify(tampered))
  expect(() => new PresentationStore(f.root).production('project', 'doc', 'child')).toThrow(
    'invalid_state',
  )
})
it('retries only the derived page and preserves inherited outputs through another derivation', () => {
  const f = setup(),
    parent = compiledParent(f)
  let child = f.store.deriveProduction('project', 'doc', 'request', 'child', 'second', deck())
  child = f.store.updateProductionPage(child, 'second', { state: 'building', attempt: 1 })
  child = f.store.updateProductionPage(child, 'second', {
    state: 'failed',
    attempt: 1,
    error: 'compile_failed',
  })
  expect(f.store.deriveProduction('project', 'doc', 'request', 'child', 'second', deck())).toEqual(
    child,
  )
  child = f.store.updateProductionPage(child, 'second', { state: 'building', attempt: 2 })
  child = f.store.updateProductionPage(child, 'second', {
    state: 'compiled',
    attempt: 2,
    result: { ...result(), pptxBase64: 'UEsDBAEAAAA=' },
  })
  expect(child.pages[0]).toEqual(parent.pages[0])
  expect(f.store.production('project', 'doc', 'request')).toEqual(parent)
  const next = f.store.deriveProduction('project', 'doc', 'child', 'grandchild', 'first', deck())
  expect(next.pages[1]).toEqual(child.pages[1])
  expect(new PresentationStore(f.root).production('project', 'doc', 'grandchild')).toEqual(next)
})
it('keeps derivation atomic, cleans failed writes and enforces the shared request cap', () => {
  const f = setup(),
    parent = compiledParent(f)
  vi.mocked(renameSync).mockImplementationOnce(() => {
    throw new Error('disk failure')
  })
  expect(() =>
    f.store.deriveProduction('project', 'doc', 'request', 'child', 'second', deck()),
  ).toThrow('disk failure')
  expect(f.store.production('project', 'doc', 'child')).toBeUndefined()
  expect(f.store.production('project', 'doc', 'request')).toEqual(parent)
  expect(readdirSync(join(f.path(), '..')).some((name) => name.endsWith('.tmp'))).toBe(false)
  for (let i = 0; i < 31; i++)
    f.store.deriveProduction('project', 'doc', 'request', `child-${i}`, 'second', deck())
  expect(() =>
    f.store.deriveProduction('project', 'doc', 'request', 'extra', 'second', deck()),
  ).toThrow('output_too_large')
  expect(
    f.store.deriveProduction('project', 'doc', 'request', 'child-0', 'second', deck()).requestId,
  ).toBe('child-0')
})
it('rejects malformed revision, missing parent and altered inherited page records', () => {
  const f = setup()
  compiledParent(f)
  const child = f.store.deriveProduction('project', 'doc', 'request', 'child', 'second', deck())
  for (const revision of [
    { ...child.revision!, extra: true },
    { ...child.revision!, parentRequestId: 'child' },
    { ...child.revision!, pageId: 'missing' },
    { ...child.revision!, parentRequestId: 'missing' },
  ]) {
    writeFileSync(f.path('child'), JSON.stringify({ ...child, revision }))
    expect(() => f.store.production('project', 'doc', 'child')).toThrow('invalid_state')
  }
  writeFileSync(
    f.path('child'),
    JSON.stringify({ ...child, pages: [{ ...child.pages[0], attempt: 2 }, child.pages[1]] }),
  )
  expect(() => f.store.production('project', 'doc', 'child')).toThrow('invalid_state')
})
