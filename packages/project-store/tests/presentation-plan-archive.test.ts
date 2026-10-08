import {
  mkdtempSync,
  rmSync,
  readFileSync,
  writeFileSync,
  readdirSync,
  unlinkSync,
  symlinkSync,
  renameSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { PresentationStore } from '../src/presentation-store.js'
vi.mock('node:fs', async (original) => {
  const fs = await original<typeof import('node:fs')>()
  return { ...fs, renameSync: vi.fn(fs.renameSync) }
})
const roots: string[] = []
afterEach(async () => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
  const actual = await vi.importActual<typeof import('node:fs')>('node:fs')
  vi.mocked(renameSync).mockReset().mockImplementation(actual.renameSync)
})
function setup() {
  const root = mkdtempSync(join(tmpdir(), 'ppt-plan-archive-'))
  roots.push(root)
  const store = new PresentationStore(root)
  const a = store.savePlan('p', 'd', 0, { title: 'A', sources: ['research'], assets: ['image'] })
  const directory = join(
    root,
    'projects',
    'presentations',
    readdirSync(join(root, 'projects', 'presentations'))[0]!,
  )
  return {
    root,
    store,
    a,
    path: (revision: number) => join(directory, `plan-revision-${revision}.json`),
    directory,
  }
}
describe('complete historical presentation plans', () => {
  it('archives the last complete pre-history plan on its first update without inventing earlier revisions', () => {
    const f = setup()
    const { revisions: _history, ...legacy } = f.a
    writeFileSync(join(f.directory, 'plan.json'), JSON.stringify({ ...legacy, revision: 7 }))
    const next = f.store.savePlan('p', 'd', 7, { title: 'updated legacy' })
    expect(next.revisions?.map((event) => event.revision)).toEqual([7, 8])
    expect(f.store.planRevision('p', 'd', 7)?.plan).toEqual(f.a.plan)
    expect(f.store.planRevision('p', 'd', 6)).toBeUndefined()
  })
  it('recovers exact historical content after reopening without changing the current revision', () => {
    const f = setup()
    f.store.savePlan('p', 'd', 1, { title: 'B' })
    const c = f.store.savePlan('p', 'd', 2, { title: 'C' })
    const reopened = new PresentationStore(f.root)
    expect(reopened.planRevision('p', 'd', 1)?.plan).toEqual(f.a.plan)
    expect(reopened.planRevision('p', 'd', 2)?.plan).toEqual({ title: 'B' })
    expect(reopened.planRevision('p', 'd', 3)).toEqual(c)
    expect(reopened.plan('p', 'd')).toEqual(c)
    const old = reopened.planRevision('p', 'd', 1)!
    ;(old.plan as { title: string }).title = 'local mutation'
    expect(reopened.planRevision('p', 'd', 1)?.plan).toEqual(f.a.plan)
    expect(() => reopened.planRevision('p', 'foreign', 1)).toThrow('document_mismatch')
    expect(() => reopened.planRevision('p', 'd', 0)).toThrow('invalid_request')
  })

  it('does not advance the current plan if archival or current commit fails', async () => {
    const f = setup()
    vi.mocked(renameSync).mockImplementationOnce(() => {
      throw new Error('archive failure')
    })
    expect(() => f.store.savePlan('p', 'd', 1, { title: 'B' })).toThrow('archive failure')
    expect(f.store.plan('p', 'd')).toEqual(f.a)
    const actual = await vi.importActual<typeof import('node:fs')>('node:fs')
    vi.mocked(renameSync)
      .mockImplementationOnce(actual.renameSync)
      .mockImplementationOnce(() => {
        throw new Error('current failure')
      })
    expect(() => f.store.savePlan('p', 'd', 1, { title: 'B' })).toThrow('current failure')
    expect(f.store.plan('p', 'd')).toEqual(f.a)
    f.store.savePlan('p', 'd', 1, { title: 'B' })
    expect(f.store.planRevision('p', 'd', 1)?.plan).toEqual(f.a.plan)
  })

  it('rejects altered archive content and symbolic links rather than falling back to other data', () => {
    const f = setup()
    f.store.savePlan('p', 'd', 1, { title: 'B' })
    const saved = readFileSync(f.path(1), 'utf8')
    const forged = JSON.parse(saved)
    forged.plan.title = 'forged'
    writeFileSync(f.path(1), JSON.stringify(forged))
    expect(() => f.store.planRevision('p', 'd', 1)).toThrow('invalid_state')
    unlinkSync(f.path(1))
    symlinkSync(`${f.path(1)}.missing`, f.path(1))
    expect(() => f.store.planRevision('p', 'd', 1)).toThrow('invalid_state')
    expect(f.store.plan('p', 'd')?.revision).toBe(2)
  })

  it('retains only the current 32-revision window and rejects historical versions outside it', () => {
    const f = setup()
    for (let revision = 1; revision < 35; revision++)
      f.store.savePlan('p', 'd', revision, { title: String(revision) })
    expect(f.store.plan('p', 'd')?.revision).toBe(35)
    expect(
      readdirSync(f.directory).filter((file) => /^plan-revision-\d+\.json$/.test(file)),
    ).toHaveLength(31)
    expect(f.store.planRevision('p', 'd', 3)).toBeUndefined()
    expect(f.store.planRevision('p', 'd', 4)?.revision).toBe(4)
  })

  it('recovers a missing old archive only from an exact frozen plan binding', () => {
    const f = setup()
    f.store.beginProduction(
      'p',
      'd',
      'old-task',
      { slides: [{ id: 'a' }] },
      { revision: 1, plan: f.a.plan },
    )
    f.store.savePlan('p', 'd', 1, { title: 'B' })
    unlinkSync(f.path(1))
    expect(f.store.planRevision('p', 'd', 1)?.plan).toEqual(f.a.plan)
    expect(f.store.planRevision('p', 'd', 999)).toBeUndefined()
    const other = setup()
    other.store.savePlan('p', 'd', 1, { title: 'B' })
    unlinkSync(other.path(1))
    expect(other.store.planRevision('p', 'd', 1)).toBeUndefined()
  })
})
