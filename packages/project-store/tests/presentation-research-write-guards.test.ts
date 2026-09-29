import { afterEach, expect, it, vi } from 'vitest'
import { mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PresentationResearchStore } from '../src/presentation-research-store'
import { PresentationLifecycleStore } from '../src/presentation-lifecycle'
const roots: string[] = []
afterEach(() => {
  vi.restoreAllMocks()
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'research-fence-'))
  roots.push(root)
  const store = new PresentationResearchStore(root),
    lifecycle = new PresentationLifecycleStore(root),
    scope = { documentId: 'doc', projectId: 'p' }
  lifecycle.initialize(scope)
  return { root, store, lifecycle, scope, guard: () => lifecycle.assertActive(scope, 0) }
}
const draft = { scope: 'Research', sources: [], facts: [] }
const freeze = (f: ReturnType<typeof fixture>) =>
  f.lifecycle.beginDeletion(f.scope, 0, {
    deletionId: 'delete',
    reason: 'user',
    resources: [{ resourceId: 'research', kind: 'research', ownership: 'project_exclusive' }],
  })
it('refuses a frozen begin before creating the research namespace', async () => {
  const f = fixture()
  freeze(f)
  await expect(f.store.begin('doc', 'p', 0, 'first', draft, f.guard)).rejects.toThrow(
    'revision_conflict',
  )
  expect(readdirSync(f.root)).toEqual(['presentation-project-lifecycles'])
})
it('rechecks the original lease after a private load await and cannot publish finished body', async () => {
  const f = fixture()
  await f.store.begin('doc', 'p', 0, 'first', draft)
  const before = await f.store.read('doc', 'p', 'first')
  const privateStore = f.store as unknown as {
    load(doc: string, project: string): Promise<unknown>
  }
  const load = privateStore.load.bind(f.store)
  let arrived!: () => void, resume!: () => void
  const paused = new Promise<void>((r) => {
      arrived = r
    }),
    wait = new Promise<void>((r) => {
      resume = r
    })
  vi.spyOn(privateStore, 'load').mockImplementation(async (...args) => {
    const value = await load(...args)
    arrived()
    await wait
    return value
  })
  const pending = f.store.finish('doc', 'p', 'first', { state: 'completed', sources: [] }, f.guard)
  await paused
  freeze(f)
  resume()
  await expect(pending).rejects.toThrow('revision_conflict')
  vi.restoreAllMocks()
  expect(await new PresentationResearchStore(f.root).read('doc', 'p', 'first')).toEqual(before)
})
it('does not remove an unowned staging file while publishing its own next revision', async () => {
  const f = fixture()
  await f.store.begin('doc', 'p', 0, 'first', draft)
  const folder = join(
    f.root,
    'presentation-research',
    readdirSync(join(f.root, 'presentation-research'))[0]!,
    readdirSync(
      join(f.root, 'presentation-research', readdirSync(join(f.root, 'presentation-research'))[0]!),
    )[0]!,
  )
  const foreign = join(folder, 'state.json.12345678-1234-4123-8123-123456789abc.tmp')
  writeFileSync(foreign, 'other process staging')
  await f.store.finish('doc', 'p', 'first', { state: 'completed', sources: [] }, f.guard)
  expect(readFileSync(foreign, 'utf8')).toBe('other process staging')
})
