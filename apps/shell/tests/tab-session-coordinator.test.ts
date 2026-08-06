import { readFileSync } from 'node:fs'
import { describe, expect, it, vi } from 'vitest'
import { TabSessionPersistenceCoordinator } from '../src/main/tab-session'

function deferred() {
  let resolve!: () => void
  const promise = new Promise<void>((done) => (resolve = done))
  return { promise, resolve }
}

describe('tab session persistence coordinator', () => {
  it('serializes immutable snapshots and final flush waits for the last write', async () => {
    const first = deferred()
    const writes: string[][] = []
    const save = vi.fn(async (state: { projectPaths: string[] }) => {
      writes.push(state.projectPaths)
      if (writes.length === 1) await first.promise
    })
    const coordinator = new TabSessionPersistenceCoordinator(save)
    const mutable = { projectPaths: ['/one'], activeProjectPath: '/one' }
    coordinator.enqueue(mutable)
    mutable.projectPaths.push('/mutated')
    coordinator.enqueue({ projectPaths: ['/two'], activeProjectPath: '/two' })
    const final = coordinator.flush({ projectPaths: ['/final'], activeProjectPath: '/final' })
    await vi.waitFor(() => expect(writes).toEqual([['/one']]))
    coordinator.enqueue({ projectPaths: ['/late'], activeProjectPath: '/late' })
    first.resolve()
    await expect(final).resolves.toBe(true)
    expect(writes).toEqual([['/one'], ['/two'], ['/final'], ['/late']])
  })

  it('reports a failed final save so close can retain the window', async () => {
    const coordinator = new TabSessionPersistenceCoordinator(async () => {
      throw new Error('disk full')
    })
    await expect(
      coordinator.flush({ projectPaths: ['/final'], activeProjectPath: '/final' }),
    ).resolves.toBe(false)
  })

  it('wires whole-window close through edit flush and final session drain', () => {
    const source = readFileSync(new URL('../src/main/index.ts', import.meta.url), 'utf8')
    const start = source.indexOf("win.on('close'")
    const end = source.indexOf("win.on('closed'", start)
    const closeHandler = source.slice(start, end)
    expect(start).toBeGreaterThan(-1)
    expect(closeHandler.indexOf('event.preventDefault()')).toBeLessThan(
      closeHandler.indexOf('manager.dirtySheetsTabs()'),
    )
    expect(closeHandler.indexOf('prepareLatexCloseTabs')).toBeGreaterThan(-1)
    expect(closeHandler.indexOf('prepareLatexCloseTabs')).toBeLessThan(
      closeHandler.indexOf('manager.dirtySheetsTabs()'),
    )
    expect(closeHandler.indexOf('prepareLatexCloseTabs')).toBeLessThan(
      closeHandler.indexOf('latexQueryDirty'),
    )
    expect(closeHandler.indexOf('tabSessionPersistence.flush')).toBeLessThan(
      closeHandler.indexOf('finalLatexCloseCheck'),
    )
    expect(closeHandler.indexOf('finalLatexCloseCheck')).toBeLessThan(
      closeHandler.indexOf('closeConfirmed = true'),
    )
    expect(closeHandler).toContain('releaseLatexCloseTabs(latexTabs)')
    expect(closeHandler).toContain('closeInProgress')
  })
})
