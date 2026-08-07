import { describe, expect, it, vi } from 'vitest'
import {
  finalLatexCloseCheck,
  prepareLatexCloseTabs,
  releaseLatexCloseTabs,
} from '../src/main/latex-final-close'

describe('whole-window final LaTeX close pass', () => {
  it('prepares every LaTeX tab concurrently and resumes all after an abort', async () => {
    let finishSecond!: (ok: boolean) => void
    const second = new Promise<boolean>((resolve) => (finishSecond = resolve))
    const frozen = new Set<object>()
    const tabs = [{ webContents: {} as never }, { webContents: {} as never }]
    const prepare = vi.fn((contents: object) => {
      frozen.add(contents)
      return contents === tabs[1].webContents ? second : Promise.resolve(true)
    })
    const preparing = prepareLatexCloseTabs(tabs, prepare)
    expect(prepare).toHaveBeenCalledTimes(2)
    expect(frozen.size).toBe(2)
    finishSecond(false)
    await expect(preparing).resolves.toBe(false)
    releaseLatexCloseTabs(tabs, (contents) => frozen.delete(contents))
    expect(frozen.size).toBe(0)
  })
  it('aborts when a late edit becomes dirty during deferred close work', async () => {
    let release!: () => void
    const deferred = new Promise<void>((resolve) => (release = resolve))
    let lateDirty = false
    const dirty = vi.fn(async () => {
      await deferred
      return lateDirty
    })
    const checking = finalLatexCloseCheck([{ webContents: {} as never }], dirty)
    lateDirty = true
    release()
    await expect(checking).resolves.toBe(false)
    expect(dirty).toHaveBeenCalledOnce()
  })

  it('allows close only when the final flush succeeds and remains clean', async () => {
    const dirty = vi.fn().mockResolvedValue(false)
    await expect(finalLatexCloseCheck([{ webContents: {} as never }], dirty)).resolves.toBe(true)
  })
})
