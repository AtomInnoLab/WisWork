import { describe, expect, it, vi } from 'vitest'
import { createSaveCoordinator } from '../src/renderer/save-coordinator'

function deferred<T>() {
  let resolve!: (value: T) => void
  const promise = new Promise<T>((settle) => {
    resolve = settle
  })
  return { promise, resolve }
}

describe('save coordinator', () => {
  it('serializes save requests', async () => {
    const first = deferred<boolean>()
    const calls: string[] = []
    const coordinator = createSaveCoordinator()

    const firstResult = coordinator.enqueue(async () => {
      calls.push('first:start')
      const result = await first.promise
      calls.push('first:end')
      return result
    })
    const secondResult = coordinator.enqueue(async () => {
      calls.push('second')
      return true
    })

    await Promise.resolve()
    expect(calls).toEqual(['first:start'])
    first.resolve(true)
    await expect(firstResult).resolves.toBe(true)
    await expect(secondResult).resolves.toBe(true)
    expect(calls).toEqual(['first:start', 'first:end', 'second'])
  })

  it('flushes edits that arrive while an earlier save is in flight', async () => {
    const first = deferred<boolean>()
    let dirty = true
    let saves = 0
    const save = vi.fn(async () => {
      saves += 1
      if (saves === 1) {
        const result = await first.promise
        dirty = true
        return result
      }
      dirty = false
      return true
    })
    const coordinator = createSaveCoordinator()

    const inFlight = coordinator.enqueue(save)
    const close = coordinator.flushDirty(() => dirty, save)
    first.resolve(true)

    await expect(inFlight).resolves.toBe(true)
    await expect(close).resolves.toBe(true)
    expect(save).toHaveBeenCalledTimes(2)
    expect(dirty).toBe(false)
  })

  it('does not write again when the in-flight save already made the document clean', async () => {
    const first = deferred<boolean>()
    let dirty = true
    const save = vi.fn(async () => {
      const result = await first.promise
      dirty = false
      return result
    })
    const coordinator = createSaveCoordinator()

    const inFlight = coordinator.enqueue(save)
    const close = coordinator.flushDirty(() => dirty, save)
    first.resolve(true)

    await expect(inFlight).resolves.toBe(true)
    await expect(close).resolves.toBe(true)
    expect(save).toHaveBeenCalledTimes(1)
  })

  it('fails the close flush without retrying forever when a save fails', async () => {
    const save = vi.fn(async () => false)
    const coordinator = createSaveCoordinator()

    await expect(coordinator.flushDirty(() => true, save)).resolves.toBe(false)
    expect(save).toHaveBeenCalledTimes(1)
  })

  it('fails closed when edits keep arriving throughout the bounded close flush', async () => {
    const save = vi.fn(async () => true)
    const coordinator = createSaveCoordinator()

    await expect(coordinator.flushDirty(() => true, save)).resolves.toBe(false)
    expect(save).toHaveBeenCalledTimes(8)
  })
})
