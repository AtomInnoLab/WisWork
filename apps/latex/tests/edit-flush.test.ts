import { describe, expect, it, vi } from 'vitest'
import { LatexEditFlushCoordinator } from '../src/main/edit-flush.js'
import {
  PendingUpdateRegistry,
  RendererCloseFreeze,
  flushRendererCloseFence,
} from '../src/renderer/workbench-coordination.js'
import { LATEX_CHANNELS } from '../src/shared/ipc.js'

function deferred<T>() {
  let resolve!: (value: T) => void
  const promise = new Promise<T>((done) => (resolve = done))
  return { promise, resolve }
}

describe('last-edit close fence', () => {
  it('waits for every tracked update and reports failures', async () => {
    const registry = new PendingUpdateRegistry()
    const first = deferred<boolean>()
    registry.track(first.promise)
    let settled = false
    const flush = registry.settleAll().then((ok) => {
      settled = true
      return ok
    })
    await Promise.resolve()
    expect(settled).toBe(false)
    first.resolve(true)
    await expect(flush).resolves.toBe(true)
    registry.track(Promise.resolve(false))
    await expect(registry.settleAll()).resolves.toBe(false)
    registry.track(Promise.reject(new Error('invoke failed')))
    await Promise.resolve()
    await Promise.resolve()
    await expect(registry.settleAll()).resolves.toBe(false)
    await expect(registry.settleAll()).resolves.toBe(false)
    registry.track(Promise.resolve(true))
    await expect(registry.settleAll()).resolves.toBe(true)
  })

  it('cancels timers and drains save queues including work created while flushing', async () => {
    const first = deferred<boolean>()
    const late = deferred<boolean>()
    const timers = new Map<string, number>([['main.tex', 1]])
    const queues = new Map<string, Promise<boolean>>([['main.tex', first.promise]])
    const clearTimer = vi.fn()
    const cancelAutoCompile = vi.fn()
    const flushing = flushRendererCloseFence({
      saveTimers: timers,
      saveQueues: queues,
      clearTimer,
      cancelAutoCompile,
      settleUpdates: async () => true,
    })
    await Promise.resolve()
    timers.set('late.tex', 2)
    queues.set('late.tex', late.promise)
    first.resolve(false)
    await Promise.resolve()
    queues.delete('main.tex')
    late.resolve(true)
    queues.delete('late.tex')
    await expect(flushing).resolves.toBe(true)
    expect(clearTimer).toHaveBeenCalledWith(1)
    expect(clearTimer).toHaveBeenCalledWith(2)
    expect(cancelAutoCompile).toHaveBeenCalled()
    expect(timers.size).toBe(0)
  })

  it('freezes synchronously through prepare and resumes only for the owning release', async () => {
    const freeze = new RendererCloseFreeze()
    const fence = deferred<boolean>()
    const preparing = freeze.prepare('r1', () => fence.promise)
    expect(freeze.isFrozen()).toBe(true)
    expect(freeze.release('wrong')).toBe(false)
    expect(freeze.isFrozen()).toBe(true)
    fence.resolve(true)
    await expect(preparing).resolves.toBe(true)
    expect(freeze.isFrozen()).toBe(true)
    expect(freeze.release('r1')).toBe(true)
    expect(freeze.isFrozen()).toBe(false)
  })

  it('accepts only the owning sender and matching one-shot request id', async () => {
    let ack!: (event: { sender: object }, payload: unknown) => void
    const ipcMain = { on: vi.fn((_channel, handler) => (ack = handler)), removeListener: vi.fn() }
    const coordinator = new LatexEditFlushCoordinator(ipcMain, {
      timeoutMs: 100,
      randomId: () => 'r1',
    })
    const destroyed = vi.fn()
    const contents = {
      id: 7,
      isDestroyed: () => false,
      send: vi.fn(),
      once: vi.fn((_e, h) => destroyed.mockImplementation(h)),
      removeListener: vi.fn(),
    }
    const pending = coordinator.request(contents)
    const joined = coordinator.request(contents)
    expect(contents.send).toHaveBeenCalledWith(LATEX_CHANNELS.editFlushRequest, {
      requestId: 'r1',
    })
    ack({ sender: {} }, { requestId: 'r1', ok: true })
    ack({ sender: contents }, { requestId: 'wrong', ok: true })
    let done = false
    void pending.then(() => (done = true))
    await Promise.resolve()
    expect(done).toBe(false)
    ack({ sender: contents }, { requestId: 'r1', ok: true })
    await expect(pending).resolves.toBe(true)
    await expect(joined).resolves.toBe(true)
    expect(contents.send).not.toHaveBeenCalledWith(LATEX_CHANNELS.editFlushRelease, {
      requestId: 'r1',
    })
    coordinator.release(contents)
    expect(contents.send).not.toHaveBeenCalledWith(LATEX_CHANNELS.editFlushRelease, {
      requestId: 'r1',
    })
    coordinator.release(contents)
    expect(contents.send).toHaveBeenCalledWith(LATEX_CHANNELS.editFlushRelease, {
      requestId: 'r1',
    })
    coordinator.release(contents)
    expect(
      contents.send.mock.calls.filter(([channel]) => channel === LATEX_CHANNELS.editFlushRelease),
    ).toHaveLength(1)
    ack({ sender: contents }, { requestId: 'r1', ok: false })
    coordinator.dispose()
  })

  it('blocks on timeout and renderer destruction', async () => {
    vi.useFakeTimers()
    let destroyed!: () => void
    const coordinator = new LatexEditFlushCoordinator(
      { on: vi.fn(), removeListener: vi.fn() },
      { timeoutMs: 10 },
    )
    const contents = {
      id: 8,
      isDestroyed: () => false,
      send: vi.fn(),
      once: vi.fn((_e, h) => (destroyed = h)),
      removeListener: vi.fn(),
    }
    const timedOut = coordinator.request(contents)
    const joinedTimeout = coordinator.request(contents)
    await vi.advanceTimersByTimeAsync(10)
    await expect(timedOut).resolves.toBe(false)
    await expect(joinedTimeout).resolves.toBe(false)
    expect(
      contents.send.mock.calls.some(([channel]) => channel === LATEX_CHANNELS.editFlushRelease),
    ).toBe(true)
    const destroyedResult = coordinator.request(contents)
    destroyed()
    await expect(destroyedResult).resolves.toBe(false)
    coordinator.dispose()
    vi.useRealTimers()
  })
})
