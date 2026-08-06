import { describe, expect, it, vi } from 'vitest'
import { CompileQueue } from '../src/queue.js'

describe('per-project compile queue', () => {
  it('single-flights the same project revision', async () => {
    const queue = new CompileQueue<string>()
    const run = vi.fn(async () => 'pdf-v1')
    const first = queue.request({ projectId: 'p', revision: 'r1', run })
    const second = queue.request({ projectId: 'p', revision: 'r1', run })
    expect(second).toBe(first)
    await expect(first).resolves.toBe('pdf-v1')
    expect(run).toHaveBeenCalledTimes(1)
  })

  it('cancels an old revision and never publishes its late result', async () => {
    const queue = new CompileQueue<string>()
    let finishOld!: (value: string) => void
    let oldSignal!: AbortSignal
    const oldRun = vi.fn(
      ({ signal }: { signal: AbortSignal }) =>
        new Promise<string>((resolve) => {
          oldSignal = signal
          signal.addEventListener('abort', () => undefined)
          finishOld = resolve
        }),
    )
    const publish = vi.fn()
    const old = queue.request({ projectId: 'p', revision: 'r1', run: oldRun, publish })
    const oldResult = expect(old).rejects.toMatchObject({ code: 'TECTONIC_STALE_RESULT' })
    await Promise.resolve()
    const fresh = queue.request({ projectId: 'p', revision: 'r2', run: async () => 'new', publish })
    expect(oldSignal.aborted).toBe(true)
    finishOld('old')
    await oldResult
    await expect(fresh).resolves.toBe('new')
    expect(publish).toHaveBeenCalledTimes(1)
    expect(publish).toHaveBeenCalledWith('new')
  })
  it('linearizes an in-progress publish before starting a newer revision', async () => {
    const queue = new CompileQueue<string>()
    let releasePublish!: () => void
    let publishStarted!: () => void
    const started = new Promise<void>((resolve) => (publishStarted = resolve))
    const gate = new Promise<void>((resolve) => (releasePublish = resolve))
    const events: string[] = []
    const old = queue.request({
      projectId: 'p',
      revision: 'r1',
      run: async () => 'old',
      publish: async () => {
        events.push('publish-old')
        publishStarted()
        await gate
      },
    })
    await started
    const fresh = queue.request({
      projectId: 'p',
      revision: 'r2',
      run: async () => {
        events.push('run-new')
        return 'new'
      },
      publish: (value) => {
        events.push(`publish-${value}`)
      },
    })
    expect(queue.cancel('p')).toBe(false)
    expect(events).toEqual(['publish-old'])
    releasePublish()
    await expect(old).resolves.toBe('old')
    await expect(fresh).resolves.toBe('new')
    expect(events).toEqual(['publish-old', 'run-new', 'publish-new'])
  })
})
