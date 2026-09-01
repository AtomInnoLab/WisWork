import { describe, expect, it, vi } from 'vitest'
import { createBeforeQuitBarrier } from '../src/main/before-quit-barrier'

describe('before-quit cleanup barrier', () => {
  it('defers quit, coalesces repeated events and quits once after cleanup', async () => {
    let release!: () => void
    const cleanup = vi.fn(() => new Promise<void>((resolve) => (release = resolve)))
    const quit = vi.fn()
    const preventDefault = vi.fn()
    const handle = createBeforeQuitBarrier({ cleanup, quit })
    handle({ preventDefault })
    handle({ preventDefault })
    expect(cleanup).toHaveBeenCalledOnce()
    expect(quit).not.toHaveBeenCalled()
    release()
    await vi.waitFor(() => expect(quit).toHaveBeenCalledOnce())
  })

  it('fails safe at the cleanup deadline by keeping the app alive', async () => {
    vi.useFakeTimers()
    const quit = vi.fn()
    const diagnostics = vi.fn()
    const handle = createBeforeQuitBarrier({
      cleanup: async () => await new Promise<void>(() => undefined),
      quit,
      diagnostics,
      deadlineMs: 10,
    })
    handle({ preventDefault: vi.fn() })
    await vi.advanceTimersByTimeAsync(10)
    expect(diagnostics).toHaveBeenCalledWith('enhanced_quit_cleanup_deadline')
    expect(quit).not.toHaveBeenCalled()
    vi.useRealTimers()
  })

  it('continues awaiting cleanup after the warning deadline and then quits once', async () => {
    vi.useFakeTimers()
    let release!: () => void
    const quit = vi.fn()
    const handle = createBeforeQuitBarrier({
      cleanup: () => new Promise<void>((resolve) => (release = resolve)),
      quit,
      deadlineMs: 10,
    })
    const preventDefault = vi.fn()
    handle({ preventDefault })
    await vi.advanceTimersByTimeAsync(10)
    expect(quit).not.toHaveBeenCalled()
    handle({ preventDefault })
    release()
    await vi.runAllTimersAsync()
    expect(preventDefault).toHaveBeenCalledTimes(2)
    expect(quit).toHaveBeenCalledOnce()
    vi.useRealTimers()
  })
})
