import { afterEach, describe, expect, it, vi } from 'vitest'
import { createTurnIdleDeadline } from '../src/main/codex-engine'

afterEach(() => vi.useRealTimers())

describe('Enhanced turn idle deadline', () => {
  it('expires stalled work promptly but extends the deadline after real progress', async () => {
    vi.useFakeTimers()
    const expired = vi.fn()
    const deadline = createTurnIdleDeadline(expired, 60_000)

    await vi.advanceTimersByTimeAsync(59_000)
    expect(expired).not.toHaveBeenCalled()
    deadline.touch()
    await vi.advanceTimersByTimeAsync(59_000)
    expect(expired).not.toHaveBeenCalled()
    await vi.advanceTimersByTimeAsync(1_000)
    expect(expired).toHaveBeenCalledOnce()
  })

  it('disarms the deadline after terminal settlement', async () => {
    vi.useFakeTimers()
    const expired = vi.fn()
    const deadline = createTurnIdleDeadline(expired, 60_000)
    deadline.disarm()
    await vi.advanceTimersByTimeAsync(60_000)
    expect(expired).not.toHaveBeenCalled()
  })
})
