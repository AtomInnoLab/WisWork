import { describe, expect, it, vi } from 'vitest'

import { readUntilConverged } from '../src/skills/shared/office-write-transaction.js'

describe('Office write transaction convergence', () => {
  it('accepts a semantic state that becomes visible after two stale readbacks', async () => {
    const read = vi
      .fn()
      .mockResolvedValueOnce('before')
      .mockResolvedValueOnce('before')
      .mockResolvedValue('after')
    const delay = vi.fn().mockResolvedValue(undefined)

    await expect(
      readUntilConverged({ read, accept: (value) => value === 'after', delay }),
    ).resolves.toBe('after')
    expect(read).toHaveBeenCalledTimes(3)
    expect(delay).toHaveBeenCalledTimes(2)
  })

  it('aborts rather than performing another delayed readback', async () => {
    const controller = new AbortController()
    const read = vi.fn().mockResolvedValue('before')
    const delay = vi.fn(async () => controller.abort())

    await expect(
      readUntilConverged({ read, accept: () => false, delay, signal: controller.signal }),
    ).rejects.toThrow('cancelled')
    expect(read).toHaveBeenCalledOnce()
  })

  it('releases the abort listener after a completed retry delay', async () => {
    vi.useFakeTimers()
    try {
      const controller = new AbortController()
      const add = vi.spyOn(controller.signal, 'addEventListener')
      const remove = vi.spyOn(controller.signal, 'removeEventListener')
      const read = vi.fn().mockResolvedValueOnce('before').mockResolvedValue('after')
      const result = readUntilConverged({
        read,
        accept: (value) => value === 'after',
        signal: controller.signal,
      })
      await vi.advanceTimersByTimeAsync(50)
      await expect(result).resolves.toBe('after')
      expect(add).toHaveBeenCalledOnce()
      expect(remove).toHaveBeenCalledWith('abort', expect.any(Function))
    } finally {
      vi.useRealTimers()
    }
  })

  it('releases the abort listener when the retry is cancelled', async () => {
    vi.useFakeTimers()
    try {
      const controller = new AbortController()
      const remove = vi.spyOn(controller.signal, 'removeEventListener')
      const read = vi.fn().mockResolvedValue('before')
      const result = readUntilConverged({
        read,
        accept: () => false,
        signal: controller.signal,
      })
      await Promise.resolve()
      controller.abort()
      await expect(result).rejects.toThrow('cancelled')
      expect(read).toHaveBeenCalledOnce()
      expect(remove).toHaveBeenCalledWith('abort', expect.any(Function))
    } finally {
      vi.useRealTimers()
    }
  })
})
