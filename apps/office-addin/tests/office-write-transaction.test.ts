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
})
