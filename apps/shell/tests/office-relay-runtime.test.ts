import { describe, expect, it, vi } from 'vitest'

import {
  logoutWithOfficeRelay,
  startOfficeRelayPersistence,
} from '../src/main/office-relay-runtime'

describe('Office relay runtime integration', () => {
  it('always performs auth logout after binding invalidation fails and reports the binding diagnostic', async () => {
    const order: string[] = []
    let suspended = false
    const diagnostic = vi.fn()
    const result = await logoutWithOfficeRelay({
      invalidate: async () => {
        suspended = true
        order.push('invalidate')
        throw new Error('disk_failure')
      },
      logout: async () => {
        expect(suspended).toBe(true)
        order.push('logout')
        return true
      },
      onBindingFailure: diagnostic,
    })
    expect(result).toBe(true)
    expect(order).toEqual(['invalidate', 'logout'])
    expect(diagnostic).toHaveBeenCalledOnce()
  })

  it('suspends persistent relay and lets bootstrap continue when startup storage sync fails', async () => {
    const suspend = vi.fn()
    const diagnostic = vi.fn()
    await expect(
      startOfficeRelayPersistence({
        sync: async () => {
          throw new Error('secure_storage_unavailable')
        },
        suspend,
        onBindingFailure: diagnostic,
      }),
    ).resolves.toBe(false)
    expect(suspend).toHaveBeenCalledWith('binding_lifecycle')
    expect(diagnostic).toHaveBeenCalledOnce()
  })

  it('continues bootstrap even when diagnostic shutdown itself throws', async () => {
    await expect(
      startOfficeRelayPersistence({
        sync: async () => {
          throw new Error('invalid_schema')
        },
        suspend: () => {
          throw new Error('partial_pool')
        },
        onBindingFailure: () => {
          throw new Error('renderer_gone')
        },
      }),
    ).resolves.toBe(false)
  })
})
