import { describe, expect, it } from 'vitest'
import { resolveIterationIdentity } from '../src/main/iteration-identity'

describe('packaged iteration identity', () => {
  it('leaves production unchanged only when metadata is absent', () => {
    expect(resolveIterationIdentity(undefined)).toBeNull()
    expect(() => resolveIterationIdentity({ mode: 'unknown' })).toThrow()
  })
  it('derives separate deterministic profiles', () => {
    const meta = { commit: 'abc1234', builtAt: '2026-09-03T00:00:00.000Z' }
    expect(resolveIterationIdentity({ ...meta, mode: 'dogfood' })?.productName).toBe(
      'WisWork Dogfood',
    )
    expect(resolveIterationIdentity({ ...meta, mode: 'preview', pr: 123 })?.productName).toBe(
      'WisWork Preview PR123',
    )
    expect(() => resolveIterationIdentity({ ...meta, mode: 'preview', pr: '../WisWork' })).toThrow()
  })
})
