import { describe, expect, it } from 'vitest'

import { fingerprintSemanticValue } from '../src/index'

describe('semantic fingerprint', () => {
  it('is deterministic across object insertion order', async () => {
    const left = await fingerprintSemanticValue({
      text: 'café',
      geometry: { y: 2, x: 1 },
      tags: ['a', 'b'],
    })
    const right = await fingerprintSemanticValue({
      tags: ['a', 'b'],
      geometry: { x: 1, y: 2 },
      text: 'café',
    })
    expect(left).toBe(right)
    expect(left).toMatch(/^sha256:[0-9a-f]{64}$/)
  })

  it('preserves array order and semantic type distinctions', async () => {
    await expect(fingerprintSemanticValue(['a', 'b'])).resolves.not.toBe(
      await fingerprintSemanticValue(['b', 'a']),
    )
    await expect(fingerprintSemanticValue('1')).resolves.not.toBe(await fingerprintSemanticValue(1))
  })

  it.each([NaN, Infinity, -Infinity, undefined, { nested: undefined }])(
    'rejects unsupported value %s',
    async (value) => {
      await expect(fingerprintSemanticValue(value)).rejects.toThrow()
    },
  )

  it('rejects cycles and unsafe prototype keys', async () => {
    const cyclic: Record<string, unknown> = {}
    cyclic.self = cyclic
    await expect(fingerprintSemanticValue(cyclic)).rejects.toThrow(/cyclic/i)
    await expect(fingerprintSemanticValue(JSON.parse('{"__proto__":true}'))).rejects.toThrow(
      /unsafe/i,
    )
  })

  it('rejects hidden symbol and accessor state instead of silently omitting it', async () => {
    await expect(fingerprintSemanticValue({ [Symbol('hidden')]: true })).rejects.toThrow()
    const accessor = {}
    Object.defineProperty(accessor, 'secret', { enumerable: true, get: () => 'value' })
    await expect(fingerprintSemanticValue(accessor)).rejects.toThrow(/accessor/i)
  })

  it('rejects sparse arrays so holes cannot collide with empty arrays', async () => {
    const sparse = Array(1)
    await expect(fingerprintSemanticValue(sparse)).rejects.toThrow(/array/i)
  })

  it('rejects array extensions and non-standard prototypes', async () => {
    const values: unknown[][] = [[], [], [], [], []]
    Object.assign(values[0]!, { extra: true })
    Object.defineProperty(values[1]!, 'hidden', { value: true })
    values[2]![Symbol('hidden') as unknown as number] = true
    Object.defineProperty(values[3]!, '__proto__', { value: true, enumerable: true })
    Object.setPrototypeOf(values[4]!, null)
    for (const value of values) await expect(fingerprintSemanticValue(value)).rejects.toThrow()
  })

  it('rejects array accessors without executing them', async () => {
    let accessed = false
    const value: unknown[] = []
    Object.defineProperty(value, '0', {
      enumerable: true,
      get: () => {
        accessed = true
        return 'secret'
      },
    })
    await expect(fingerprintSemanticValue(value)).rejects.toThrow(/accessor/i)
    expect(accessed).toBe(false)
  })
})
