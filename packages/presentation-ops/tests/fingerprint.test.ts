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
})
