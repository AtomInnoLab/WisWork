import { describe, expect, it } from 'vitest'
import { selectDeterministicFailureTurn } from '../src/main/codex-engine'

describe('deterministic stream failure turn correlation', () => {
  it('selects only the matching turn while multiple documents are active', () => {
    const first = { turnId: 'turn-a' }
    const second = { turnId: 'turn-b' }

    expect(selectDeterministicFailureTurn([first, second], 'turn-b')).toBe(second)
    expect(selectDeterministicFailureTurn([first, second], 'missing')).toBeUndefined()
    expect(selectDeterministicFailureTurn([first, second])).toBeUndefined()
  })
})
