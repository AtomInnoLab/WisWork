import { describe, expect, it } from 'vitest'
import {
  selectDeterministicFailureTurn,
  shouldRetryRejectedToolInput,
} from '../src/main/codex-engine'

describe('deterministic stream failure turn correlation', () => {
  it('selects only the matching turn while multiple documents are active', () => {
    const first = { turnId: 'turn-a' }
    const second = { turnId: 'turn-b' }

    expect(selectDeterministicFailureTurn([first, second], 'turn-b')).toBe(second)
    expect(selectDeterministicFailureTurn([first, second], 'missing')).toBeUndefined()
    expect(selectDeterministicFailureTurn([first, second])).toBeUndefined()
  })
})

describe('malformed tool input recovery', () => {
  it('retries rejected exec input once without retrying wait or repeated input', () => {
    expect(shouldRetryRejectedToolInput('invalid_custom_tool_input', 0)).toBe(true)
    expect(shouldRetryRejectedToolInput('unsafe_custom_tool_input', 0)).toBe(true)
    expect(shouldRetryRejectedToolInput('invalid_custom_tool_input', 1)).toBe(false)
    expect(shouldRetryRejectedToolInput('unsafe_custom_tool_input', 1)).toBe(false)
    expect(shouldRetryRejectedToolInput('invalid_wait_input', 0)).toBe(false)
  })
})
