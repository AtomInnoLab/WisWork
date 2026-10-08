import { describe, expect, it } from 'vitest'
import {
  selectDeterministicFailureDocument,
  shouldRetryRejectedToolInput,
} from '../src/main/codex-engine'

describe('deterministic stream failure turn correlation', () => {
  it('selects only the matching turn while multiple documents are active', () => {
    const first = { threadId: 'thread-a', active: { turnId: 'turn-a' } }
    const second = { threadId: 'thread-b', active: { turnId: 'turn-b' } }

    expect(selectDeterministicFailureDocument([first, second], 'turn-b')).toBe(second)
    expect(selectDeterministicFailureDocument([first, second], 'missing')).toBeUndefined()
    expect(selectDeterministicFailureDocument([first, second])).toBeUndefined()
  })

  it('attributes a provider turn-id mismatch when exactly one document is active', () => {
    const only = { threadId: 'thread-a', active: { turnId: 'local-turn-a' } }

    expect(selectDeterministicFailureDocument([only], 'provider-turn-a')).toBe(only)
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
