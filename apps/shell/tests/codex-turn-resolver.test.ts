import { describe, expect, it, vi } from 'vitest'
import { CodexTurnResolver } from '../src/main/codex-turn-resolver'

describe('Shell-private Codex turn resolver', () => {
  it('binds sequential provider requests to one turn and rejects drift/cross-thread', () => {
    const prepareTurn = vi.fn(() => ({ messagesRequest: {}, messagesStreamToResponses: vi.fn() }))
    const issueForTurn = vi.fn(() => Object.freeze({}))
    const resolver = new CodexTurnResolver()
    resolver.arm('thread-a', () => ({ issueForTurn, prepareTurn }) as never, Object.freeze({}))
    const input = {
      client_metadata: { thread_id: 'thread-a', turn_id: 'turn-a', session_id: 'session-a' },
    }
    expect(resolver.prepare(input)).toBeTruthy()
    expect(issueForTurn).toHaveBeenCalledWith(
      expect.objectContaining({
        turnId: 'turn-a',
        method: 'mcp__wiswork__wiswork_call',
        toolName: 'wiswork_call',
      }),
    )
    expect(resolver.prepare(input)).toBeTruthy()
    expect(issueForTurn).toHaveBeenCalledTimes(2)
    expect(() =>
      resolver.prepare({
        client_metadata: { thread_id: 'thread-a', turn_id: 'turn-b', session_id: 'session-a' },
      }),
    ).toThrow('turn_unbound')
    expect(() =>
      resolver.prepare({
        client_metadata: { thread_id: 'thread-b', turn_id: 'turn-a', session_id: 'session-a' },
      }),
    ).toThrow('turn_unbound')
    const disarm = resolver.arm(
      'thread-c',
      () => ({ issueForTurn, prepareTurn }) as never,
      Object.freeze({}),
    )
    disarm()
    expect(() =>
      resolver.prepare({
        client_metadata: { thread_id: 'thread-c', turn_id: 'turn-c', session_id: 'session-c' },
      }),
    ).toThrow('turn_unbound')
  })
})
