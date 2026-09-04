import { describe, expect, it } from 'vitest'
import type { Session } from '../src/main/session-state'
import { blocksCanonicalPresentationTransaction } from '../src/main/operations/transaction-admission'

const session = (overrides: Partial<Session> = {}): Session =>
  ({
    path: '',
    opened: {} as Session['opened'],
    fitWidthPx: 960,
    undoStack: [],
    redoStack: [],
    ...overrides,
  }) as Session

describe('canonical presentation transaction admission', () => {
  it('allows a canonical content transaction inside the agent history batch', () => {
    expect(
      blocksCanonicalPresentationTransaction(
        session({
          historyBatch: {
            depth: 1,
            undoStart: 0,
            before: {} as NonNullable<Session['historyBatch']>['before'],
          },
        }),
      ),
    ).toBe(false)
  })

  it('continues to block master editing and transform previews', () => {
    expect(blocksCanonicalPresentationTransaction(session({ masterEdit: {} as never }))).toBe(true)
    expect(blocksCanonicalPresentationTransaction(session({ transformPreview: true }))).toBe(true)
  })
})
