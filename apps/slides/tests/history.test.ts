import { describe, expect, it, vi } from 'vitest'
import type { Session } from '../src/main/session-state'
import {
  beginHistoryBatch,
  acquirePresentationMutationLease,
  acquirePresentationPersistenceLease,
  acquirePresentationTransactionLease,
  endHistoryBatch,
  pushHistory,
  registerAiSnapshot,
  restoreAiSnapshot,
  restoreSnapshot,
  redoSession,
  settleStaleHistoryBatch,
  sessionHasActivePresentationTransaction,
  sessionBlocksPresentationClose,
  SlidesSessionBusyError,
  undoSession,
  withPresentationMutationLease,
} from '../src/main/session-state'

vi.mock('electron', () => ({
  BrowserWindow: { getFocusedWindow: () => null },
}))
vi.mock('../src/main/fonts', () => ({
  createSystemFontMetrics: () => ({}),
}))

function sessionWith(value: string): Session {
  return {
    path: '',
    fitWidthPx: 1280,
    undoStack: [],
    redoStack: [],
    opened: {
      deck: {
        slides: [{ value }],
        size: { cx: 1, cy: 1 },
      },
      archive: { entries: new Map([['deck', new Uint8Array([value.length])]]) },
    },
  } as unknown as Session
}

function valueOf(session: Session): string {
  return (session.opened.deck.slides[0] as unknown as { value: string }).value
}

function setValue(session: Session, value: string): void {
  ;(session.opened.deck.slides[0] as unknown as { value: string }).value = value
}

describe('Slides main-process history batching', () => {
  it('makes transaction and persistence leases mutually exclusive', () => {
    const session = sessionWith('before')
    const releaseTransaction = acquirePresentationTransactionLease(session)
    expect(releaseTransaction).not.toBeNull()
    expect(sessionHasActivePresentationTransaction(session)).toBe(true)
    expect(acquirePresentationTransactionLease(session)).toBeNull()
    expect(acquirePresentationPersistenceLease(session)).toBeNull()
    releaseTransaction!()

    const releasePersistence = acquirePresentationPersistenceLease(session)
    expect(releasePersistence).not.toBeNull()
    expect(sessionBlocksPresentationClose(session)).toBe(true)
    expect(acquirePresentationPersistenceLease(session)).toBeNull()
    expect(acquirePresentationTransactionLease(session)).toBeNull()
    expect(acquirePresentationMutationLease(session)).toBeNull()
    releasePersistence!()
    const releaseMutation = acquirePresentationMutationLease(session)
    expect(releaseMutation).not.toBeNull()
    expect(acquirePresentationTransactionLease(session)).toBeNull()
    expect(acquirePresentationPersistenceLease(session)).toBeNull()
    expect(sessionBlocksPresentationClose(session)).toBe(true)
    releaseMutation!()
    expect(acquirePresentationTransactionLease(session)).not.toBeNull()
  })

  it('blocks ordinary edit, undo, redo, and snapshot restore while a transaction owns the gate', () => {
    const session = sessionWith('before')
    pushHistory(session)
    setValue(session, 'after')
    const snapshotId = registerAiSnapshot(session, session.undoStack[0]!)
    const undoBefore = [...session.undoStack]
    const redoBefore = [...session.redoStack]
    const release = acquirePresentationTransactionLease(session)!

    const expectBusy = (action: () => unknown) => {
      try {
        action()
        throw new Error('expected busy')
      } catch (error) {
        expect(error).toBeInstanceOf(SlidesSessionBusyError)
        expect((error as SlidesSessionBusyError).code).toBe('slides_session_busy')
      }
    }
    expectBusy(() => pushHistory(session))
    expectBusy(() => undoSession(session))
    expectBusy(() => redoSession(session))
    expectBusy(() => restoreAiSnapshot(session, snapshotId))
    expect(valueOf(session)).toBe('after')
    expect(session.undoStack).toEqual(undoBefore)
    expect(session.redoStack).toEqual(redoBefore)

    release()
    expect(undoSession(session)).toBe(true)
    expect(valueOf(session)).toBe('before')
  })

  it('does not enter a production mutation wrapper while a transaction is paused', async () => {
    const session = sessionWith('before')
    const release = acquirePresentationTransactionLease(session)!
    let entered = false

    await expect(
      withPresentationMutationLease(session, () => {
        entered = true
        setValue(session, 'forbidden')
      }),
    ).rejects.toMatchObject({ code: 'slides_session_busy' })
    expect(entered).toBe(false)
    expect(valueOf(session)).toBe('before')

    release()
    await withPresentationMutationLease(session, () => setValue(session, 'after'))
    expect(valueOf(session)).toBe('after')
  })

  it('holds the ordinary mutation gate across async work and releases it afterward', async () => {
    const session = sessionWith('before')
    let finish!: () => void
    const paused = new Promise<void>((resolve) => {
      finish = resolve
    })
    const mutation = withPresentationMutationLease(session, async () => {
      pushHistory(session)
      setValue(session, 'intermediate')
      await paused
      setValue(session, 'after')
    })

    await Promise.resolve()
    expect(valueOf(session)).toBe('intermediate')
    expect(acquirePresentationTransactionLease(session)).toBeNull()
    expect(acquirePresentationPersistenceLease(session)).toBeNull()
    expect(sessionBlocksPresentationClose(session)).toBe(true)

    finish()
    await mutation
    expect(valueOf(session)).toBe('after')
    const releaseTransaction = acquirePresentationTransactionLease(session)
    expect(releaseTransaction).not.toBeNull()
    releaseTransaction!()
  })

  it('collapses several edits into one pre-run snapshot', () => {
    const session = sessionWith('before')
    beginHistoryBatch(session)
    pushHistory(session)
    setValue(session, 'first')
    pushHistory(session)
    setValue(session, 'second')
    endHistoryBatch(session)

    expect(session.undoStack).toHaveLength(1)
    restoreSnapshot(session, session.undoStack[0]!)
    expect(valueOf(session)).toBe('before')
  })

  it('supports nested tool batching inside an AI-run batch', () => {
    const session = sessionWith('before')
    beginHistoryBatch(session)
    beginHistoryBatch(session)
    pushHistory(session)
    setValue(session, 'after')
    endHistoryBatch(session)
    expect(session.historyBatch?.depth).toBe(1)
    endHistoryBatch(session)
    expect(session.undoStack).toHaveLength(1)
  })

  it('does not create a history step when every edit is rolled back', () => {
    const session = sessionWith('before')
    beginHistoryBatch(session)
    pushHistory(session)
    session.undoStack.pop()
    endHistoryBatch(session)
    expect(session.undoStack).toHaveLength(0)
  })

  it('restores the deck size on undo', () => {
    const session = sessionWith('before')
    pushHistory(session)
    session.opened.deck.size = { cx: 2, cy: 3 }
    setValue(session, 'after')

    restoreSnapshot(session, session.undoStack.pop()!)
    expect(session.opened.deck.size).toEqual({ cx: 1, cy: 1 })
    expect(valueOf(session)).toBe('before')
  })

  it('returns the pre-run snapshot from the outermost batch end with edits', () => {
    const session = sessionWith('before')
    beginHistoryBatch(session)
    beginHistoryBatch(session)
    pushHistory(session)
    setValue(session, 'after')
    expect(endHistoryBatch(session)).toBeNull()
    const before = endHistoryBatch(session)
    expect(before).not.toBeNull()
    expect((before!.slides[0] as unknown as { value: string }).value).toBe('before')

    const emptyRun = sessionWith('untouched')
    beginHistoryBatch(emptyRun)
    expect(endHistoryBatch(emptyRun)).toBeNull()
  })

  it('rolls back to a registered AI snapshot and makes the rollback undoable', () => {
    const session = sessionWith('before')
    beginHistoryBatch(session)
    pushHistory(session)
    setValue(session, 'ai edited')
    const id = registerAiSnapshot(session, endHistoryBatch(session)!)

    setValue(session, 'user edited on top')
    expect(restoreAiSnapshot(session, id)).toBe(true)
    expect(valueOf(session)).toBe('before')
    expect(restoreAiSnapshot(session, id)).toBe(false) // consumed

    restoreSnapshot(session, session.undoStack.pop()!) // ⌘Z returns to the pre-rollback state
    expect(valueOf(session)).toBe('user edited on top')
  })

  it('keeps registered snapshots isolated from later in-place deck mutations', () => {
    const session = sessionWith('before')
    beginHistoryBatch(session)
    pushHistory(session)
    setValue(session, 'ai edited')
    const id = registerAiSnapshot(session, endHistoryBatch(session)!)

    restoreSnapshot(session, session.undoStack.pop()!) // undo hands the stack snapshot to the live deck
    setValue(session, 'mutated after undo')
    expect(restoreAiSnapshot(session, id)).toBe(true)
    expect(valueOf(session)).toBe('before')
  })

  it('undo → edit → redo replays the state that was undone, not a mutated copy', () => {
    const session = sessionWith('before')
    pushHistory(session)
    setValue(session, 'edited')

    // ⌘Z
    session.redoStack.push({
      slides: structuredClone(session.opened.deck.slides),
      entries: new Map(session.opened.archive.entries),
      size: { ...session.opened.deck.size },
    })
    restoreSnapshot(session, session.undoStack.pop()!)
    expect(valueOf(session)).toBe('before')

    // typing after the undo must not rewrite the redo snapshot in place
    setValue(session, 'typed after undo')
    restoreSnapshot(session, session.redoStack.pop()!)
    expect(valueOf(session)).toBe('edited')
  })

  it('a batch left open by a crashed tool path is collapsed so undo still works', () => {
    const session = sessionWith('before')
    // run begins a batch, a tool nests another, then the tool path dies without ending either
    beginHistoryBatch(session)
    pushHistory(session)
    setValue(session, 'ai edited')
    beginHistoryBatch(session)
    expect(session.historyBatch).toBeDefined()

    settleStaleHistoryBatch(session)
    expect(session.historyBatch).toBeUndefined()
    expect(session.undoStack.length).toBe(1)

    restoreSnapshot(session, session.undoStack.pop()!)
    expect(valueOf(session)).toBe('before')
  })
})
