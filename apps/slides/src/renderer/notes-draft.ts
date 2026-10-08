export interface NotesDraft {
  index: number
  text: string
  version: number
}

export interface NotesDraftRef {
  current: NotesDraft | null
}

export interface NotesFlushReceipt {
  flushedVersion: number | null
  persisted: boolean
  stillDirty: boolean
}

/**
 * Persist exactly the captured draft. A later draft is never cleared by this
 * completion, including when the host write itself succeeded.
 */
export async function flushNotesDraft(
  draft: NotesDraftRef,
  persist: (pending: NotesDraft) => Promise<boolean>,
  expected: NotesDraft | null = draft.current,
): Promise<NotesFlushReceipt> {
  if (draft.current !== expected) {
    return {
      flushedVersion: expected?.version ?? null,
      persisted: false,
      stillDirty: draft.current !== null,
    }
  }
  if (!expected) return { flushedVersion: null, persisted: true, stillDirty: false }
  const persisted = await persist(expected).catch(() => false)
  if (persisted && draft.current === expected) draft.current = null
  return {
    flushedVersion: expected.version,
    persisted,
    stillDirty: draft.current !== null,
  }
}

export interface SpeakerNotesDraftPreparation {
  ready: boolean
  expectedDraftVersion: number
}

/** Bind AI preparation to the draft and editor context visible at invocation start. */
export async function prepareSpeakerNotesDraft(options: {
  draft: NotesDraftRef
  currentVersion: () => number
  contextKey: () => unknown
  flush: (expected: NotesDraft | null) => Promise<NotesFlushReceipt>
}): Promise<SpeakerNotesDraftPreparation> {
  const expectedDraft = options.draft.current
  const expectedDraftVersion = options.currentVersion()
  const expectedContext = options.contextKey()
  const receipt = await options.flush(expectedDraft)
  const stable =
    receipt.persisted &&
    !receipt.stillDirty &&
    receipt.flushedVersion === (expectedDraft?.version ?? null) &&
    options.draft.current === null &&
    options.currentVersion() === expectedDraftVersion &&
    options.contextKey() === expectedContext
  return { ready: stable, expectedDraftVersion }
}
