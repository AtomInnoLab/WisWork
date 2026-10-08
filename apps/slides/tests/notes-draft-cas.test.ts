import { describe, expect, it } from 'vitest'
import {
  flushNotesDraft,
  prepareSpeakerNotesDraft,
  type NotesDraft,
} from '../src/renderer/notes-draft'

function deferred<T>() {
  let resolve!: (value: T) => void
  let reject!: (reason?: unknown) => void
  const promise = new Promise<T>((res, rej) => {
    resolve = res
    reject = rej
  })
  return { promise, resolve, reject }
}

describe('speaker notes draft CAS', () => {
  it('fails closed and preserves input typed while the captured draft is flushing', async () => {
    const draft = { current: { index: 0, text: 'P', version: 1 } as NotesDraft | null }
    const write = deferred<boolean>()
    const preparation = prepareSpeakerNotesDraft({
      draft,
      currentVersion: () => 1,
      contextKey: () => 'session-a:slide-0',
      flush: (expected) => flushNotesDraft(draft, () => write.promise, expected),
    })
    draft.current = { index: 0, text: 'Q', version: 2 }
    write.resolve(true)
    await expect(preparation).resolves.toEqual({ ready: false, expectedDraftVersion: 1 })
    expect(draft.current).toEqual({ index: 0, text: 'Q', version: 2 })
  })

  it.each([
    ['false', () => Promise.resolve(false)],
    ['reject', () => Promise.reject(new Error('setNotes failed'))],
  ])('fails closed and retains the draft when setNotes returns %s', async (_label, persist) => {
    const pending: NotesDraft = { index: 0, text: 'P', version: 4 }
    const draft = { current: pending as NotesDraft | null }
    const result = await prepareSpeakerNotesDraft({
      draft,
      currentVersion: () => 4,
      contextKey: () => 'session-a:slide-0',
      flush: (expected) => flushNotesDraft(draft, persist, expected),
    })
    expect(result).toEqual({ ready: false, expectedDraftVersion: 4 })
    expect(draft.current).toBe(pending)
  })

  it('fails closed when the slide or session changes during flush', async () => {
    const draft = { current: { index: 0, text: 'P', version: 7 } as NotesDraft | null }
    const write = deferred<boolean>()
    let context = 'session-a:slide-0'
    const preparation = prepareSpeakerNotesDraft({
      draft,
      currentVersion: () => 7,
      contextKey: () => context,
      flush: (expected) => flushNotesDraft(draft, () => write.promise, expected),
    })
    context = 'session-b:slide-1'
    write.resolve(true)
    await expect(preparation).resolves.toEqual({ ready: false, expectedDraftVersion: 7 })
  })
})
