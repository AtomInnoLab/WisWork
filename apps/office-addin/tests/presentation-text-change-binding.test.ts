import { expect, it, vi } from 'vitest'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import {
  validatePresentationTextChange,
  type PresentationTextChange,
} from '../src/skills/powerpoint/presentation-text-change'
const key = 'wiswork.presentation.text-change.v1'
async function fixture() {
  const values = new Map<string, string>(),
    save = vi.fn(async () => {})
  let location = 'file://deck.pptx'
  const settings = {
    get: (key: string) => values.get(key),
    set: (key: string, value: string) => {
      values.set(key, value)
    },
    save,
    location: () => location,
  }
  const create = () => createPresentationDocumentBinding(settings, () => 'document')
  const binding = create()
  const record: PresentationTextChange = {
    version: 1,
    changeId: 'change-1',
    documentId: await binding.documentId(),
    projectId: 'project',
    requestId: 'request',
    artifactDigest: 'a'.repeat(64),
    pageId: 'page',
    hostSlideId: 'host',
    shapeId: 'shape',
    before: 'original',
    after: 'replacement',
    state: 'pending',
  }
  return {
    values,
    save,
    settings,
    create,
    binding,
    record,
    move: () => {
      location = 'file://other.pptx'
    },
  }
}
it('validates bounded strict identity, text, state and source', async () => {
  const { record } = await fixture()
  expect(validatePresentationTextChange(record)).toBe(true)
  expect(validatePresentationTextChange({ ...record, source: 'production' })).toBe(true)
  for (const patch of [
    { extra: true },
    { version: 2 },
    { source: 'legacy' },
    { changeId: '../bad' },
    { documentId: '' },
    { shapeId: 'bad\n' },
    { artifactDigest: 'x'.repeat(64) },
    { state: 'complete' },
    { before: 1 },
    { after: 'a'.repeat(12001) },
    { after: null },
    { before: {} },
    { documentId: 'a'.repeat(4097) },
  ])
    expect(validatePresentationTextChange({ ...record, ...patch })).toBe(false)
})
it('persists the monotonic lifecycle, idempotency and replaces only ended records', async () => {
  const f = await fixture()
  expect(f.binding.readTextChange()).toBeUndefined()
  let prior: PresentationTextChange | undefined
  for (const state of ['pending', 'applied', 'undo_pending', 'undone'] as const) {
    const next = { ...f.record, state }
    await f.binding.writeTextChange(next, prior)
    expect(f.create().readTextChange()).toEqual(next)
    const saves = f.save.mock.calls.length
    await f.binding.writeTextChange(next, next)
    expect(f.save.mock.calls.length).toBe(saves)
    prior = next
  }
  const next = { ...f.record, changeId: 'change-2' }
  await f.binding.writeTextChange(next, prior)
  expect(f.create().readTextChange()).toEqual(next)
})
it('rejects stale CAS, skipped transitions, content rewrites and pending replacement', async () => {
  const f = await fixture()
  await f.binding.writeTextChange(f.record, undefined)
  await expect(
    f.binding.writeTextChange({ ...f.record, state: 'applied' }, undefined),
  ).rejects.toThrow('presentation_text_change_stale')
  for (const patch of [
    { state: 'undone' as const },
    { changeId: 'other' },
    { source: 'production' as const },
    { after: 'changed' },
  ])
    await expect(f.binding.writeTextChange({ ...f.record, ...patch }, f.record)).rejects.toThrow(
      'presentation_text_change_state_invalid',
    )
  const applied = { ...f.record, state: 'applied' as const }
  await f.binding.writeTextChange(applied, f.record)
  const undo = { ...f.record, state: 'undo_pending' as const }
  await f.binding.writeTextChange(undo, applied)
  await expect(f.binding.writeTextChange({ ...f.record, changeId: 'other' }, undo)).rejects.toThrow(
    'presentation_text_change_state_invalid',
  )
})
it('copies record and expected before queueing and rejects a competing stale save', async () => {
  const f = await fixture()
  let release!: () => void
  f.save.mockImplementationOnce(
    () =>
      new Promise<void>((resolve) => {
        release = resolve
      }),
  )
  const first = f.binding.writeTextChange(f.record, undefined)
  await vi.waitFor(() => expect(release).toBeTypeOf('function'))
  const expected = structuredClone(f.record),
    next = { ...structuredClone(f.record), state: 'applied' as const }
  const second = f.binding.writeTextChange(next, expected)
  next.after = 'changed'
  expected.after = 'changed'
  release()
  await first
  await second
  expect(f.binding.readTextChange()?.after).toBe('replacement')
  const results = await Promise.allSettled([
    f.binding.writeTextChange(
      { ...f.record, state: 'undo_pending' },
      { ...f.record, state: 'applied' },
    ),
    f.binding.writeTextChange(
      { ...f.record, changeId: 'other' },
      { ...f.record, state: 'applied' },
    ),
  ])
  expect(results.map((result) => result.status)).toEqual(['fulfilled', 'rejected'])
})
it('restores previous settings on failure and locks the binding on document switch', async () => {
  const f = await fixture()
  f.save.mockRejectedValueOnce(new Error('save_failed'))
  await expect(f.binding.writeTextChange(f.record, undefined)).rejects.toThrow('save_failed')
  expect(f.create().readTextChange()).toBeUndefined()
  await f.binding.writeTextChange(f.record, undefined)
  f.save.mockImplementationOnce(async () => {
    f.move()
  })
  await expect(
    f.binding.writeTextChange({ ...f.record, state: 'applied' }, f.record),
  ).rejects.toThrow('presentation_document_changed')
  expect(() => f.binding.readTextChange()).toThrow('presentation_text_change_state_invalid')
})
it('rejects corrupted or oversized settings and wrong document identity', async () => {
  const f = await fixture()
  await expect(
    f.binding.writeTextChange({ ...f.record, documentId: 'other' }, undefined),
  ).rejects.toThrow('presentation_document_changed')
  for (const raw of [
    '{',
    '{}',
    'null',
    ' '.repeat(192 * 1024 + 1),
    JSON.stringify({ ...f.record, extra: true }),
  ]) {
    f.values.set(key, raw)
    expect(() => f.create().readTextChange()).toThrow('presentation_text_change_state_invalid')
  }
})
it('retains a pending marker when completion save fails and locks after a changed document identity', async () => {
  const f = await fixture()
  await expect(
    f.binding.writeTextChange({ ...f.record, state: 'applied' }, undefined),
  ).rejects.toThrow('presentation_text_change_state_invalid')
  await f.binding.writeTextChange(f.record, undefined)
  f.save.mockRejectedValueOnce(new Error('save_failed'))
  await expect(
    f.binding.writeTextChange({ ...f.record, state: 'applied' }, f.record),
  ).rejects.toThrow('save_failed')
  expect(f.create().readTextChange()).toEqual(f.record)
  f.save.mockImplementationOnce(async () => {
    f.values.set('wiswork.presentation.document.v1', 'different')
  })
  await expect(
    f.binding.writeTextChange({ ...f.record, state: 'applied' }, f.record),
  ).rejects.toThrow('presentation_document_changed')
  expect(() => f.binding.readTextChange()).toThrow('presentation_text_change_state_invalid')
})
it('reserves the full JSON byte limit and locks if failed save rollback cannot restore settings', async () => {
  const f = await fixture()
  expect(validatePresentationTextChange({ ...f.record, before: 'a'.repeat(12001) })).toBe(false)
  await f.binding.writeTextChange(f.record, undefined)
  const original = f.settings.set
  f.save.mockImplementationOnce(async () => {
    f.settings.set = () => {
      throw new Error('set_failed')
    }
    throw new Error('save_failed')
  })
  await expect(
    f.binding.writeTextChange({ ...f.record, state: 'applied' }, f.record),
  ).rejects.toThrow('save_failed')
  f.settings.set = original
  expect(() => f.binding.readTextChange()).toThrow('presentation_text_change_state_invalid')
})

it('accepts full bounded text and enumerates validated image copies without saves', async () => {
  const f = await fixture()
  const maximal = {
    ...f.record,
    before: '\u0000'.repeat(12000),
    after: '\u0000'.repeat(12000),
    documentId: '\u0000'.repeat(4096),
  }
  expect(validatePresentationTextChange(maximal)).toBe(true)
  const image = {
    version: 1,
    documentId: f.record.documentId,
    projectId: 'project',
    requestId: 'request',
    pageId: 'page',
    hostSlideId: 'host',
    oldShapeId: 'old',
    assetDigest: 'b'.repeat(64),
    state: 'pending',
  }
  const imageKey = 'wiswork.presentation.image-replacements.v1'
  f.values.set(imageKey, JSON.stringify({ ['a'.repeat(64)]: image }))
  const saves = f.save.mock.calls.length
  const records = f.binding.listImageReplacements()
  expect(records).toEqual([image])
  records[0].oldShapeId = 'changed'
  expect(f.binding.listImageReplacements()).toEqual([image])
  expect(f.save.mock.calls.length).toBe(saves)
  f.values.set(imageKey, '{')
  expect(() => f.binding.listImageReplacements()).toThrow(
    'presentation_image_replacement_state_invalid',
  )
})
