import { expect, it, vi } from 'vitest'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import {
  validatePresentationExistingImageChange,
  type PresentationExistingImageChange,
} from '../src/skills/powerpoint/presentation-existing-image'

async function fixture() {
  const values = new Map<string, string>()
  const save = vi.fn(async () => {})
  const settings = {
    get: (key: string) => values.get(key),
    set: (key: string, value: string) => {
      values.set(key, value)
    },
    save,
    location: () => 'deck',
  }
  const binding = createPresentationDocumentBinding(settings, () => 'doc')
  const digest = (c: string) => c.repeat(64)
  const original = {
    slideId: 's1',
    shapeId: 'p1',
    geometry: { left: 0, top: 0, width: 10, height: 10 },
    rotation: 0,
    name: 'Picture',
    altTextTitle: '',
    altTextDescription: '',
    zOrderPosition: 0,
    shapeIds: ['p1'],
    pictureFingerprint: digest('a'),
    mediaDigest: digest('b'),
  }
  const record: PresentationExistingImageChange = {
    version: 1,
    changeId: 'img1',
    documentId: await binding.documentId(),
    baselineId: 'baseline',
    baselineDigest: digest('c'),
    scope: { slideIds: ['s1'], shapeIds: ['p1'] },
    hostSlideId: 's1',
    oldShapeId: 'p1',
    assetDigest: digest('d'),
    original,
    backup: { attachmentId: digest('b'), sizeBytes: 123, mime: 'image/png' },
    state: 'pending',
  }
  return {
    binding,
    record,
    save,
    values,
    reopen: () => createPresentationDocumentBinding(settings, () => 'doc'),
  }
}

it('persists native image recovery evidence and ordered transitions across reopen', async () => {
  const f = await fixture()
  await f.binding.writeExistingImageChange(f.record, undefined)
  const inserted = { ...f.record, insertedShapeId: 'p2' }
  await f.reopen().writeExistingImageChange(inserted, f.record)
  const after = {
    ...f.record.original,
    shapeId: 'p2',
    shapeIds: ['p2'],
    mediaDigest: f.record.assetDigest,
  }
  const complete = { ...inserted, state: 'complete' as const, after }
  await f.binding.writeExistingImageChange(complete, inserted)
  const undo = { ...complete, state: 'undo_pending' as const, undoBaseline: after }
  await f.binding.writeExistingImageChange(undo, complete)
  const restored = { ...undo, restoredShapeId: 'p3' }
  await f.binding.writeExistingImageChange(restored, undo)
  const undone = { ...restored, state: 'undone' as const }
  await f.reopen().writeExistingImageChange(undone, restored)
  expect(f.reopen().readExistingImageChange('img1')).toEqual(undone)
  expect(f.reopen().listChangeHistory()).toMatchObject([{ kind: 'existing_image', record: undone }])
})

it('rejects malformed scope, backup, snapshots and impossible transitions', async () => {
  const f = await fixture()
  expect(
    validatePresentationExistingImageChange({ ...f.record, scope: { slideIds: ['other'] } }),
  ).toBe(false)
  expect(
    validatePresentationExistingImageChange({
      ...f.record,
      backup: { ...f.record.backup, attachmentId: 'e'.repeat(64) },
    }),
  ).toBe(false)
  expect(
    validatePresentationExistingImageChange({
      ...f.record,
      original: { ...f.record.original, shapeIds: ['other'] },
    }),
  ).toBe(false)
  await expect(
    f.binding.writeExistingImageChange({ ...f.record, state: 'complete' }, undefined),
  ).rejects.toThrow('state_invalid')
  await f.binding.writeExistingImageChange(f.record, undefined)
  await expect(
    f.binding.writeExistingImageChange({ ...f.record, insertedShapeId: 'p2' }, undefined),
  ).rejects.toThrow('stale')
  await expect(
    f.binding.writeExistingImageChange({ ...f.record, oldShapeId: 'other' }, f.record),
  ).rejects.toThrow('state_invalid')
  await expect(
    f.binding.writeExistingImageChange({ ...f.record, state: 'undo_pending' }, f.record),
  ).rejects.toThrow('state_invalid')
})

it('blocks other changes while image write is unresolved and restores both settings on failed save', async () => {
  const f = await fixture()
  await f.binding.writeExistingImageChange(f.record, undefined)
  await expect(
    f.binding.writeExistingImageChange({ ...f.record, changeId: 'img2' }, undefined),
  ).rejects.toThrow('pending')
  const before = new Map(f.values)
  f.save.mockRejectedValueOnce(new Error('save failed'))
  await expect(
    f.binding.writeExistingImageChange({ ...f.record, insertedShapeId: 'p2' }, f.record),
  ).rejects.toThrow('save failed')
  for (const [key, value] of before) expect(f.values.get(key)).toBe(value)
  expect(f.reopen().readExistingImageChange('img1')).toEqual(f.record)
})

it('stores an image review only at a terminal savepoint and clears it before undo', async () => {
  const f = await fixture()
  const inserted = { ...f.record, insertedShapeId: 'p2' }
  const after = { ...f.record.original, shapeId: 'p2', shapeIds: ['p2'], mediaDigest: f.record.assetDigest }
  const complete = { ...inserted, state: 'complete' as const, after }
  const review = { hostSlideId: 's1', screenshotDigest: 'e'.repeat(64), capturedAt: '2026-09-24T00:00:00.000Z', reviewedAt: '2026-09-24T00:01:00.000Z', status: 'pass' as const, notes: 'visual check' }
  await f.binding.writeExistingImageChange(f.record, undefined)
  await f.binding.writeExistingImageChange(inserted, f.record)
  await f.binding.writeExistingImageChange(complete, inserted)
  const reviewed = { ...complete, review }
  await f.reopen().writeExistingImageChange(reviewed, complete)
  expect(f.reopen().readExistingImageChange('img1')?.review).toEqual(review)
  await expect(f.binding.writeExistingImageChange({ ...reviewed, state: 'undo_pending', undoBaseline: after }, reviewed)).rejects.toThrow('state_invalid')
  const undo = { ...complete, state: 'undo_pending' as const, undoBaseline: after }
  await f.binding.writeExistingImageChange(undo, reviewed)
  expect(f.reopen().readExistingImageChange('img1')?.review).toBeUndefined()
})
