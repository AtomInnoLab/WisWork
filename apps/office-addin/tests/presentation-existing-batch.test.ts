import { expect, it, vi } from 'vitest'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import {
  validatePresentationExistingBatch,
  type PresentationExistingBatch,
} from '../src/skills/powerpoint/presentation-existing-batch'

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
  const batch: PresentationExistingBatch = {
    version: 1,
    changeId: 'batch',
    documentId: await binding.documentId(),
    baselineId: 'baseline',
    baselineDigest: 'a'.repeat(64),
    scope: { slideIds: ['s1', 's2'], shapeIds: ['a', 'b'] },
    intent: 'Update title and position',
    preserved: ['Other objects'],
    validation: ['Readback', 'Screenshots'],
    risk: 'medium',
    operations: [
      {
        kind: 'text',
        hostSlideId: 's1',
        shapeId: 'a',
        shapeType: 'text',
        before: 'old',
        after: 'new',
      },
      {
        kind: 'geometry',
        hostSlideId: 's2',
        shapeId: 'b',
        shapeType: 'text',
        before: { left: 0, top: 0, width: 10, height: 10 },
        after: { left: 1, top: 0, width: 10, height: 10 },
      },
    ],
    state: 'applying',
    cursor: 0,
  }
  return {
    values,
    save,
    binding,
    batch,
    reopen: () => createPresentationDocumentBinding(settings, () => 'doc'),
  }
}

it('journals ordered forward and reverse progress across reopen', async () => {
  const f = await fixture()
  await f.binding.writeExistingBatch(f.batch, undefined)
  const one = { ...f.batch, cursor: 1 }
  await f.binding.writeExistingBatch(one, f.batch)
  expect(f.reopen().readExistingBatch('batch')).toEqual(one)
  const applied = { ...one, cursor: 2, state: 'applied' as const }
  await f.reopen().writeExistingBatch(applied, one)
  const reviewed = {
    ...applied,
    reviews: [
      {
        hostSlideId: 's1',
        screenshotDigest: 'b'.repeat(64),
        capturedAt: '2026-09-24T00:00:00.000Z',
        reviewedAt: '2026-09-24T00:01:00.000Z',
        status: 'pass' as const,
        notes: 'Looks correct',
      },
    ],
  }
  await f.binding.writeExistingBatch(reviewed, applied)
  expect(f.reopen().readExistingBatch('batch')?.reviews).toEqual(reviewed.reviews)
  await expect(f.binding.writeExistingBatch(applied, reviewed)).rejects.toThrow('state_invalid')
  const undoing = { ...applied, state: 'undoing' as const }
  await f.binding.writeExistingBatch(undoing, reviewed)
  const reverse = { ...undoing, cursor: 1 }
  await f.binding.writeExistingBatch(reverse, undoing)
  const undone = { ...reverse, cursor: 0, state: 'undone' as const }
  await f.binding.writeExistingBatch(undone, reverse)
  expect(f.reopen().listChangeHistory()).toMatchObject([{ kind: 'existing_batch', record: undone }])
})

it('accepts only bounded terminal reviews on affected pages and clears them before undo', async () => {
  const f = await fixture()
  const review = {
    hostSlideId: 's1',
    screenshotDigest: 'b'.repeat(64),
    capturedAt: '2026-09-24T00:00:00.000Z',
    reviewedAt: '2026-09-24T00:01:00.000Z',
    status: 'pass',
    notes: '',
  }
  expect(validatePresentationExistingBatch({ ...f.batch, reviews: [review] })).toBe(false)
  const applied = { ...f.batch, state: 'applied', cursor: 2 }
  expect(
    validatePresentationExistingBatch({
      ...applied,
      reviews: [{ ...review, hostSlideId: 'other' }],
    }),
  ).toBe(false)
  expect(validatePresentationExistingBatch({ ...applied, reviews: [review, review] })).toBe(false)
  expect(
    validatePresentationExistingBatch({
      ...applied,
      reviews: [{ ...review, notes: 'x'.repeat(8192) }],
    }),
  ).toBe(false)
  expect(validatePresentationExistingBatch({ ...applied, reviews: [review] })).toBe(true)
})

it('keeps large records from the previous batch format readable without retroactive review reservation', async () => {
  const f = await fixture()
  const legacy: PresentationExistingBatch = {
    ...f.batch,
    scope: { slideIds: ['s1'] },
    operations: Array.from({ length: 8 }, (_, index) => ({
      kind: 'text' as const,
      hostSlideId: 's1',
      shapeId: `shape${index}`,
      shapeType: 'TextBox',
      before: 'a'.repeat(11800),
      after: 'b'.repeat(11800),
    })),
  }
  expect(validatePresentationExistingBatch(legacy)).toBe(true)
  expect(validatePresentationExistingBatch({ ...legacy, reviewCapacity: true })).toBe(false)
  await f.binding.writeExistingBatch(legacy, undefined)
  expect(f.reopen().readExistingBatch('batch')).toEqual(legacy)
})

it('rejects invalid scope, repeated target operations, stale CAS, skipped progress and parallel changes', async () => {
  const f = await fixture()
  expect(validatePresentationExistingBatch({ ...f.batch, scope: { slideIds: ['s1'] } })).toBe(false)
  expect(
    validatePresentationExistingBatch({
      ...f.batch,
      operations: [f.batch.operations[0], f.batch.operations[0]],
    }),
  ).toBe(false)
  await f.binding.writeExistingBatch(f.batch, undefined)
  await expect(f.binding.writeExistingBatch({ ...f.batch, cursor: 2 }, f.batch)).rejects.toThrow(
    'state_invalid',
  )
  await expect(f.binding.writeExistingBatch({ ...f.batch, cursor: 1 }, undefined)).rejects.toThrow(
    'stale',
  )
  await expect(
    f.binding.writeExistingBatch({ ...f.batch, changeId: 'other' }, undefined),
  ).rejects.toThrow('history_pending')
})

it('restores both batch slot and history if settings save fails', async () => {
  const f = await fixture()
  await f.binding.writeExistingBatch(f.batch, undefined)
  const history = f.values.get('wiswork.presentation.change-history.v1')
  const slot = f.values.get('wiswork.presentation.existing-batch.v1')
  f.save.mockRejectedValueOnce(new Error('failed'))
  await expect(f.binding.writeExistingBatch({ ...f.batch, cursor: 1 }, f.batch)).rejects.toThrow(
    'failed',
  )
  expect(f.values.get('wiswork.presentation.change-history.v1')).toBe(history)
  expect(f.values.get('wiswork.presentation.existing-batch.v1')).toBe(slot)
  expect(f.reopen().readExistingBatch('batch')).toEqual(f.batch)
})

it('does not leave a forged historical review when its settings save fails', async () => {
  const f = await fixture()
  const pending = { ...f.batch, reviewCapacity: true as const }
  await f.binding.writeExistingBatch(pending, undefined)
  const one = { ...pending, cursor: 1 }
  await f.binding.writeExistingBatch(one, pending)
  const applied = { ...one, state: 'applied' as const, cursor: 2 }
  await f.binding.writeExistingBatch(applied, one)
  const priorHistory = f.values.get('wiswork.presentation.change-history.v1')
  const priorSlot = f.values.get('wiswork.presentation.existing-batch.v1')
  const reviewed = {
    ...applied,
    reviews: [
      {
        hostSlideId: 's1',
        screenshotDigest: 'b'.repeat(64),
        capturedAt: '2026-09-24T00:00:00.000Z',
        reviewedAt: '2026-09-24T00:01:00.000Z',
        status: 'pass' as const,
        notes: 'Checked',
      },
    ],
  }
  f.save.mockRejectedValueOnce(new Error('failed'))
  await expect(f.binding.writeExistingBatch(reviewed, applied)).rejects.toThrow('failed')
  expect(f.values.get('wiswork.presentation.change-history.v1')).toBe(priorHistory)
  expect(f.values.get('wiswork.presentation.existing-batch.v1')).toBe(priorSlot)
  expect(f.reopen().readExistingBatch('batch')?.reviews).toBeUndefined()
})
