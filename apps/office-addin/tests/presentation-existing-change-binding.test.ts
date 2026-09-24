import { expect, it, vi } from 'vitest'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import {
  validatePresentationExistingChange,
  type PresentationExistingChange,
} from '../src/skills/powerpoint/presentation-existing-change'
const historyKey = 'wiswork.presentation.change-history.v1'
const slotKey = 'wiswork.presentation.existing-change.v1'
async function fixture() {
  const values = new Map<string, string>()
  const save = vi.fn(async () => {})
  let location = 'deck'
  const settings = {
    get: (key: string) => values.get(key),
    set: (key: string, value: string) => {
      values.set(key, value)
    },
    save,
    location: () => location,
  }
  const create = () => createPresentationDocumentBinding(settings, () => 'doc')
  const binding = create()
  const record: PresentationExistingChange = {
    version: 1,
    changeId: 'one',
    documentId: await binding.documentId(),
    baselineId: 'baseline',
    baselineDigest: 'a'.repeat(64),
    scope: { slideIds: ['slide'], shapeIds: ['shape'] },
    hostSlideId: 'slide',
    shapeId: 'shape',
    shapeType: 'TextBox',
    kind: 'text',
    before: 'one',
    after: 'two',
    state: 'pending',
  }
  return {
    values,
    save,
    create,
    binding,
    record,
    move: () => {
      location = 'other'
    },
  }
}
it('validates exact bounded records and rejects forged scope and reviews on pending records', async () => {
  const { record } = await fixture()
  expect(validatePresentationExistingChange(record)).toBe(true)
  expect(
    validatePresentationExistingChange({ ...record, scope: { slideIds: ['elsewhere'] } }),
  ).toBe(false)
  expect(validatePresentationExistingChange({ ...record, projectId: 'fake' })).toBe(false)
})
it('persists exact-ID records across reopen and requires CAS on transitions', async () => {
  const f = await fixture()
  await f.binding.writeExistingChange(f.record, undefined)
  const applied = { ...f.record, state: 'applied' as const }
  await f.binding.writeExistingChange(applied, f.record)
  const second = { ...f.record, changeId: 'two' }
  await f.binding.writeExistingChange(second, undefined)
  const reopened = f.create()
  expect(reopened.readExistingChange('one')).toEqual(applied)
  expect(reopened.readExistingChange('missing')).toBeUndefined()
  expect(reopened.listChangeHistory().map((e) => e.id)).toEqual(['existing:one', 'existing:two'])
  await expect(reopened.writeExistingChange(applied, f.record)).rejects.toThrow('stale')
})
const review = {
  screenshotDigest: 'b'.repeat(64),
  capturedAt: '2026-09-24T00:00:00.000Z',
  reviewedAt: '2026-09-24T00:00:01.000Z',
  status: 'pass' as const,
  notes: 'checked',
}
it('allows terminal review updates, preserves immutable core and clears review on state advance', async () => {
  const f = await fixture()
  expect(validatePresentationExistingChange({ ...f.record, review })).toBe(false)
  await f.binding.writeExistingChange(f.record, undefined)
  const applied = { ...f.record, state: 'applied' as const }
  await f.binding.writeExistingChange(applied, f.record)
  const reviewed = { ...applied, review }
  await f.binding.writeExistingChange(reviewed, applied)
  const saves = f.save.mock.calls.length
  await f.binding.writeExistingChange(reviewed, reviewed)
  expect(f.save).toHaveBeenCalledTimes(saves)
  await expect(
    f.binding.writeExistingChange({ ...reviewed, before: 'forged' }, reviewed),
  ).rejects.toThrow('state_invalid')
  await expect(
    f.binding.writeExistingChange({ ...reviewed, state: 'undone' }, reviewed),
  ).rejects.toThrow('state_invalid')
  await expect(
    f.binding.writeExistingChange({ ...reviewed, state: 'undo_pending' }, reviewed),
  ).rejects.toThrow('state_invalid')
  const undo = { ...applied, state: 'undo_pending' as const }
  await f.binding.writeExistingChange(undo, reviewed)
  const undone = { ...undo, state: 'undone' as const }
  await f.binding.writeExistingChange(undone, undo)
  expect(f.create().readExistingChange('one')).toEqual(undone)
  expect(f.create().listChangeHistory()[0].sequence).toBe(1)
})
it('blocks new transactions across existing and generated sources while allowing pending completion', async () => {
  const f = await fixture()
  const generated = {
    version: 1 as const,
    changeId: 'generated',
    documentId: f.record.documentId,
    projectId: 'project',
    requestId: 'request',
    pageId: 'page',
    artifactDigest: 'a'.repeat(64),
    hostSlideId: 'slide',
    shapeId: 'shape',
    before: 'one',
    after: 'two',
    state: 'pending' as const,
  }
  await f.binding.writeExistingChange(f.record, undefined)
  await expect(f.binding.writeTextChange(generated, undefined)).rejects.toThrow('history_pending')
  await expect(
    f.binding.writeExistingChange({ ...f.record, changeId: 'two' }, undefined),
  ).rejects.toThrow('history_pending')
  await f.binding.writeExistingChange({ ...f.record, state: 'applied' }, f.record)
  await f.binding.writeTextChange(generated, undefined)
  await expect(
    f.binding.writeExistingChange({ ...f.record, changeId: 'two' }, undefined),
  ).rejects.toThrow('history_pending')
})
it('rolls back slot and history together on save failure and detects old-slot divergence', async () => {
  const f = await fixture()
  await f.binding.writeExistingChange(f.record, undefined)
  const prior = f.values.get(slotKey),
    history = f.values.get(historyKey)
  f.save.mockRejectedValueOnce(new Error('disk'))
  await expect(
    f.binding.writeExistingChange({ ...f.record, state: 'applied' }, f.record),
  ).rejects.toThrow('disk')
  expect(f.values.get(slotKey)).toBe(prior)
  expect(f.values.get(historyKey)).toBe(history)
  expect(f.create().readExistingChange('one')).toEqual(f.record)
  f.values.set(slotKey, JSON.stringify({ ...f.record, state: 'applied' }))
  expect(() => f.binding.listChangeHistory()).toThrow('history_state_invalid')
})
it('rejects changed document identity before saving and poisons reads after SaveAs during saving', async () => {
  const f = await fixture()
  await expect(
    f.binding.writeExistingChange({ ...f.record, documentId: 'other' }, undefined),
  ).rejects.toThrow('document_changed')
  f.save.mockImplementationOnce(async () => {
    f.move()
  })
  await expect(f.binding.writeExistingChange(f.record, undefined)).rejects.toThrow(
    'document_changed',
  )
  expect(() => f.binding.readExistingChange('one')).toThrow('state_invalid')
})
it('retains 64 entries and reserves full review and undo growth before writes', async () => {
  const f = await fixture()
  for (let index = 0; index < 64; index++) {
    const record = { ...f.record, changeId: `entry_${index}` }
    await f.binding.writeExistingChange(record, undefined)
    await f.binding.writeExistingChange({ ...record, state: 'applied' }, record)
  }
  await expect(
    f.binding.writeExistingChange({ ...f.record, changeId: 'extra' }, undefined),
  ).rejects.toThrow('history_full')
  const first = f.binding.readExistingChange('entry_0')!
  await f.binding.writeExistingChange(
    { ...first, review: { ...review, notes: 'x'.repeat(7900) } },
    first,
  )
  expect(f.create().listChangeHistory()).toHaveLength(64)
  const g = await fixture()
  let count = 0
  for (; count < 64; count++) {
    const record = {
      ...g.record,
      changeId: `large_${count}`,
      before: 'x'.repeat(12000),
      after: 'y'.repeat(12000),
    }
    try {
      await g.binding.writeExistingChange(record, undefined)
    } catch (e) {
      expect(String(e)).toContain('history_full')
      break
    }
    await g.binding.writeExistingChange({ ...record, state: 'applied' }, record)
  }
  expect(count).toBeGreaterThan(0)
  expect(count).toBeLessThan(64)
  for (const e of g.binding.listChangeHistory()) {
    if (e.kind !== 'existing') throw new Error('unexpected')
    const reviewed = { ...e.record, review: { ...review, notes: 'x'.repeat(7900) } }
    await g.binding.writeExistingChange(reviewed, e.record)
    const undo = { ...e.record, state: 'undo_pending' as const }
    await g.binding.writeExistingChange(undo, reviewed)
    await g.binding.writeExistingChange({ ...undo, state: 'undone' }, undo)
  }
  expect(g.create().listChangeHistory()).toHaveLength(count)
})
it('rejects corrupted records and invalid review timestamps and geometry', async () => {
  const f = await fixture()
  expect(
    validatePresentationExistingChange({
      ...f.record,
      kind: 'geometry',
      before: { left: 0, top: 0, width: 1, height: 1 },
      after: { left: 0, top: 0, width: -1, height: 1 },
    }),
  ).toBe(false)
  expect(
    validatePresentationExistingChange({
      ...f.record,
      state: 'applied',
      review: { ...review, notes: 'x'.repeat(8192) },
    }),
  ).toBe(false)
  expect(
    validatePresentationExistingChange({
      ...f.record,
      state: 'applied',
      review: { ...review, reviewedAt: 'bad' },
    }),
  ).toBe(false)
  f.values.set(slotKey, '{')
  expect(() => f.binding.readExistingChange('one')).toThrow('state_invalid')
})
