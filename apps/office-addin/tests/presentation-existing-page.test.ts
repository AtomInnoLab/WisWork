import { expect, it, vi } from 'vitest'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import {
  validatePresentationExistingPageChange,
  type PresentationExistingPageChange,
} from '../src/skills/powerpoint/presentation-existing-page'

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
  const d = (c: string) => c.repeat(64)
  const record: PresentationExistingPageChange = {
    version: 1,
    changeId: 'page1',
    documentId: await binding.documentId(),
    baselineId: 'baseline',
    baselineDigest: d('a'),
    scope: { slideIds: ['s1'] },
    oldSlideId: 's1',
    beforeSlideIds: ['s1', 's2'],
    originalPackageDigest: d('b'),
    replacementPackageDigest: d('c'),
    sourceSlideId: '256#',
    backup: { backupId: 'backup1', sha256: d('d'), sizeBytes: 120 },
    state: 'pending',
  }
  return {
    binding,
    record,
    values,
    save,
    reopen: () => createPresentationDocumentBinding(settings, () => 'doc'),
  }
}

it('persists native page replacement phases and undo across reopen', async () => {
  const f = await fixture()
  const phases: PresentationExistingPageChange[] = [
    f.record,
    { ...f.record, state: 'inserted', newSlideId: 's3' },
    { ...f.record, state: 'staged', newSlideId: 's3' },
    { ...f.record, state: 'commit_pending', newSlideId: 's3' },
    { ...f.record, state: 'applied', newSlideId: 's3' },
    { ...f.record, state: 'undo_pending', newSlideId: 's3' },
    { ...f.record, state: 'restore_inserted', newSlideId: 's3', restoredSlideId: 's4' },
    { ...f.record, state: 'undone', newSlideId: 's3', restoredSlideId: 's4' },
  ]
  for (let i = 0; i < phases.length; i++)
    await f.reopen().writeExistingPageChange(phases[i], phases[i - 1])
  expect(f.reopen().readExistingPageChange('page1')).toEqual(phases.at(-1))
  expect(f.reopen().listChangeHistory()).toMatchObject([
    { kind: 'existing_page', record: phases.at(-1) },
  ])
})

it('rejects bad backup, IDs, and skipped phases', async () => {
  const f = await fixture()
  expect(validatePresentationExistingPageChange({ ...f.record, scope: { slideIds: ['s9'] } })).toBe(
    false,
  )
  expect(
    validatePresentationExistingPageChange({ ...f.record, scope: { slideIds: ['s1', 's9'] } }),
  ).toBe(false)
  expect(
    validatePresentationExistingPageChange({
      ...f.record,
      backup: { ...f.record.backup, sha256: 'bad' },
    }),
  ).toBe(false)
  await expect(
    f.binding.writeExistingPageChange({ ...f.record, state: 'applied' }, undefined),
  ).rejects.toThrow('state_invalid')
  await f.binding.writeExistingPageChange(f.record, undefined)
  await expect(
    f.binding.writeExistingPageChange({ ...f.record, state: 'staged', newSlideId: 's3' }, f.record),
  ).rejects.toThrow('state_invalid')
  await expect(
    f.binding.writeExistingPageChange({ ...f.record, oldSlideId: 's2' }, f.record),
  ).rejects.toThrow('state_invalid')
  await expect(
    f.binding.writeExistingPageChange(
      { ...f.record, state: 'inserted', newSlideId: 's3' },
      undefined,
    ),
  ).rejects.toThrow('stale')
})

it('persists picture readback identity and forbids changing it after staging', async () => {
  const f = await fixture()
  const target = {
    shapeId: '7',
    name: 'Picture 1',
    beforeDigest: 'e'.repeat(64),
    afterDigest: 'f'.repeat(64),
  }
  const pending = { ...f.record, pictureTarget: target }
  expect(validatePresentationExistingPageChange(pending)).toBe(true)
  expect(
    validatePresentationExistingPageChange({
      ...pending,
      pictureTarget: { ...target, afterDigest: target.beforeDigest },
    }),
  ).toBe(false)
  await f.binding.writeExistingPageChange(pending, undefined)
  const inserted = { ...pending, state: 'inserted' as const, newSlideId: 's3' }
  await f.reopen().writeExistingPageChange(inserted, pending)
  expect(f.reopen().readExistingPageChange('page1')?.pictureTarget).toEqual(target)
  await expect(
    f.binding.writeExistingPageChange(
      { ...inserted, state: 'staged', pictureTarget: { ...target, afterDigest: 'a'.repeat(64) } },
      inserted,
    ),
  ).rejects.toThrow('state_invalid')
})

it('persists an exact original-edit restore source and forbids changing its provenance', async () => {
  const f = await fixture()
  const restores = {
    sourceKind: 'single' as const,
    sourceChangeId: 'single',
    sourceHostSlideId: 's1',
    originalBackupId: 'source-backup',
    originalPackageDigest: f.record.replacementPackageDigest,
  }
  const pending = { ...f.record, restores }
  expect(validatePresentationExistingPageChange(pending)).toBe(true)
  expect(
    validatePresentationExistingPageChange({
      ...pending,
      restores: { ...restores, sourceHostSlideId: 's2' },
    }),
  ).toBe(false)
  expect(
    validatePresentationExistingPageChange({
      ...pending,
      restores: { ...restores, originalPackageDigest: f.record.originalPackageDigest },
    }),
  ).toBe(false)
  await f.binding.writeExistingPageChange(pending, undefined)
  const inserted = { ...pending, state: 'inserted' as const, newSlideId: 's3' }
  await f.binding.writeExistingPageChange(inserted, pending)
  await expect(
    f.binding.writeExistingPageChange(
      { ...inserted, state: 'staged', restores: { ...restores, sourceChangeId: 'other' } },
      inserted,
    ),
  ).rejects.toThrow('state_invalid')
})

it('allows staged replacement to be discarded and releases the global pending guard', async () => {
  const f = await fixture()
  const inserted = { ...f.record, state: 'inserted' as const, newSlideId: 's3' }
  const staged = { ...inserted, state: 'staged' as const }
  const discarding = { ...inserted, state: 'discard_pending' as const }
  const discarded = { ...inserted, state: 'discarded' as const }
  await f.binding.writeExistingPageChange(f.record, undefined)
  await f.binding.writeExistingPageChange(inserted, f.record)
  await f.binding.writeExistingPageChange(staged, inserted)
  await f.binding.writeExistingPageChange(discarding, staged)
  await f.binding.writeExistingPageChange(discarded, discarding)
  const released = { ...discarded, backupReleasedAt: '2026-09-24T00:00:00.000Z' }
  await f.binding.writeExistingPageChange(released, discarded)
  await expect(
    f.binding.writeExistingPageChange(
      { ...released, backupReleasedAt: '2026-09-24T00:00:01.000Z' },
      released,
    ),
  ).rejects.toThrow('state_invalid')
  const reviewed = {
    ...released,
    captures: [
      {
        hostSlideId: 's1',
        screenshotDigest: 'e'.repeat(64),
        capturedAt: '2026-09-24T00:01:00.000Z',
      },
    ],
  }
  await f.binding.writeExistingPageChange(reviewed, released)
  await f.binding.writeExistingPageChange({ ...f.record, changeId: 'page2' }, undefined)
  expect(f.reopen().listChangeHistory()).toMatchObject([
    { kind: 'existing_page', record: reviewed },
    { kind: 'existing_page', record: { changeId: 'page2', state: 'pending' } },
  ])
})

it('blocks concurrent changes and rolls back a failed save', async () => {
  const f = await fixture()
  await f.binding.writeExistingPageChange(f.record, undefined)
  await expect(
    f.binding.writeExistingPageChange({ ...f.record, changeId: 'page2' }, undefined),
  ).rejects.toThrow('pending')
  const before = new Map(f.values)
  f.save.mockRejectedValueOnce(new Error('save failed'))
  await expect(
    f.binding.writeExistingPageChange(
      { ...f.record, state: 'inserted', newSlideId: 's3' },
      f.record,
    ),
  ).rejects.toThrow('save failed')
  expect([...f.values]).toEqual([...before])
  expect(f.reopen().readExistingPageChange('page1')).toEqual(f.record)
})

it('stores each staged page review and invalidates both at commit', async () => {
  const f = await fixture()
  const inserted = { ...f.record, state: 'inserted' as const, newSlideId: 's3' }
  const staged = { ...inserted, state: 'staged' as const }
  const base = {
    screenshotDigest: 'e'.repeat(64),
    capturedAt: '2026-09-24T00:00:00.000Z',
    reviewedAt: '2026-09-24T00:01:00.000Z',
    status: 'pass' as const,
    notes: 'checked',
  }
  await f.binding.writeExistingPageChange(f.record, undefined)
  await f.binding.writeExistingPageChange(inserted, f.record)
  await f.binding.writeExistingPageChange(staged, inserted)
  const first = { ...staged, reviews: [{ ...base, hostSlideId: 's1' }] }
  const both = { ...staged, reviews: [...first.reviews, { ...base, hostSlideId: 's3' }] }
  await f.binding.writeExistingPageChange(first, staged)
  await f.reopen().writeExistingPageChange(both, first)
  await expect(
    f.binding.writeExistingPageChange({ ...both, state: 'commit_pending' }, both),
  ).rejects.toThrow('state_invalid')
  const pending = { ...staged, state: 'commit_pending' as const }
  await f.binding.writeExistingPageChange(pending, both)
  expect(f.reopen().readExistingPageChange('page1')?.reviews).toBeUndefined()
})

it('bounds durable source metadata and protects both backup identities and provenance after reopen', async () => {
  const f = await fixture()
  const sourceBackup = { backupId: 'source1', sha256: 'e'.repeat(64), sizeBytes: 123 }
  const record = { ...f.record, sourceBackup, reapplies: 'previous' }
  expect(validatePresentationExistingPageChange(record)).toBe(true)
  for (const invalid of [
    { ...record, sourceBackup: { ...sourceBackup, extra: true } },
    { ...record, sourceBackup: { ...sourceBackup, sizeBytes: 104857601 } },
    { ...record, sourceBackup: { ...sourceBackup, sha256: 'bad' } },
    { ...record, sourceBackup: { ...sourceBackup, backupId: record.backup.backupId } },
    { ...record, sourceBackup: undefined },
    { ...record, reapplies: record.changeId },
    { ...record, reapplies: 'a'.repeat(129) },
  ])
    expect(validatePresentationExistingPageChange(invalid)).toBe(false)
  await f.binding.writeExistingPageChange(record, undefined)
  expect(f.reopen().readExistingPageChange(record.changeId)).toEqual(record)
  for (const changed of [
    { ...record, sourceBackup: { ...sourceBackup, sizeBytes: 124 } },
    { ...record, sourceBackup: { ...sourceBackup, backupId: 'another' } },
    { ...record, sourceSlideId: '257#' },
    { ...record, reapplies: 'another' },
    { ...record, reapplies: undefined },
  ])
    await expect(
      f
        .reopen()
        .writeExistingPageChange({ ...changed, state: 'inserted', newSlideId: 's3' }, record),
    ).rejects.toThrow('state_invalid')
})

it('permits historical restore source host identity only on a source-backed reapply record', async () => {
  const f = await fixture()
  const restores = {
    sourceKind: 'single' as const,
    sourceChangeId: 'original',
    sourceHostSlideId: 'historical',
    originalBackupId: 'historical-backup',
    originalPackageDigest: f.record.replacementPackageDigest,
  }
  expect(validatePresentationExistingPageChange({ ...f.record, restores })).toBe(false)
  const record = {
    ...f.record,
    restores,
    reapplies: 'previous',
    sourceBackup: { backupId: 'source1', sha256: 'e'.repeat(64), sizeBytes: 123 },
  }
  expect(validatePresentationExistingPageChange(record)).toBe(true)
  await f.binding.writeExistingPageChange(record, undefined)
  expect(f.reopen().readExistingPageChange(record.changeId)?.restores).toEqual(restores)
  await expect(
    f.binding.writeExistingPageChange(
      {
        ...record,
        restores: { ...restores, sourceHostSlideId: f.record.oldSlideId },
        state: 'inserted',
        newSlideId: 's3',
      },
      record,
    ),
  ).rejects.toThrow('state_invalid')
})
