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
    set: (key: string, value: string) => { values.set(key, value) },
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
  return { binding, record, values, save, reopen: () => createPresentationDocumentBinding(settings, () => 'doc') }
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
  expect(f.reopen().listChangeHistory()).toMatchObject([{ kind: 'existing_page', record: phases.at(-1) }])
})

it('rejects bad backup, IDs, and skipped phases', async () => {
  const f = await fixture()
  expect(validatePresentationExistingPageChange({ ...f.record, scope: { slideIds: ['s9'] } })).toBe(false)
  expect(validatePresentationExistingPageChange({ ...f.record, scope: { slideIds: ['s1', 's9'] } })).toBe(false)
  expect(validatePresentationExistingPageChange({ ...f.record, backup: { ...f.record.backup, sha256: 'bad' } })).toBe(false)
  await expect(f.binding.writeExistingPageChange({ ...f.record, state: 'applied' }, undefined)).rejects.toThrow('state_invalid')
  await f.binding.writeExistingPageChange(f.record, undefined)
  await expect(f.binding.writeExistingPageChange({ ...f.record, state: 'staged', newSlideId: 's3' }, f.record)).rejects.toThrow('state_invalid')
  await expect(f.binding.writeExistingPageChange({ ...f.record, oldSlideId: 's2' }, f.record)).rejects.toThrow('state_invalid')
  await expect(f.binding.writeExistingPageChange({ ...f.record, state: 'inserted', newSlideId: 's3' }, undefined)).rejects.toThrow('stale')
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
  await f.binding.writeExistingPageChange({ ...f.record, changeId: 'page2' }, undefined)
  expect(f.reopen().listChangeHistory()).toMatchObject([
    { kind: 'existing_page', record: discarded },
    { kind: 'existing_page', record: { changeId: 'page2', state: 'pending' } },
  ])
})

it('blocks concurrent changes and rolls back a failed save', async () => {
  const f = await fixture()
  await f.binding.writeExistingPageChange(f.record, undefined)
  await expect(f.binding.writeExistingPageChange({ ...f.record, changeId: 'page2' }, undefined)).rejects.toThrow('pending')
  const before = new Map(f.values)
  f.save.mockRejectedValueOnce(new Error('save failed'))
  await expect(f.binding.writeExistingPageChange({ ...f.record, state: 'inserted', newSlideId: 's3' }, f.record)).rejects.toThrow('save failed')
  expect([...f.values]).toEqual([...before])
  expect(f.reopen().readExistingPageChange('page1')).toEqual(f.record)
})

it('stores each staged page review and invalidates both at commit', async () => {
  const f = await fixture()
  const inserted = { ...f.record, state: 'inserted' as const, newSlideId: 's3' }
  const staged = { ...inserted, state: 'staged' as const }
  const base = { screenshotDigest: 'e'.repeat(64), capturedAt: '2026-09-24T00:00:00.000Z', reviewedAt: '2026-09-24T00:01:00.000Z', status: 'pass' as const, notes: 'checked' }
  await f.binding.writeExistingPageChange(f.record, undefined)
  await f.binding.writeExistingPageChange(inserted, f.record)
  await f.binding.writeExistingPageChange(staged, inserted)
  const first = { ...staged, reviews: [{ ...base, hostSlideId: 's1' }] }
  const both = { ...staged, reviews: [...first.reviews, { ...base, hostSlideId: 's3' }] }
  await f.binding.writeExistingPageChange(first, staged)
  await f.reopen().writeExistingPageChange(both, first)
  await expect(f.binding.writeExistingPageChange({ ...both, state: 'commit_pending' }, both)).rejects.toThrow('state_invalid')
  const pending = { ...staged, state: 'commit_pending' as const }
  await f.binding.writeExistingPageChange(pending, both)
  expect(f.reopen().readExistingPageChange('page1')?.reviews).toBeUndefined()
})
