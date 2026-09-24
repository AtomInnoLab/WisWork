import { expect, it } from 'vitest'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import {
  validatePresentationExistingChartChange,
  validExistingChartTransition,
  type PresentationExistingChartChange,
} from '../src/skills/powerpoint/presentation-existing-chart'

async function fixture() {
  const values = new Map<string, string>()
  const binding = createPresentationDocumentBinding({
    get: (key) => values.get(key),
    set: (key, value) => { values.set(key, value) },
    save: async () => {},
    location: () => 'deck',
  }, () => 'doc')
  const record: PresentationExistingChartChange = {
    version: 1,
    changeId: 'chart_one',
    documentId: await binding.documentId(),
    oldSlideId: 'slide-1',
    shapeId: '7',
    slideIndex: 0,
    beforeSlideIds: ['slide-1', 'slide-2'],
    beforePackageDigest: 'a'.repeat(64),
    afterPackageDigest: 'b'.repeat(64),
    backup: { backupId: 'backup', sha256: 'c'.repeat(64), sizeBytes: 1024 },
    state: 'pending',
  }
  return { binding, values, record }
}

it('persists chart write intent, applied result and undo as document-bound history', async () => {
  const { binding, record } = await fixture()
  await binding.writeExistingChartChange(record, undefined)
  const writing = { ...record, state: 'write_pending' as const }
  await binding.writeExistingChartChange(writing, record)
  const applied = { ...writing, state: 'applied' as const, newSlideId: 'slide-3' }
  await binding.writeExistingChartChange(applied, writing)
  const undoing = { ...applied, state: 'undo_pending' as const }
  await binding.writeExistingChartChange(undoing, applied)
  const undone = { ...undoing, state: 'undone' as const, restoredSlideId: 'slide-4' }
  await binding.writeExistingChartChange(undone, undoing)
  expect(binding.readExistingChartChange(record.changeId)).toEqual(undone)
  expect(binding.listChangeHistory()).toMatchObject([{ kind: 'existing_chart', record: undone }])
  await expect(binding.writeExistingChartChange(applied, writing)).rejects.toThrow('stale')
})

it('rejects malformed chart identity, invented transitions and allows cancellation before host mutation', async () => {
  const { binding, record } = await fixture()
  expect(validatePresentationExistingChartChange({ ...record, shapeId: 'NaN' })).toBe(false)
  expect(validatePresentationExistingChartChange({ ...record, slideIndex: 1 })).toBe(false)
  expect(validExistingChartTransition(record, { ...record, state: 'applied', newSlideId: 'slide-3' })).toBe(false)
  await binding.writeExistingChartChange(record, undefined)
  const cancelled = { ...record, state: 'cancelled' as const }
  await binding.writeExistingChartChange(cancelled, record)
  expect(binding.readExistingChartChange(record.changeId)?.state).toBe('cancelled')
})

it('records one backup release receipt only after a terminal chart change', async () => {
  const { binding, record } = await fixture()
  const releasedAt = '2026-09-24T09:00:00.000Z'
  await binding.writeExistingChartChange(record, undefined)
  const cancelled = { ...record, state: 'cancelled' as const }
  await binding.writeExistingChartChange(cancelled, record)
  const released = { ...cancelled, backupReleasedAt: releasedAt }
  expect(validExistingChartTransition(record, { ...record, backupReleasedAt: releasedAt })).toBe(false)
  expect(validExistingChartTransition(cancelled, released)).toBe(true)
  await binding.writeExistingChartChange(released, cancelled)
  expect(binding.readExistingChartChange(record.changeId)).toEqual(released)
  expect(validExistingChartTransition(released, { ...released, backupReleasedAt: '2026-09-24T10:00:00.000Z' })).toBe(false)
  expect(validatePresentationExistingChartChange({ ...released, backupReleasedAt: 'yesterday' })).toBe(false)
})
