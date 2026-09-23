import { expect, it, vi } from 'vitest'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import {
  validatePresentationPageReplacement,
  type PresentationPageReplacement,
} from '../src/skills/powerpoint/presentation-page-replacement-record'
const key = 'wiswork.presentation.page-replacement.v1'
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
  const record: PresentationPageReplacement = {
    version: 1,
    changeId: 'change',
    documentId: await binding.documentId(),
    projectId: 'project',
    parentRequestId: 'parent',
    requestId: 'child',
    pageId: 'page',
    backupId: 'backup',
    parentArtifactDigest: 'a'.repeat(64),
    backupDigest: 'b'.repeat(64),
    originalPackageDigest: 'c'.repeat(64),
    replacementPackageDigest: 'd'.repeat(64),
    sourceSlideId: '256#',
    oldSlideId: 'old',
    beforeSlideIds: ['first', 'old', 'last'],
    state: 'pending',
  }
  const inserted = { ...record, state: 'inserted' as const, newSlideId: 'new' }
  return {
    values,
    save,
    settings,
    create,
    binding,
    record,
    inserted,
    move: () => {
      location = 'file://other.pptx'
    },
  }
}
it('validates strict bounded identity, original order and new page lifecycle', async () => {
  const { record, inserted } = await fixture()
  expect(validatePresentationPageReplacement(record)).toBe(true)
  expect(validatePresentationPageReplacement(inserted)).toBe(true)
  for (const patch of [
    { extra: true },
    { version: 2 },
    { changeId: '../bad' },
    { parentRequestId: 'child' },
    { projectId: 'a'.repeat(81) },
    { backupId: '' },
    { documentId: '' },
    { oldSlideId: 'bad\n' },
    { sourceSlideId: '255#' },
    { backupDigest: 'x'.repeat(64) },
    { state: 'complete' },
    { beforeSlideIds: [] },
    { beforeSlideIds: ['first'] },
    { beforeSlideIds: ['old', 'old'] },
    { beforeSlideIds: Array.from({ length: 513 }, (_, i) => (i === 0 ? 'old' : String(i))) },
    { newSlideId: 'new' },
    { state: 'inserted' },
    { state: 'staged', newSlideId: 'old' },
  ])
    expect(validatePresentationPageReplacement({ ...record, ...patch })).toBe(false)
  const large = {
    ...record,
    beforeSlideIds: ['old', ...Array.from({ length: 511 }, (_, i) => String(i) + '界'.repeat(250))],
  }
  expect(validatePresentationPageReplacement(large)).toBe(false)
})
it('persists monotonic lifecycle and idempotency; permits only discarded predecessor replacement', async () => {
  const f = await fixture()
  let previous: PresentationPageReplacement | undefined
  for (const state of ['pending', 'inserted', 'staged', 'discard_pending', 'discarded'] as const) {
    const next = state === 'pending' ? f.record : { ...f.inserted, state }
    await f.binding.writePageReplacement(next, previous)
    expect(f.create().readPageReplacement()).toEqual(next)
    const saves = f.save.mock.calls.length
    await f.binding.writePageReplacement(next, next)
    expect(f.save.mock.calls.length).toBe(saves)
    if (state !== 'discarded')
      await expect(
        f.binding.writePageReplacement(
          { ...f.record, changeId: 'other', projectId: 'other' },
          next,
        ),
      ).rejects.toThrow('presentation_page_replacement_state_invalid')
    previous = next
  }
  await f.binding.writePageReplacement({ ...f.record, changeId: 'other' }, previous)
  expect(f.create().readPageReplacement()?.changeId).toBe('other')
})
it('rejects stale CAS, skipped transitions and changes to frozen fields or assigned new page ID', async () => {
  const f = await fixture()
  await f.binding.writePageReplacement(f.record, undefined)
  await expect(f.binding.writePageReplacement(f.inserted, undefined)).rejects.toThrow(
    'presentation_page_replacement_stale',
  )
  for (const next of [
    { ...f.inserted, state: 'staged' as const },
    { ...f.inserted, backupDigest: 'e'.repeat(64) },
    { ...f.inserted, beforeSlideIds: ['old', 'first', 'last'] },
  ])
    await expect(f.binding.writePageReplacement(next, f.record)).rejects.toThrow(
      'presentation_page_replacement_state_invalid',
    )
  await f.binding.writePageReplacement(f.inserted, f.record)
  await expect(
    f.binding.writePageReplacement(
      { ...f.inserted, state: 'staged', newSlideId: 'another' },
      f.inserted,
    ),
  ).rejects.toThrow('presentation_page_replacement_state_invalid')
})
it('clones before queueing and serializes competing writes', async () => {
  const f = await fixture()
  let release!: () => void
  f.save.mockImplementationOnce(
    () =>
      new Promise<void>((r) => {
        release = r
      }),
  )
  const first = f.binding.writePageReplacement(f.record, undefined)
  await vi.waitFor(() => expect(release).toBeTypeOf('function'))
  const expected = structuredClone(f.record),
    next = structuredClone(f.inserted)
  const second = f.binding.writePageReplacement(next, expected)
  expected.beforeSlideIds.reverse()
  next.beforeSlideIds.reverse()
  release()
  await first
  await second
  expect(f.binding.readPageReplacement()).toEqual(f.inserted)
  const result = await Promise.allSettled([
    f.binding.writePageReplacement({ ...f.inserted, state: 'staged' }, f.inserted),
    f.binding.writePageReplacement({ ...f.inserted, state: 'staged' }, f.inserted),
  ])
  expect(result.map((r) => r.status)).toEqual(['fulfilled', 'rejected'])
})
it('restores prior marker on failed saves and retains pending after insertion receipt failure', async () => {
  const f = await fixture()
  f.save.mockRejectedValueOnce(new Error('save_failed'))
  await expect(f.binding.writePageReplacement(f.record, undefined)).rejects.toThrow('save_failed')
  expect(f.create().readPageReplacement()).toBeUndefined()
  await f.binding.writePageReplacement(f.record, undefined)
  f.save.mockRejectedValueOnce(new Error('save_failed'))
  await expect(f.binding.writePageReplacement(f.inserted, f.record)).rejects.toThrow('save_failed')
  expect(f.create().readPageReplacement()).toEqual(f.record)
})
it('locks after URL, document identity or journal changes during save', async () => {
  for (const change of ['location', 'identity', 'journal']) {
    const f = await fixture()
    await f.binding.writePageReplacement(f.record, undefined)
    f.save.mockImplementationOnce(async () => {
      if (change === 'location') f.move()
      else
        f.values.set(change === 'identity' ? 'wiswork.presentation.document.v1' : key, 'different')
    })
    await expect(f.binding.writePageReplacement(f.inserted, f.record)).rejects.toThrow(
      'presentation_document_changed',
    )
    expect(() => f.binding.readPageReplacement()).toThrow(
      'presentation_page_replacement_state_invalid',
    )
  }
})
it('rejects corrupt settings and wrong document identity', async () => {
  const f = await fixture()
  await expect(
    f.binding.writePageReplacement({ ...f.record, documentId: 'other' }, undefined),
  ).rejects.toThrow('presentation_document_changed')
  await expect(f.binding.writePageReplacement(f.inserted, undefined)).rejects.toThrow(
    'presentation_page_replacement_state_invalid',
  )
  for (const raw of [
    '{',
    '{}',
    'null',
    ' '.repeat(192 * 1024 + 1),
    JSON.stringify({ ...f.record, extra: true }),
  ]) {
    f.values.set(key, raw)
    expect(() => f.create().readPageReplacement()).toThrow(
      'presentation_page_replacement_state_invalid',
    )
  }
})
it('locks if rollback cannot restore settings', async () => {
  const f = await fixture()
  await f.binding.writePageReplacement(f.record, undefined)
  const original = f.settings.set
  f.save.mockImplementationOnce(async () => {
    f.settings.set = () => {
      throw new Error('set_failed')
    }
    throw new Error('save_failed')
  })
  await expect(f.binding.writePageReplacement(f.inserted, f.record)).rejects.toThrow('save_failed')
  f.settings.set = original
  expect(() => f.binding.readPageReplacement()).toThrow(
    'presentation_page_replacement_state_invalid',
  )
})

async function committedFixture() {
  const f = await fixture()
  const parentReceipt = {
    state: 'complete' as const,
    documentId: f.record.documentId,
    slideIds: ['old'],
    checkpoint: {
      version: 2 as const,
      pageIds: ['page'],
      artifactDigest: f.record.parentArtifactDigest,
      sourceSlideIds: ['256#'],
      baselineSlideIds: ['first'],
      completed: [{ sourceSlideId: '256#', slideId: 'old' }],
    },
  }
  const childReceipt = structuredClone(parentReceipt)
  childReceipt.slideIds = ['new']
  childReceipt.checkpoint.completed[0].slideId = 'new'
  childReceipt.checkpoint.artifactDigest = 'e'.repeat(64)
  const staged = { ...f.inserted, state: 'staged' as const }
  const commit = { ...staged, state: 'commit_pending' as const, parentReceipt, childReceipt }
  await f.binding.writeReceipt('production/project/parent', parentReceipt)
  await f.binding.writePageReplacement(f.record, undefined)
  await f.binding.writePageReplacement(f.inserted, f.record)
  await f.binding.writePageReplacement(staged, f.inserted)
  return { ...f, staged, commit, parentReceipt, childReceipt }
}
it('atomically switches complete receipt mappings and restores parent mapping on undo', async () => {
  const f = await committedFixture()
  await f.binding.writePageReplacement(f.commit, f.staged)
  const applied = { ...f.commit, state: 'applied' as const }
  await f.binding.writePageReplacement(applied, f.commit)
  expect(() => f.create().readReceipt('production/project/parent')).toThrow(
    'presentation_import_superseded',
  )
  expect(f.create().readReceipt('production/project/child')).toEqual(f.childReceipt)
  await expect(f.binding.writeReceipt('production/project/child', undefined)).rejects.toThrow(
    'presentation_import_superseded',
  )
  const pending = { ...applied, state: 'undo_pending' as const }
  const restored = { ...pending, state: 'restore_inserted' as const, restoredSlideId: 'restored' }
  const undone = { ...restored, state: 'undone' as const }
  await f.binding.writePageReplacement(pending, applied)
  await f.binding.writePageReplacement(restored, pending)
  await f.binding.writePageReplacement(undone, restored)
  expect(f.create().readReceipt('production/project/parent')?.slideIds).toEqual(['restored'])
  expect(() => f.create().readReceipt('production/project/child')).toThrow(
    'presentation_import_superseded',
  )
  await f.binding.writePageReplacement({ ...f.record, changeId: 'next' }, undone)
  expect(f.create().readReceipt('production/project/parent')?.slideIds).toEqual(['restored'])
})
it('keeps mappings unchanged if terminal save fails and rejects corrupt overlays/Save As reads', async () => {
  const f = await committedFixture()
  await f.binding.writePageReplacement(f.commit, f.staged)
  f.save.mockRejectedValueOnce(new Error('save_failed'))
  await expect(
    f.binding.writePageReplacement({ ...f.commit, state: 'applied' }, f.commit),
  ).rejects.toThrow('save_failed')
  expect(f.create().readReceipt('production/project/parent')).toEqual(f.parentReceipt)
  await f.binding.writePageReplacement({ ...f.commit, state: 'applied' }, f.commit)
  const envelope = JSON.parse(f.values.get(key)!)
  envelope.receipts['production/project/parent'] = f.parentReceipt
  f.values.set(key, JSON.stringify(envelope))
  expect(() => f.create().readPageReplacement()).toThrow(
    'presentation_page_replacement_state_invalid',
  )
  f.values.set(key, JSON.stringify(f.record))
  expect(f.create().readPageReplacement()).toEqual(f.record)
  f.move()
  expect(() => f.create().readPageReplacement()).toThrow('presentation_document_changed')
})
it('preflights overlay capacity and validates frozen receipt identities before commit', async () => {
  const f = await committedFixture()
  for (const patch of [
    { parentReceipt: { ...f.parentReceipt, documentId: 'other' } },
    { childReceipt: { ...f.childReceipt, slideIds: ['wrong'] } },
    { restoredSlideId: 'early' },
  ])
    expect(validatePresentationPageReplacement({ ...f.commit, ...patch })).toBe(false)
  const envelope = JSON.parse(f.values.get(key)!)
  for (let i = 0; i < 31; i++) envelope.receipts[`production/history/task${i}`] = null
  f.values.set(key, JSON.stringify(envelope))
  await expect(f.binding.writePageReplacement(f.commit, f.staged)).rejects.toThrow(
    'presentation_import_history_full',
  )
  expect(f.binding.readPageReplacement()).toEqual(f.staged)
  expect(f.binding.readReceipt('production/project/parent')).toEqual(f.parentReceipt)
})
it('freezes pending commit receipt keys and immutable recovery receipt fields', async () => {
  const f = await committedFixture()
  await f.binding.writePageReplacement(f.commit, f.staged)
  await expect(f.binding.writeReceipt('production/project/parent', undefined)).rejects.toThrow(
    'presentation_import_superseded',
  )
  await expect(f.binding.writeReceipt('production/project/child', f.childReceipt)).rejects.toThrow(
    'presentation_import_superseded',
  )
  const changed = structuredClone(f.commit)
  changed.childReceipt.checkpoint.artifactDigest = 'f'.repeat(64)
  await expect(
    f.binding.writePageReplacement({ ...changed, state: 'applied' }, f.commit),
  ).rejects.toThrow('presentation_page_replacement_state_invalid')
})
it('reserves receipt byte budget before committing a large mapping', async () => {
  const f = await committedFixture()
  const baseline = Array.from({ length: 400 }, (_, i) => String(i) + 'x'.repeat(200))
  f.commit.parentReceipt.checkpoint.baselineSlideIds = baseline
  f.commit.childReceipt.checkpoint.baselineSlideIds = baseline
  await f.binding.writeReceipt('production/project/parent', f.commit.parentReceipt)
  const envelope = JSON.parse(f.values.get(key)!)
  envelope.receipts['production/history/task'] = f.commit.parentReceipt
  f.values.set(key, JSON.stringify(envelope))
  expect(validatePresentationPageReplacement(f.commit)).toBe(true)
  await expect(f.binding.writePageReplacement(f.commit, f.staged)).rejects.toThrow(
    'presentation_import_history_full',
  )
  expect(f.binding.readPageReplacement()).toEqual(f.staged)
})
