import { expect, it, vi } from 'vitest'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import {
  imageReplacementKey,
  validateImageReplacementRecord,
  type ImageReplacementRecord,
} from '../src/skills/powerpoint/presentation-image-replacement-record'
async function setup() {
  const values = new Map<string, string>(),
    save = vi.fn(async () => {})
  const settings = {
    get: (key: string) => values.get(key),
    set: (key: string, value: string) => {
      values.set(key, value)
    },
    save,
    location: () => 'file://image.pptx',
  }
  const create = () => createPresentationDocumentBinding(settings, () => 'doc')
  const binding = create(),
    record: ImageReplacementRecord = {
      version: 1,
      documentId: await binding.documentId(),
      projectId: 'project',
      requestId: 'request',
      pageId: 'page',
      hostSlideId: 'host',
      oldShapeId: 'old',
      assetDigest: 'a'.repeat(64),
      state: 'pending',
    }
  const key = await imageReplacementKey('project', 'request', 'page', 'old')
  return { binding, create, record, key, save, values }
}
it('persists pending, candidate and completion across recreated settings bindings', async () => {
  const f = await setup()
  await f.binding.writeImageReplacement(f.key, f.record)
  expect(f.create().readImageReplacement(f.key)).toEqual(f.record)
  const candidate = { ...f.record, newShapeId: 'new' }
  await f.binding.writeImageReplacement(f.key, candidate)
  const complete = { ...candidate, state: 'complete' as const }
  await f.binding.writeImageReplacement(f.key, complete)
  expect(f.create().readImageReplacement(f.key)).toEqual(complete)
  await expect(f.binding.writeImageReplacement(f.key, f.record)).rejects.toThrow(
    'presentation_image_replacement_state_invalid',
  )
})
it('keeps the pending marker after candidate save fails and prevents scope or candidate rewrites', async () => {
  const f = await setup()
  await f.binding.writeImageReplacement(f.key, f.record)
  f.save.mockRejectedValueOnce(new Error('save_failed'))
  await expect(
    f.binding.writeImageReplacement(f.key, { ...f.record, newShapeId: 'new' }),
  ).rejects.toThrow('save_failed')
  expect(f.binding.readImageReplacement(f.key)).toEqual(f.record)
  await expect(
    f.binding.writeImageReplacement(f.key, { ...f.record, assetDigest: 'b'.repeat(64) }),
  ).rejects.toThrow()
  await expect(f.binding.writeImageReplacement('f'.repeat(64), f.record)).rejects.toThrow()
  await f.binding.writeImageReplacement(f.key, { ...f.record, newShapeId: 'new' })
  await expect(
    f.binding.writeImageReplacement(f.key, { ...f.record, newShapeId: 'another' }),
  ).rejects.toThrow()
})
it('strictly validates records and rejects corrupted or oversized settings', async () => {
  const f = await setup()
  expect(validateImageReplacementRecord(f.record)).toBe(true)
  for (const record of [
    { ...f.record, extra: true },
    { ...f.record, state: 'complete' },
    { ...f.record, newShapeId: 'old' },
    { ...f.record, state: 'unknown' },
  ])
    expect(validateImageReplacementRecord(record)).toBe(false)
  f.values.set('wiswork.presentation.image-replacements.v1', ' '.repeat(128 * 1024 + 1))
  expect(() => f.binding.readImageReplacement(f.key)).toThrow(
    'presentation_image_replacement_state_invalid',
  )
  f.values.set(
    'wiswork.presentation.image-replacements.v1',
    JSON.stringify({ [f.key]: { ...f.record, extra: true } }),
  )
  expect(() => f.binding.readImageReplacement(f.key)).toThrow(
    'presentation_image_replacement_state_invalid',
  )
})
it('requires a saved candidate before completion and bounds history before insertion', async () => {
  const f = await setup()
  await expect(
    f.binding.writeImageReplacement(f.key, { ...f.record, state: 'complete', newShapeId: 'new' }),
  ).rejects.toThrow()
  for (let i = 0; i < 32; i++) {
    const record = { ...f.record, oldShapeId: `old-${i}` }
    await f.binding.writeImageReplacement(
      await imageReplacementKey(
        record.projectId,
        record.requestId,
        record.pageId,
        record.oldShapeId,
      ),
      record,
    )
  }
  await expect(f.binding.writeImageReplacement(f.key, f.record)).rejects.toThrow(
    'presentation_image_replacement_history_full',
  )
  const key = await imageReplacementKey('project', 'request', 'page', 'old-0')
  const candidate = { ...f.record, oldShapeId: 'old-0', newShapeId: '\\'.repeat(256) }
  await f.binding.writeImageReplacement(key, candidate)
  await f.binding.writeImageReplacement(key, { ...candidate, state: 'complete' })
  expect(f.create().readImageReplacement(key)?.state).toBe('complete')
})
it('reserves serialized space for candidate ids and snapshots queued input', async () => {
  const f = await setup()
  // Populate valid maximum-size document identities to approach the settings budget.
  const records: Record<string, ImageReplacementRecord> = {}
  for (let i = 0; i < 23; i++) {
    const record = { ...f.record, documentId: 'd'.repeat(4096), oldShapeId: `old-${i}` }
    records[
      await imageReplacementKey(
        record.projectId,
        record.requestId,
        record.pageId,
        record.oldShapeId,
      )
    ] = record
  }
  f.values.set('wiswork.presentation.image-replacements.v1', JSON.stringify(records))
  expect(() => f.binding.readImageReplacement(f.key)).toThrow(
    'presentation_image_replacement_state_invalid',
  )
  f.values.delete('wiswork.presentation.image-replacements.v1')
  const record = { ...f.record }
  const write = f.binding.writeImageReplacement(f.key, record)
  record.assetDigest = 'b'.repeat(64)
  await write
  expect(f.binding.readImageReplacement(f.key)?.assetDigest).toBe('a'.repeat(64))
})

const baseline = {
  slideId: 'host',
  shapeId: 'old',
  geometry: { left: 1, top: 2, width: 100, height: 50 },
  rotation: 0,
  name: 'Picture',
  altTextTitle: '',
  altTextDescription: '',
  zOrderPosition: 0,
  shapeIds: ['old', 'title'],
  pictureFingerprint: 'b'.repeat(64),
  mediaDigest: 'c'.repeat(64),
}
it('persists immutable recovery evidence and never upgrades legacy records with guessed evidence', async () => {
  const f = await setup(),
    record = { ...f.record, baseline }
  expect(validateImageReplacementRecord(record)).toBe(true)
  await f.binding.writeImageReplacement(f.key, record)
  await f.binding.writeImageReplacement(f.key, { ...record, newShapeId: 'new' })
  expect(f.create().readImageReplacement(f.key)?.baseline).toEqual(baseline)
  await expect(
    f.binding.writeImageReplacement(f.key, {
      ...record,
      newShapeId: 'new',
      baseline: { ...baseline, pictureFingerprint: 'd'.repeat(64) },
    }),
  ).rejects.toThrow()
  await expect(
    f.binding.writeImageReplacement(f.key, { ...record, newShapeId: 'new', baseline: undefined }),
  ).rejects.toThrow()
  const legacy = await setup()
  await legacy.binding.writeImageReplacement(legacy.key, legacy.record)
  await expect(
    legacy.binding.writeImageReplacement(legacy.key, { ...legacy.record, baseline }),
  ).rejects.toThrow()
})
it('rejects unbound, malformed, oversized or candidate-aliasing recovery evidence', async () => {
  const f = await setup()
  for (const invalid of [
    { ...baseline, slideId: 'other' },
    { ...baseline, shapeId: 'other' },
    { ...baseline, geometry: { ...baseline.geometry, width: -1 } },
    { ...baseline, shapeIds: ['old', 'old'] },
    { ...baseline, shapeIds: ['title', 'old'] },
    { ...baseline, zOrderPosition: 0.5 },
    { ...baseline, mediaDigest: 'bad' },
    { ...baseline, extra: true },
    { ...baseline, altTextTitle: 'a'.repeat(17000) },
  ])
    expect(validateImageReplacementRecord({ ...f.record, baseline: invalid })).toBe(false)
  expect(validateImageReplacementRecord({ ...f.record, baseline, newShapeId: 'title' })).toBe(false)
})
