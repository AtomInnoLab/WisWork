import { expect, it, vi } from 'vitest'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import {
  validatePresentationGeometryChange,
  type PresentationGeometryChange,
} from '../src/skills/powerpoint/presentation-geometry-change'
const key = 'wiswork.presentation.geometry-change.v1'
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
  const record: PresentationGeometryChange = {
    version: 1,
    changeId: 'change-1',
    documentId: await binding.documentId(),
    projectId: 'project',
    requestId: 'request',
    artifactDigest: 'a'.repeat(64),
    pageId: 'page',
    hostSlideId: 'host',
    shapeId: 'shape',
    before: { left: 0, top: 0, width: 100, height: 50 },
    after: { left: 10, top: 20, width: 200, height: 100 },
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
it('validates bounded strict identity, geometry, state and source', async () => {
  const { record } = await fixture()
  expect(validatePresentationGeometryChange(record)).toBe(true)
  expect(validatePresentationGeometryChange({ ...record, source: 'production' })).toBe(true)
  for (const patch of [
    { extra: true },
    { version: 2 },
    { source: 'legacy' },
    { changeId: '../bad' },
    { documentId: '' },
    { shapeId: 'bad\n' },
    { artifactDigest: 'x'.repeat(64) },
    { state: 'complete' },
    { before: { ...record.before, width: -1 } },
    { after: { ...record.after, left: 100001 } },
    { after: { ...record.after, height: NaN } },
    { before: { ...record.before, extra: 1 } },
    { documentId: 'a'.repeat(4097) },
  ])
    expect(validatePresentationGeometryChange({ ...record, ...patch })).toBe(false)
})
it('persists the monotonic lifecycle, idempotency and replaces only ended records', async () => {
  const f = await fixture()
  expect(f.binding.readGeometryChange()).toBeUndefined()
  let prior: PresentationGeometryChange | undefined
  for (const state of ['pending', 'applied', 'undo_pending', 'undone'] as const) {
    const next = { ...f.record, state }
    await f.binding.writeGeometryChange(next, prior)
    expect(f.create().readGeometryChange()).toEqual(next)
    const saves = f.save.mock.calls.length
    await f.binding.writeGeometryChange(next, next)
    expect(f.save.mock.calls.length).toBe(saves)
    prior = next
  }
  const next = { ...f.record, changeId: 'change-2' }
  await f.binding.writeGeometryChange(next, prior)
  expect(f.create().readGeometryChange()).toEqual(next)
})
it('rejects stale CAS, skipped transitions, content rewrites and pending replacement', async () => {
  const f = await fixture()
  await f.binding.writeGeometryChange(f.record, undefined)
  await expect(
    f.binding.writeGeometryChange({ ...f.record, state: 'applied' }, undefined),
  ).rejects.toThrow('presentation_geometry_change_stale')
  for (const patch of [
    { state: 'undone' as const },
    { changeId: 'other' },
    { source: 'production' as const },
    { after: { ...f.record.after, left: 999 } },
  ])
    await expect(
      f.binding.writeGeometryChange({ ...f.record, ...patch }, f.record),
    ).rejects.toThrow('presentation_geometry_change_state_invalid')
  const applied = { ...f.record, state: 'applied' as const }
  await f.binding.writeGeometryChange(applied, f.record)
  const undo = { ...f.record, state: 'undo_pending' as const }
  await f.binding.writeGeometryChange(undo, applied)
  await expect(
    f.binding.writeGeometryChange({ ...f.record, changeId: 'other' }, undo),
  ).rejects.toThrow('presentation_geometry_change_state_invalid')
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
  const first = f.binding.writeGeometryChange(f.record, undefined)
  await vi.waitFor(() => expect(release).toBeTypeOf('function'))
  const expected = structuredClone(f.record),
    next = { ...structuredClone(f.record), state: 'applied' as const }
  const second = f.binding.writeGeometryChange(next, expected)
  next.after.left = 999
  expected.after.left = 999
  release()
  await first
  await second
  expect(f.binding.readGeometryChange()?.after.left).toBe(10)
  const results = await Promise.allSettled([
    f.binding.writeGeometryChange(
      { ...f.record, state: 'undo_pending' },
      { ...f.record, state: 'applied' },
    ),
    f.binding.writeGeometryChange(
      { ...f.record, changeId: 'other' },
      { ...f.record, state: 'applied' },
    ),
  ])
  expect(results.map((result) => result.status)).toEqual(['fulfilled', 'rejected'])
})
it('restores previous settings on failure and locks the binding on document switch', async () => {
  const f = await fixture()
  f.save.mockRejectedValueOnce(new Error('save_failed'))
  await expect(f.binding.writeGeometryChange(f.record, undefined)).rejects.toThrow('save_failed')
  expect(f.create().readGeometryChange()).toBeUndefined()
  await f.binding.writeGeometryChange(f.record, undefined)
  f.save.mockImplementationOnce(async () => {
    f.move()
  })
  await expect(
    f.binding.writeGeometryChange({ ...f.record, state: 'applied' }, f.record),
  ).rejects.toThrow('presentation_document_changed')
  expect(() => f.binding.readGeometryChange()).toThrow('presentation_geometry_change_state_invalid')
})
it('rejects corrupted or oversized settings and wrong document identity', async () => {
  const f = await fixture()
  await expect(
    f.binding.writeGeometryChange({ ...f.record, documentId: 'other' }, undefined),
  ).rejects.toThrow('presentation_document_changed')
  for (const raw of [
    '{',
    '{}',
    'null',
    ' '.repeat(16385),
    JSON.stringify({ ...f.record, extra: true }),
  ]) {
    f.values.set(key, raw)
    expect(() => f.create().readGeometryChange()).toThrow(
      'presentation_geometry_change_state_invalid',
    )
  }
})
it('retains a pending marker when completion save fails and locks after a changed document identity', async () => {
  const f = await fixture()
  await expect(
    f.binding.writeGeometryChange({ ...f.record, state: 'applied' }, undefined),
  ).rejects.toThrow('presentation_geometry_change_state_invalid')
  await f.binding.writeGeometryChange(f.record, undefined)
  f.save.mockRejectedValueOnce(new Error('save_failed'))
  await expect(
    f.binding.writeGeometryChange({ ...f.record, state: 'applied' }, f.record),
  ).rejects.toThrow('save_failed')
  expect(f.create().readGeometryChange()).toEqual(f.record)
  f.save.mockImplementationOnce(async () => {
    f.values.set('wiswork.presentation.document.v1', 'different')
  })
  await expect(
    f.binding.writeGeometryChange({ ...f.record, state: 'applied' }, f.record),
  ).rejects.toThrow('presentation_document_changed')
  expect(() => f.binding.readGeometryChange()).toThrow('presentation_geometry_change_state_invalid')
})
it('reserves the full JSON byte limit and locks if failed save rollback cannot restore settings', async () => {
  const f = await fixture()
  expect(
    validatePresentationGeometryChange({ ...f.record, documentId: '\u0000'.repeat(4096) }),
  ).toBe(false)
  await f.binding.writeGeometryChange(f.record, undefined)
  const original = f.settings.set
  f.save.mockImplementationOnce(async () => {
    f.settings.set = () => {
      throw new Error('set_failed')
    }
    throw new Error('save_failed')
  })
  await expect(
    f.binding.writeGeometryChange({ ...f.record, state: 'applied' }, f.record),
  ).rejects.toThrow('save_failed')
  f.settings.set = original
  expect(() => f.binding.readGeometryChange()).toThrow('presentation_geometry_change_state_invalid')
})
