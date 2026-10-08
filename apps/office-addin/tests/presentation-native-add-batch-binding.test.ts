import { expect, it, vi } from 'vitest'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import {
  type PresentationNativeAddBatch,
  validatePresentationExistingBatch,
  validExistingBatchTransition,
  existingBatchReservedBytes,
} from '../src/skills/powerpoint/presentation-existing-batch'
const operation = (index = 0) => ({
  op: 'add_text_box' as const,
  slide_index: 0,
  name: `text-${index}`,
  text: 'text',
  left: 1,
  top: 2,
  width: 30,
  height: 20,
})
async function fixture() {
  const values = new Map<string, string>(),
    save = vi.fn(async () => {})
  const settings = {
    get: (k: string) => values.get(k),
    set: (k: string, v: string) => {
      values.set(k, v)
    },
    save,
    location: () => 'synthetic',
  }
  const create = () => createPresentationDocumentBinding(settings, () => 'doc')
  const binding = create()
  const record = {
    version: 2 as const,
    kind: 'native_page_add' as const,
    changeId: 'add',
    documentId: await binding.documentId(),
    baselineId: 'baseline',
    baselineDigest: 'a'.repeat(64),
    hostSlideId: 'slide',
    slideIndex: 0,
    beforeSlideIds: ['slide'],
    scope: { slideIds: ['slide'] },
    intent: 'add objects',
    preserved: ['original objects'],
    validation: ['readback'],
    risk: 'high' as const,
    backups: [
      {
        hostSlideId: 'slide',
        backupId: 'backup',
        sha256: 'b'.repeat(64),
        sizeBytes: 100,
        packageDigest: 'a'.repeat(64),
      },
    ],
    operations: [operation(), operation(1)],
    nextIndex: 0,
    createdShapeIds: [] as string[],
    state: 'applying' as const,
  }
  return { binding, create, values, save, record }
}
it('persists exact native add start/inflight/completion across reopen and lost ACK without fake target IDs', async () => {
  const f = await fixture(),
    r = f.record
  expect(validatePresentationExistingBatch(r)).toBe(true)
  await f.binding.writeExistingBatch(r, undefined)
  const pending = { ...r, inFlightIndex: 0 }
  await f.binding.writeExistingBatch(pending, r)
  expect(f.create().readExistingBatch('add')).toEqual(pending)
  const { inFlightIndex: _omit, ...rest } = pending
  const first = { ...rest, nextIndex: 1, createdShapeIds: ['real-1'] }
  await f.binding.writeExistingBatch(first, pending)
  const secondPending = { ...first, inFlightIndex: 1 }
  await f.binding.writeExistingBatch(secondPending, first)
  const end = {
    ...first,
    nextIndex: 2,
    createdShapeIds: ['real-1', 'real-2'],
    state: 'applied' as const,
  }
  await f.binding.writeExistingBatch(end, secondPending)
  const saves = f.save.mock.calls.length
  await f.create().writeExistingBatch(end, end)
  expect(f.save).toHaveBeenCalledTimes(saves)
  expect(f.create().readExistingBatch('add')).toEqual(end)
  expect(f.create().listChangeHistory()[0]?.record).toEqual(end)
})
it('rejects invalid shapes, scope, real identities and nonmonotonic progress', async () => {
  const { record: r } = await fixture()
  for (const patch of [
    { cursor: 0 },
    { operations: [] },
    { operations: Array.from({ length: 33 }, (_, i) => operation(i)) },
    { operations: [{ ...operation(), shape_id: 'fake' }] },
    { operations: [{ ...operation(), bold: 'yes' }] },
    { operations: [{ ...operation(), op: 'delete_shape' }] },
    { scope: { slideIds: ['other'] } },
    { baselineDigest: 'c'.repeat(64) },
    { createdShapeIds: ['x'] },
    { inFlightIndex: 1 },
    { state: 'applied' },
  ])
    expect(validatePresentationExistingBatch({ ...r, ...patch })).toBe(false)
  const pending = { ...r, inFlightIndex: 0 }
  expect(validExistingBatchTransition(r, { ...r, intent: 'replaced' })).toBe(false)
  expect(validExistingBatchTransition(r, { ...r, nextIndex: 1, createdShapeIds: ['x'] })).toBe(
    false,
  )
  expect(
    validExistingBatchTransition(pending, {
      ...r,
      nextIndex: 2,
      createdShapeIds: ['x', 'y'],
      state: 'applied',
    }),
  ).toBe(false)
})
it('rolls back failed pending save and rejects stale CAS while retaining original across restart', async () => {
  const f = await fixture(),
    r = f.record
  await f.binding.writeExistingBatch(r, undefined)
  f.save.mockRejectedValueOnce(Error('save_failed'))
  await expect(f.binding.writeExistingBatch({ ...r, inFlightIndex: 0 }, r)).rejects.toThrow(
    'save_failed',
  )
  expect(f.create().readExistingBatch('add')).toEqual(r)
  await expect(f.binding.writeExistingBatch({ ...r, inFlightIndex: 0 }, undefined)).rejects.toThrow(
    'presentation_existing_batch_stale',
  )
})
it('reserves UTF8 maximum terminal identities without raising the original 192KiB cap', async () => {
  const { record: r } = await fixture()
  const full = {
    ...r,
    operations: Array.from({ length: 32 }, (_, i) => ({
      ...operation(i),
      text: '界'.repeat(1400),
    })),
  }
  expect(validatePresentationExistingBatch(full)).toBe(true)
  expect(existingBatchReservedBytes(full)).toBeGreaterThan(32 * 256 * 6)
  const huge = {
    ...full,
    operations: full.operations.map((op) => ({ ...op, text: '界'.repeat(1900) })),
  }
  expect(validatePresentationExistingBatch(huge)).toBe(false)
})
it('keeps every real created identity through whole-page restore and immutable terminal retry', async () => {
  const f = await fixture(),
    r = f.record
  await f.binding.writeExistingBatch(r, undefined)
  const pending = { ...r, inFlightIndex: 0 }
  await f.binding.writeExistingBatch(pending, r)
  const progress = { ...r, nextIndex: 1, createdShapeIds: ['real'] }
  await f.binding.writeExistingBatch(progress, pending)
  const restoring = { ...progress, state: 'undoing' as const }
  await f.binding.writeExistingBatch(restoring, progress)
  const restored = { ...restoring, state: 'undone' as const, restoredSlideId: 'restored' }
  await f.binding.writeExistingBatch(restored, restoring)
  expect(f.create().readExistingBatch('add')).toEqual(restored)
  expect(validExistingBatchTransition(restored, r)).toBe(false)
  expect(
    validExistingBatchTransition(restored, { ...restored, restoredSlideId: 'replacement' }),
  ).toBe(false)
  const released = { ...restored, backupReleasedAt: '2026-09-29T00:00:00.000Z' }
  await f.binding.writeExistingBatch(released, restored)
  await f.create().writeExistingBatch(released, released)
})
it('validates every real action field and rejects foreign scope or corrupted persisted records', async () => {
  const f = await fixture(),
    r = f.record
  const table = {
    op: 'add_native_table' as const,
    slide_index: 0,
    name: 'table',
    rows: [['', 'value']],
    left: 0,
    top: 0,
    width: 10,
    height: 10,
    fontFace: 'Arial',
    fontSize: 16,
    color: 'abcdef',
    borderColor: 'aBcDeF',
    cellMargin: 0.04,
  }
  const shape = {
    op: 'add_geometric_shape' as const,
    slide_index: 0,
    name: 'shape',
    left: 0,
    top: 0,
    width: 10,
    height: 10,
    shape: 'rect' as const,
    fill: 'abcdef',
    lineColor: 'fedcba',
  }
  expect(validatePresentationExistingBatch({ ...r, operations: [operation(), shape, table] })).toBe(
    true,
  )
  for (const op of [
    { ...table, fontSize: 49 },
    { ...table, rows: [['a'], ['b', 'c']] },
    { ...table, cellMargin: Infinity },
    { ...shape, fill: 'transparent' },
    { ...shape, shape: 'freeform' },
    { ...operation(), margin: 73 },
    { ...operation(), verticalAlignment: 'unknown' },
    { ...operation(), color: undefined },
    { ...operation(), width: 0 },
    { ...operation(), slide_index: 1 },
  ])
    expect(validatePresentationExistingBatch({ ...r, operations: [op] })).toBe(false)
  await expect(
    f.binding.writeExistingBatch({ ...r, documentId: 'other' }, undefined),
  ).rejects.toThrow('presentation_document_changed')
  await f.binding.writeExistingBatch(r, undefined)
  f.values.set(
    'wiswork.presentation.existing-batch.v1',
    JSON.stringify({ ...r, createdShapeIds: ['fake'] }),
  )
  expect(() => f.create().readExistingBatch('add')).toThrow(
    'presentation_existing_batch_state_invalid',
  )
})
it.each(['界', '\ud800'])(
  'allows reserved terminal identities including JSON escapes (%j) at the original quota',
  async (unit) => {
    const f = await fixture()
    let current: PresentationNativeAddBatch = {
      ...f.record,
      operations: Array.from({ length: 32 }, (_, i) => ({
        ...operation(i),
        text: '界'.repeat(1400),
      })),
    }
    // Fill the remaining budget with valid text before any host identity has been saved.
    while (
      validatePresentationExistingBatch({
        ...current,
        operations: current.operations.map((op) => ({
          ...op,
          text: (op as Extract<typeof op, { op: 'add_text_box' }>).text + '界',
        })),
      })
    ) {
      current = {
        ...current,
        operations: current.operations.map((op) => ({
          ...op,
          text: (op as Extract<typeof op, { op: 'add_text_box' }>).text + '界',
        })),
      }
    }
    await f.binding.writeExistingBatch(current, undefined)
    for (let i = 0; i < 32; i++) {
      const pending = { ...current, inFlightIndex: i }
      await f.binding.writeExistingBatch(pending, current)
      const next = {
        ...current,
        nextIndex: i + 1,
        createdShapeIds: [
          ...current.createdShapeIds,
          String.fromCharCode(0xe000 + i) + unit.repeat(255),
        ],
        state: i === 31 ? ('applied' as const) : ('applying' as const),
      }
      await f.binding.writeExistingBatch(next, pending)
      current = next
    }
    const restoring = { ...current, state: 'undoing' as const }
    await f.binding.writeExistingBatch(restoring, current)
    const terminal = { ...restoring, state: 'undone' as const, restoredSlideId: unit.repeat(256) }
    await f.binding.writeExistingBatch(terminal, restoring)
    const released = { ...terminal, backupReleasedAt: '2026-09-29T00:00:00.000Z' }
    await f.binding.writeExistingBatch(released, terminal)
    expect(f.create().readExistingBatch('add')).toEqual(released)
    expect(new TextEncoder().encode(JSON.stringify(released)).byteLength).toBeLessThanOrEqual(
      192 * 1024,
    )
  },
)
