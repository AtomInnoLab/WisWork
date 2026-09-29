import { expect, it, vi } from 'vitest'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import {
  validatePresentationNativeMasterChange,
  type PresentationNativeMasterChange,
} from '../src/skills/powerpoint/presentation-native-master-change'

const ref = (key: string) => ({ key, sha256: 'a'.repeat(64), sizeBytes: 100 })
async function fixture() {
  const values = new Map<string, string>(),
    save = vi.fn(async () => {})
  let location = 'deck'
  const settings = {
    get: (k: string) => values.get(k),
    set: (k: string, v: string) => {
      values.set(k, v)
    },
    save,
    location: () => location,
  }
  const create = () => createPresentationDocumentBinding(settings, () => 'doc')
  const binding = create()
  const record: PresentationNativeMasterChange = {
    version: 1,
    kind: 'native_master',
    changeId: 'master-change',
    documentId: await binding.documentId(),
    intent: 'theme',
    snapshotRef: ref('snapshot'),
    operations: [
      { op: 'set_master_theme_color', master_id: 'm1', theme_color: 'Accent1', color: '#112233' },
    ],
    inverseOperations: [
      { op: 'set_master_theme_color', master_id: 'm1', theme_color: 'Accent1', color: '#FFFFFF' },
    ],
    scope: { masterIds: ['m1'], affectedPageCount: 600 },
    nextIndex: 0,
    state: 'applying',
    currentProofRef: ref('receipt-0'),
    receipts: [],
    reviews: [],
  }
  expect(validatePresentationNativeMasterChange(record)).toBe(true)
  return {
    binding,
    create,
    values,
    save,
    record,
    move: () => {
      location = 'other'
    },
  }
}
it('persists standalone master pending and actual receipt across reopen with checkpoint lineage', async () => {
  const f = await fixture(),
    r = f.record
  await f.binding.writeNativeMasterChange(r, undefined)
  const pending = {
    ...r,
    pending: { direction: 'forward' as const, index: 0, beforeProofRef: r.currentProofRef },
  }
  await f.binding.writeNativeMasterChange(pending, r)
  expect(f.create().readNativeMasterChange(r.changeId)).toEqual(pending)
  const observed = { ...pending, pending: { ...pending.pending, afterProofRef: ref('receipt-1') } }
  await f.binding.writeNativeMasterChange(observed, pending)
  const applied = {
    ...r,
    nextIndex: 1,
    state: 'applied' as const,
    currentProofRef: ref('receipt-1'),
    receipts: [{ direction: 'forward' as const, index: 0, proofRef: ref('receipt-1') }],
  }
  await f.binding.writeNativeMasterChange(applied, observed)
  expect(f.create().readNativeMasterChange(r.changeId)).toEqual(applied)
  expect(f.create().readExistingBatch(r.changeId)).toBeUndefined()
  expect(f.create().listChangeHistory()[0]).toMatchObject({
    kind: 'native_master',
    checkpointCreatedAt: expect.any(String),
    record: applied,
  })
  const count = f.save.mock.calls.length
  await f.create().writeNativeMasterChange(applied, applied)
  expect(f.save).toHaveBeenCalledTimes(count)
})
it('rejects immutable savepoint replacement and stale CAS without changing durable state', async () => {
  const f = await fixture(),
    r = f.record
  await f.binding.writeNativeMasterChange(r, undefined)
  await expect(
    f.binding.writeNativeMasterChange({ ...r, snapshotRef: ref('receipt-9') }, r),
  ).rejects.toThrow('presentation_native_master_state_invalid')
  await expect(f.binding.writeNativeMasterChange(r, undefined)).rejects.toThrow(
    'presentation_native_master_stale',
  )
  expect(f.create().readNativeMasterChange(r.changeId)).toEqual(r)
})
it('rolls back failed settings save before native mutation and remains readable on restart', async () => {
  const f = await fixture(),
    r = f.record
  await f.binding.writeNativeMasterChange(r, undefined)
  f.save.mockRejectedValueOnce(new Error('synthetic-save'))
  await expect(
    f.binding.writeNativeMasterChange(
      { ...r, pending: { direction: 'forward', index: 0, beforeProofRef: r.currentProofRef } },
      r,
    ),
  ).rejects.toThrow('synthetic-save')
  expect(f.create().readNativeMasterChange(r.changeId)).toEqual(r)
})
it('captures caller records before asynchronous settings writes and rejects Save As identity', async () => {
  const f = await fixture(),
    r = f.record
  const pending = f.binding.writeNativeMasterChange(r, undefined)
  r.intent = 'external alias mutation'
  r.scope.masterIds[0] = 'external'
  await pending
  expect(f.create().readNativeMasterChange(r.changeId)?.intent).toBe('theme')
  f.move()
  const saved = f.create().listChangeHistory()[0]?.record as PresentationNativeMasterChange
  await expect(f.create().writeNativeMasterChange(saved, saved)).rejects.toThrow(
    'presentation_document_changed',
  )
})
it('a pending master transaction blocks unrelated historical writes', async () => {
  const f = await fixture(),
    r = f.record
  await f.binding.writeNativeMasterChange(r, undefined)
  await expect(
    f.binding.writeTextChange(
      {
        version: 1,
        changeId: 'text',
        documentId: r.documentId,
        projectId: 'project',
        requestId: 'request',
        artifactDigest: 'b'.repeat(64),
        pageId: 'page',
        hostSlideId: 's1',
        shapeId: 'shape',
        before: 'before',
        after: 'after',
        state: 'pending',
      },
      undefined,
    ),
  ).rejects.toThrow('presentation_change_history_pending')
  expect(f.create().listChangeHistory()).toHaveLength(1)
})
it('confirmed unchanged pending closure records restoration without inventing a native write', async () => {
  const f = await fixture(),
    r = f.record
  await f.binding.writeNativeMasterChange(r, undefined)
  const p = {
    ...r,
    pending: { direction: 'forward' as const, index: 0, beforeProofRef: r.currentProofRef },
  }
  await f.binding.writeNativeMasterChange(p, r)
  const undone = { ...r, state: 'undone' as const }
  await f.binding.writeNativeMasterChange(undone, p)
  expect(f.create().listChangeHistory()[0]).toMatchObject({
    kind: 'native_master',
    checkpointRestoredAt: expect.any(String),
    record: { state: 'undone', nextIndex: 0, receipts: [] },
  })
})

it('checks cached document identity synchronously at the native write boundary and rejects Save As', async () => {
  const f = await fixture()
  expect(() => f.binding.assertDocumentId(f.record.documentId)).not.toThrow()
  f.move()
  expect(() => f.binding.assertDocumentId(f.record.documentId)).toThrow(
    'presentation_document_changed',
  )
  expect(f.save).toHaveBeenCalledOnce()
})
