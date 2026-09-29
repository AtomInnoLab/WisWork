import { expect, it, vi } from 'vitest'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
const ref = (key: string) => ({ key, sha256: 'a'.repeat(64), sizeBytes: 100 })
async function fixture() {
  const values = new Map<string, string>(),
    save = vi.fn(async () => {})
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
  const binding = create() as any
  const record = {
    version: 1 as const,
    kind: 'package_xml' as const,
    changeId: 'xml-change',
    documentId: await binding.documentId(),
    sourceKind: 'slide' as const,
    sourceSlideId: 'host-slide',
    packageSourceSlideId: '256#',
    intent: 'XML背景修改',
    snapshotRef: ref('snapshot'),
    originalRef: ref('page-0'),
    preparedRef: ref('page-1'),
    currentProofRef: ref('receipt-0'),
    state: 'prepared' as const,
    receipts: [],
    reviews: [],
  }
  return {
    binding,
    create: () => create() as any,
    values,
    save,
    record,
    move: () => {
      location = 'other'
    },
  }
}
it('persists package savepoints and import intent across reopen without page replacement routing', async () => {
  const f = await fixture(),
    r = f.record
  await f.binding.writePackageChange(r, undefined)
  const pending = { ...r, pending: { action: 'import', beforeProofRef: r.currentProofRef } }
  await f.binding.writePackageChange(pending, r)
  expect(f.create().readPackageChange(r.changeId)).toEqual(pending)
  expect(f.create().readExistingPageChange(r.changeId)).toBeUndefined()
  expect(f.create().listChangeHistory()[0]).toMatchObject({
    kind: 'package_xml',
    checkpointCreatedAt: expect.any(String),
    record: pending,
  })
  const count = f.save.mock.calls.length
  await f.create().writePackageChange(pending, pending)
  expect(f.save).toHaveBeenCalledTimes(count)
})
it('rejects immutable backup replacement and stale package CAS', async () => {
  const f = await fixture(),
    r = f.record
  await f.binding.writePackageChange(r, undefined)
  await expect(
    f.binding.writePackageChange({ ...r, snapshotRef: ref('receipt-9') }, r),
  ).rejects.toThrow('presentation_package_state_invalid')
  await expect(f.binding.writePackageChange(r, undefined)).rejects.toThrow(
    'presentation_package_stale',
  )
  expect(f.create().readPackageChange(r.changeId)).toEqual(r)
})
it('rolls back settings save failure before package host mutation', async () => {
  const f = await fixture(),
    r = f.record
  await f.binding.writePackageChange(r, undefined)
  f.save.mockRejectedValueOnce(Error('synthetic-save'))
  await expect(
    f.binding.writePackageChange(
      { ...r, pending: { action: 'import', beforeProofRef: r.currentProofRef } },
      r,
    ),
  ).rejects.toThrow('synthetic-save')
  expect(f.create().readPackageChange(r.changeId)).toEqual(r)
})
it('copies package caller input before async settings and rejects Save As', async () => {
  const f = await fixture(),
    r = f.record
  const pending = f.binding.writePackageChange(r, undefined)
  r.intent = 'external alias'
  r.originalRef.sha256 = 'f'.repeat(64)
  await pending
  const saved = f.create().readPackageChange(r.changeId)
  expect(saved.intent).toBe('XML背景修改')
  expect(saved.originalRef.sha256).toBe('a'.repeat(64))
  f.move()
  await expect(f.create().writePackageChange(saved, saved)).rejects.toThrow(
    'presentation_document_changed',
  )
})
it('prepared package transactions block unrelated writes until safely closed', async () => {
  const f = await fixture(),
    r = f.record
  await f.binding.writePackageChange(r, undefined)
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
it('rejects corrupt package head instead of dropping it from history', async () => {
  const f = await fixture(),
    r = f.record
  await f.binding.writePackageChange(r, undefined)
  f.values.set('wiswork.presentation.package-xml.v1', '{}')
  expect(() => f.create().readPackageChange(r.changeId)).toThrow(
    'presentation_package_state_invalid',
  )
})

it('records verified original-state closure after confirmed XML savepoint discard', async () => {
  const f = await fixture(),
    r = f.record
  await f.binding.writePackageChange(r, undefined)
  const discarded = { ...r, state: 'discarded' }
  await f.binding.writePackageChange(discarded, r)
  expect(f.create().listChangeHistory()[0]).toMatchObject({
    kind: 'package_xml',
    checkpointRestoredAt: expect.any(String),
    record: { state: 'discarded' },
  })
})

it('blocks unrelated transactions while an applied XML change has an uncertain restore intent', async () => {
  const f = await fixture()
  let current: any = f.record
  await f.binding.writePackageChange(current, undefined)
  async function acknowledge(action: string, state: string, index: number, slideId?: string) {
    const pending = { ...current, pending: { action, beforeProofRef: current.currentProofRef } }
    await f.binding.writePackageChange(pending, current)
    const observed = {
      ...pending,
      pending: {
        ...pending.pending,
        ...(slideId ? { insertedSlideId: slideId } : {}),
        afterProofRef: ref(`receipt-${index}`),
      },
    }
    await f.binding.writePackageChange(observed, pending)
    const next = {
      ...observed,
      ...(slideId ? { replacementSlideId: slideId } : {}),
      pending: undefined,
      state,
      currentProofRef: ref(`receipt-${index}`),
      receipts: [
        ...current.receipts,
        { action, proofRef: ref(`receipt-${index}`), ...(slideId ? { slideId } : {}) },
      ],
    }
    await f.binding.writePackageChange(next, observed)
    current = next
  }
  await acknowledge('import', 'staged', 1, 'inserted')
  await acknowledge('delete_source', 'applied', 2)
  const pending = {
    ...current,
    pending: { action: 'restore', beforeProofRef: current.currentProofRef },
  }
  await f.binding.writePackageChange(pending, current)
  const reopened = f.create()
  await expect(
    reopened.writePackageChange({ ...f.record, changeId: 'unrelated' }, undefined),
  ).rejects.toThrow('presentation_change_history_pending')
  expect(reopened.readPackageChange(current.changeId)).toEqual(pending)
})
