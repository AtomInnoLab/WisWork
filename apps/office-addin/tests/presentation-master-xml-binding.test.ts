import { expect, it, vi } from 'vitest'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document.js'
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
  const create = () => createPresentationDocumentBinding(settings, () => 'doc') as any
  const binding = create()
  const record = {
    version: 1,
    kind: 'master_xml',
    changeId: 'master-xml',
    documentId: await binding.documentId(),
    intent: '修改完整母版',
    sourceSlideId: 'source',
    packageSourceSlideId: '256#',
    snapshotRef: ref('snapshot'),
    preparedRef: ref('page-0'),
    currentProofRef: ref('receipt-0'),
    state: 'prepared',
    cursor: { phase: 'original_probe_stage', index: 0, substep: 0 },
    receiptCount: 0,
    scope: { originalMasterId: 'master', affectedPageCount: 600, originalLayoutCount: 2 },
    introducedMasterCount: 0,
    inventoryCleanupVerified: false,
    reviews: [],
  }
  return {
    values,
    save,
    binding,
    create,
    record,
    move: () => {
      location = 'other'
    },
  }
}
it('persists a compact complete-scope master XML savepoint across reopen', async () => {
  const f = await fixture()
  await f.binding.writeMasterXmlChange(f.record, undefined)
  expect(f.create().readMasterXmlChange(f.record.changeId)).toEqual(f.record)
  expect(f.create().readPackageChange(f.record.changeId)).toBeUndefined()
  expect(f.create().listChangeHistory()[0]).toMatchObject({
    kind: 'master_xml',
    checkpointCreatedAt: expect.any(String),
    record: { scope: { affectedPageCount: 600 } },
  })
  const count = f.save.mock.calls.length
  await f.create().writeMasterXmlChange(f.record, f.record)
  expect(f.save).toHaveBeenCalledTimes(count)
})
it('rejects immutable master backups and stale CAS without overwriting history', async () => {
  const f = await fixture()
  await f.binding.writeMasterXmlChange(f.record, undefined)
  await expect(
    f.binding.writeMasterXmlChange({ ...f.record, snapshotRef: ref('receipt-9') }, f.record),
  ).rejects.toThrow('presentation_master_xml_state_invalid')
  await expect(f.binding.writeMasterXmlChange(f.record, undefined)).rejects.toThrow(
    'presentation_master_xml_stale',
  )
  expect(f.create().readMasterXmlChange(f.record.changeId)).toEqual(f.record)
})
it('copies master XML inputs before async settings and rejects Save As', async () => {
  const f = await fixture()
  const pending = f.binding.writeMasterXmlChange(f.record, undefined)
  f.record.intent = 'caller alias'
  f.record.scope.affectedPageCount = 1
  await pending
  const saved = f.create().readMasterXmlChange(f.record.changeId)
  expect(saved.intent).toBe('修改完整母版')
  expect(saved.scope.affectedPageCount).toBe(600)
  f.move()
  await expect(f.create().writeMasterXmlChange(saved, saved)).rejects.toThrow(
    'presentation_document_changed',
  )
})
it('blocks unrelated page XML while a master XML transaction is prepared', async () => {
  const f = await fixture()
  await f.binding.writeMasterXmlChange(f.record, undefined)
  await expect(
    f.binding.writePackageChange(
      {
        version: 1,
        kind: 'package_xml',
        changeId: 'page-xml',
        documentId: f.record.documentId,
        intent: 'page',
        sourceKind: 'slide',
        sourceSlideId: 'source',
        packageSourceSlideId: '256#',
        snapshotRef: ref('snapshot'),
        originalRef: ref('page-0'),
        preparedRef: ref('page-1'),
        currentProofRef: ref('receipt-0'),
        state: 'prepared',
        receipts: [],
        reviews: [],
      },
      undefined,
    ),
  ).rejects.toThrow('presentation_change_history_pending')
  expect(f.create().listChangeHistory()).toHaveLength(1)
})
it('rejects corrupt master XML heads instead of silently dropping them', async () => {
  const f = await fixture()
  await f.binding.writeMasterXmlChange(f.record, undefined)
  f.values.set('wiswork.presentation.master-xml.v1', '{}')
  expect(() => f.create().readMasterXmlChange(f.record.changeId)).toThrow(
    'presentation_master_xml_state_invalid',
  )
})
