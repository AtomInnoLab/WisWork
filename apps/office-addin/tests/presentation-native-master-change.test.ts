import { expect, it } from 'vitest'
import {
  validatePresentationNativeMasterChange,
  validNativeMasterTransition,
  nativeMasterReservedBytes,
  type PresentationNativeMasterChange,
} from '../src/skills/powerpoint/presentation-native-master-change.js'
const ref = (key: string) => ({ key, sha256: 'a'.repeat(64), sizeBytes: 1 })
const op = {
  op: 'set_master_theme_color' as const,
  master_id: 'm',
  theme_color: 'Accent1',
  color: '#000000',
}
const base = (): PresentationNativeMasterChange => ({
  version: 1,
  kind: 'native_master',
  changeId: 'change',
  documentId: 'doc',
  intent: 'edit',
  snapshotRef: ref('snapshot'),
  operations: [op],
  inverseOperations: [{ ...op, color: '#FFFFFF' }],
  scope: { masterIds: ['m'], affectedPageCount: 600 },
  nextIndex: 0,
  state: 'applying',
  currentProofRef: ref('receipt-0'),
  receipts: [],
  reviews: [],
})
it('accepts more than512 affected pages using compact PC references and32 unique operations', () => {
  const r = base()
  r.operations = Array.from({ length: 32 }, (_, i) => ({ ...op, master_id: `m${i}` }))
  r.inverseOperations = r.operations.map((o) => ({ ...o, color: '#FFFFFF' }))
  r.scope.masterIds = r.operations.map((o) => o.master_id)
  expect(validatePresentationNativeMasterChange(r)).toBe(true)
  expect(nativeMasterReservedBytes(r)).toBeLessThan(192 * 1024)
})
it('rejects unknown fields, raw images, duplicate target keys and malformed references', () => {
  expect(validatePresentationNativeMasterChange({ ...base(), extra: true })).toBe(false)
  expect(
    validatePresentationNativeMasterChange({
      ...base(),
      snapshotRef: { ...ref('snapshot'), sha256: 'bad' },
    }),
  ).toBe(false)
  expect(
    validatePresentationNativeMasterChange({
      ...base(),
      operations: [op, op],
      inverseOperations: [op, op],
    }),
  ).toBe(false)
  expect(
    validatePresentationNativeMasterChange({
      ...base(),
      operations: [
        {
          op: 'set_master_background',
          master_id: 'm',
          fill: { type: 'picture_or_texture', image_base64: 'private', transparency: 0 },
        },
      ],
    }),
  ).toBe(false)
})
it('requires durable pending and observed proof before recording host progress', () => {
  const r = base(),
    pending = {
      ...r,
      pending: { direction: 'forward' as const, index: 0, beforeProofRef: r.currentProofRef },
    }
  expect(validNativeMasterTransition(undefined, r)).toBe(true)
  expect(validNativeMasterTransition(r, pending)).toBe(true)
  const advanced = {
    ...r,
    state: 'applied' as const,
    nextIndex: 1,
    currentProofRef: ref('receipt-1'),
    receipts: [{ direction: 'forward' as const, index: 0, proofRef: ref('receipt-1') }],
  }
  expect(validNativeMasterTransition(r, advanced)).toBe(false)
  expect(validNativeMasterTransition(pending, advanced)).toBe(false)
  const observed = { ...pending, pending: { ...pending.pending, afterProofRef: ref('receipt-1') } }
  expect(validNativeMasterTransition(pending, observed)).toBe(true)
  expect(validNativeMasterTransition(observed, advanced)).toBe(true)
})
it('closes explicitly reconciled before-state into undo without replay and rejects journal forgery', () => {
  const r = base(),
    pending = {
      ...r,
      pending: { direction: 'forward' as const, index: 0, beforeProofRef: r.currentProofRef },
    }
  expect(validNativeMasterTransition(pending, { ...r, state: 'undone' })).toBe(true)
  expect(validNativeMasterTransition(pending, r)).toBe(false)
  expect(validNativeMasterTransition(r, { ...r, documentId: 'other' })).toBe(false)
  expect(validNativeMasterTransition(r, { ...r, currentProofRef: ref('other') })).toBe(false)
})
