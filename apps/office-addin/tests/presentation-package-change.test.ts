import { expect, it } from 'vitest'
import {
  validatePresentationPackageChange,
  validPackageTransition,
  packageChangeReservedBytes,
  type PresentationPackageChange,
} from '../src/skills/powerpoint/presentation-package-change.js'
const ref = (key: string) => ({ key, sha256: 'a'.repeat(64), sizeBytes: 100 })
const initial = (): PresentationPackageChange => ({
  version: 1,
  kind: 'package_xml',
  changeId: 'change',
  documentId: 'doc',
  intent: 'XML',
  sourceKind: 'slide',
  sourceSlideId: 'native',
  packageSourceSlideId: '256#',
  snapshotRef: ref('snapshot'),
  originalRef: ref('page-0'),
  preparedRef: ref('page-1'),
  currentProofRef: ref('receipt-0'),
  state: 'prepared',
  receipts: [],
  reviews: [],
})
it('validates dedicated initial journal and reserves futurebytes without pagecount caps', () => {
  const r = initial()
  expect(validatePresentationPackageChange(r)).toBe(true)
  expect(validPackageTransition(undefined, r)).toBe(true)
  expect(packageChangeReservedBytes(r)).toBeGreaterThan(0)
})
it('rejects invalid native/package identities, refs and phase receipt combinations', () => {
  for (const patch of [
    { sourceSlideId: ' ' },
    { sourceSlideId: 'native\u0000' },
    { sourceSlideId: 'x'.repeat(257) },
    { packageSourceSlideId: 'native' },
    { state: 'applied' },
    { currentProofRef: ref('receipt-1') },
    { snapshotRef: ref('receipt-0') },
  ])
    expect(validatePresentationPackageChange({ ...initial(), ...patch })).toBe(false)
})
it('requires persistent import intent and observed proof before staged receipt', () => {
  const r = initial(),
    p = { ...r, pending: { action: 'import' as const, beforeProofRef: r.currentProofRef } }
  expect(validPackageTransition(r, p)).toBe(true)
  const staged = {
    ...r,
    state: 'staged' as const,
    replacementSlideId: 'actual',
    currentProofRef: ref('receipt-1'),
    receipts: [{ action: 'import' as const, proofRef: ref('receipt-1'), slideId: 'actual' }],
  }
  expect(validPackageTransition(p, staged)).toBe(false)
  expect(
    validPackageTransition(
      {
        ...p,
        pending: { ...p.pending, insertedSlideId: 'actual', afterProofRef: ref('receipt-1') },
      },
      staged,
    ),
  ).toBe(true)
})
it('allows explicit no-write closure of baseline import pending without enabling import replay', () => {
  const r = initial(),
    p = { ...r, pending: { action: 'import' as const, beforeProofRef: r.currentProofRef } }
  expect(validPackageTransition(p, { ...r, state: 'discarded' })).toBe(true)
  expect(validPackageTransition(p, r)).toBe(false)
  expect(validPackageTransition(r, { ...r, state: 'discarded' })).toBe(true)
})
