import { describe, expect, it } from 'vitest'
import {
  validatePresentationMasterXmlChange as valid,
  validMasterXmlTransition as transition,
  masterXmlReservedBytes,
  type PresentationMasterXmlChange,
} from '../src/skills/powerpoint/presentation-master-xml-change.js'
const ref = (key: string) => ({ key, sha256: 'a'.repeat(64), sizeBytes: 100 })
const seed = (): PresentationMasterXmlChange => ({
  version: 1,
  kind: 'master_xml',
  changeId: 'change',
  documentId: 'doc',
  intent: 'edit',
  sourceSlideId: 's1',
  packageSourceSlideId: '256#',
  snapshotRef: ref('snapshot'),
  preparedRef: ref('page-0'),
  currentProofRef: ref('receipt-0'),
  state: 'prepared',
  cursor: { phase: 'original_probe_stage', index: 0, substep: 0 },
  receiptCount: 0,
  scope: {
    originalMasterId: 'm1',
    affectedPageCount: 600,
    originalLayoutCount: 128,
    affectedMasterCount: 2,
  },
  introducedMasterCount: 0,
  inventoryCleanupVerified: false,
  reviews: [],
})
describe('master XML compact durable journal', () => {
  it('accepts complete scope without an artificial page count limit', () => {
    const r = seed()
    r.scope.affectedPageCount = 100_001
    expect(valid(r)).toBe(true)
    expect(transition(undefined, r)).toBe(true)
    expect(
      new TextEncoder().encode(JSON.stringify(r)).length + masterXmlReservedBytes(r),
    ).toBeLessThan(192 * 1024)
  })
  it('requires native and package source identities to stay distinct', () => {
    for (const id of ['', ' ', 'x\u0001', 'x'.repeat(257), '256#']) {
      const r = seed()
      r.sourceSlideId = id
      expect(valid(r)).toBe(id === '256#')
    }
    const r = seed()
    r.packageSourceSlideId = 's1'
    expect(valid(r)).toBe(false)
  })
  it('requires a pending intent before host receipts and immutable backup identities', () => {
    const r = seed(),
      p = {
        ...r,
        pending: {
          action: 'original_probe_stage' as const,
          index: 0,
          beforeProofRef: r.currentProofRef,
        },
      }
    expect(transition(r, p)).toBe(true)
    expect(transition(r, { ...r, receiptCount: 1, currentProofRef: ref('receipt-1') })).toBe(false)
    expect(transition(r, { ...p, snapshotRef: ref('page-1') })).toBe(false)
  })
  it('requires inserted identity and durable after proof before a stage can close', () => {
    const r = seed(),
      p = {
        ...r,
        pending: {
          action: 'original_probe_stage' as const,
          index: 0,
          beforeProofRef: r.currentProofRef,
        },
      }
    const actual = { slideId: 'probe', masterId: 'new-master', layoutId: 'new-layout' }
    expect(valid({ ...p, pending: { ...p.pending, afterProofRef: ref('receipt-1') } })).toBe(false)
    const ack = {
      ...p,
      pending: { ...p.pending, inserted: actual, afterProofRef: ref('receipt-1') },
    }
    expect(transition(p, ack)).toBe(true)
    expect(
      transition(ack, {
        ...r,
        state: 'probing_original',
        cursor: { phase: 'original_probe', index: 0, substep: 0 },
        receiptCount: 1,
        currentProofRef: ref('receipt-1'),
        originalProbeSource: actual,
      }),
    ).toBe(true)
  })
  it('rejects a receipt jumping from a probe straight to applied', () => {
    const r = seed(),
      actual = { slideId: 'probe', masterId: 'm2', layoutId: 'l2' },
      p = {
        ...r,
        pending: {
          action: 'original_probe_stage' as const,
          index: 0,
          beforeProofRef: r.currentProofRef,
          inserted: actual,
          afterProofRef: ref('receipt-1'),
        },
      }
    expect(
      transition(p, {
        ...r,
        state: 'applied',
        stagedSource: { slideId: 'fake', masterId: 'm3', layoutId: 'l3' },
        originalProbeSource: actual,
        receiptCount: 1,
        currentProofRef: ref('receipt-1'),
      }),
    ).toBe(false)
  })
  it('rejects unrelated terminal transitions even with a valid durable stage receipt', () => {
    const r = seed(),
      actual = { slideId: 'probe', masterId: 'm2', layoutId: 'l2' },
      p = {
        ...r,
        pending: {
          action: 'original_probe_stage' as const,
          index: 0,
          beforeProofRef: r.currentProofRef,
          inserted: actual,
          afterProofRef: ref('receipt-1'),
        },
      }
    expect(
      transition(p, {
        ...r,
        state: 'discarded',
        originalProbeSource: actual,
        receiptCount: 1,
        currentProofRef: ref('receipt-1'),
      }),
    ).toBe(false)
  })
  it('permits explicit unchanged initial intent closure without replay', () => {
    const r = seed(),
      p = {
        ...r,
        pending: {
          action: 'original_probe_stage' as const,
          index: 0,
          beforeProofRef: r.currentProofRef,
        },
      }
    expect(transition(p, { ...r, state: 'discarded' })).toBe(true)
  })
  it('rejects complete inventory cleanup claims and invalid cursor identities', () => {
    expect(valid({ ...seed(), inventoryCleanupVerified: true })).toBe(false)
    expect(
      valid({
        ...seed(),
        pending: { action: 'stage', index: 1, beforeProofRef: ref('receipt-0') },
      }),
    ).toBe(false)
  })
})

it('requires a precise native target for an individual layout write', () => {
  const r = seed(),
    actual = { slideId: 'staged', masterId: 'm2', layoutId: 'l2' }
  const applying = {
    ...r,
    state: 'applying' as const,
    stagedSource: actual,
    cursor: { phase: 'forward_page' as const, index: 2, substep: 0 },
  }
  expect(
    valid({
      ...applying,
      pending: { action: 'forward_page', index: 2, beforeProofRef: r.currentProofRef },
    }),
  ).toBe(false)
  expect(
    valid({
      ...applying,
      pending: {
        action: 'forward_page',
        index: 2,
        beforeProofRef: r.currentProofRef,
        targetSlideId: 's2',
        targetMasterId: 'm2',
        targetLayoutId: 'l2',
      },
    }),
  ).toBe(true)
})
it('refuses rewriting or dropping unrelated historical review entries', () => {
  const r = {
    ...seed(),
    state: 'applied' as const,
    stagedSource: { slideId: 'new', masterId: 'm2', layoutId: 'l2' },
    cursor: { phase: 'verify' as const, index: 0, substep: 0 },
    reviews: [
      { slideId: 's1', reviewRef: ref('image-0') },
      { slideId: 's2', reviewRef: ref('image-1') },
    ],
  }
  expect(valid(r)).toBe(true)
  expect(transition(r, { ...r, reviews: [{ slideId: 's1', reviewRef: ref('image-2') }] })).toBe(
    false,
  )
  expect(
    transition(r, { ...r, reviews: r.reviews.map((x) => ({ ...x, reviewRef: ref('image-3') })) }),
  ).toBe(false)
  expect(
    transition(r, {
      ...r,
      reviews: [...r.reviews, { slideId: 's3', reviewRef: ref('image-4') }],
      reviewSequence: 1,
    }),
  ).toBe(true)
})
it('accepts the last compact journal byte budget and refuses the next review row', () => {
  const r = {
    ...seed(),
    state: 'applied' as const,
    stagedSource: { slideId: 'new', masterId: 'm2', layoutId: 'l2' },
    cursor: { phase: 'verify' as const, index: 0, substep: 0 },
    reviews: [] as { slideId: string; reviewRef: ReturnType<typeof ref> }[],
  }
  while (true) {
    const next = {
      ...r,
      reviews: [
        ...r.reviews,
        {
          slideId: `page-${r.reviews.length}-${'界'.repeat(80)}`,
          reviewRef: ref(`image-${r.reviews.length}`),
        },
      ],
    }
    if (
      new TextEncoder().encode(JSON.stringify(next)).length + masterXmlReservedBytes(next) >
      192 * 1024
    ) {
      expect(valid(r)).toBe(true)
      expect(valid(next)).toBe(false)
      expect(r.reviews.length).toBeGreaterThan(100)
      break
    }
    r.reviews = next.reviews
  }
})
