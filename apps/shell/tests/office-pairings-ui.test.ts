import { describe, expect, it } from 'vitest'

import { mergeOfficePairings } from '../src/renderer/src/office-pairings'

const pairing = (pairingId: string, hostLabel: 'Word' | 'Excel' = 'Word') => ({
  pairingId,
  hostLabel,
  origin: 'https://office.example.test',
})

describe('Office pairing renderer recovery', () => {
  it('does not lose a live event when an older pending snapshot resolves later', () => {
    const snapshotCapturedEarlier = [pairing('old_pairing')]
    let current = [pairing('new_event', 'Excel')]

    current = mergeOfficePairings(current, snapshotCapturedEarlier)

    expect(current.map((entry) => entry.pairingId)).toEqual(['new_event', 'old_pairing'])
  })

  it('deduplicates a pairing delivered by both event and snapshot', () => {
    const current = mergeOfficePairings([pairing('same_pairing')], [pairing('same_pairing')])
    expect(current).toHaveLength(1)
  })
})
