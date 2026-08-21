import type { OfficePairingRequest } from '../../shared/home-api'

export function mergeOfficePairings(
  current: readonly OfficePairingRequest[],
  incoming: readonly OfficePairingRequest[],
): OfficePairingRequest[] {
  const merged = [...current]
  const known = new Set(current.map((pairing) => pairing.pairingId))
  for (const pairing of incoming) {
    if (known.has(pairing.pairingId)) continue
    known.add(pairing.pairingId)
    merged.push(pairing)
  }
  return merged
}
