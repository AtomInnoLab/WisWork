/** Browser-safe canonical serialization shared with persisted presentation hashes. */
export function canonicalPresentationValue(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(canonicalPresentationValue).join(',')}]`
  if (value && typeof value === 'object')
    return `{${Object.keys(value)
      .sort()
      .map(
        (key) =>
          `${JSON.stringify(key)}:${canonicalPresentationValue((value as Record<string, unknown>)[key])}`,
      )
      .join(',')}}`
  return JSON.stringify(value)
}

export function presentationPlanSnapshotInputs(plan: unknown) {
  const value = plan as Record<string, unknown>
  return {
    sourcesDigest: canonicalPresentationValue(value.sources),
    claimsDigest: canonicalPresentationValue(value.claims),
    slidesDigest: canonicalPresentationValue(value.slides),
    styleDigest: canonicalPresentationValue(
      value.brandKit === undefined ? value.style : { style: value.style, brandKit: value.brandKit },
    ),
  }
}
