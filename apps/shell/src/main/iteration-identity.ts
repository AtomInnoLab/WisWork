/** Only packaged metadata selects a test identity; malformed metadata fails closed. */
export function resolveIterationIdentity(value: unknown): { productName: string } | null {
  if (value === undefined) return null
  if (!value || typeof value !== 'object') throw new Error('Invalid iteration metadata')
  const meta = value as Record<string, unknown>
  if (
    typeof meta.commit !== 'string' ||
    !/^[a-f0-9]{7,40}$/.test(meta.commit) ||
    typeof meta.builtAt !== 'string' ||
    !Number.isFinite(Date.parse(meta.builtAt))
  )
    throw new Error('Invalid iteration metadata')
  if (meta.mode === 'dogfood') return { productName: 'WisWork Dogfood' }
  if (meta.mode === 'preview' && Number.isSafeInteger(meta.pr) && (meta.pr as number) > 0) {
    return { productName: `WisWork Preview PR${meta.pr}` }
  }
  throw new Error('Invalid iteration metadata')
}
