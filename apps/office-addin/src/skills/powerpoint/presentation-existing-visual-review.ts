export interface ExistingVisualReview {
  hostSlideId: string
  screenshotDigest: string
  capturedAt: string
  reviewedAt: string
  status: 'pass' | 'fail'
  notes: string
}

const bytes = (value: unknown) => new TextEncoder().encode(JSON.stringify(value)).byteLength
const timestamp = (value: unknown): value is string =>
  typeof value === 'string' && value.length <= 40 && Number.isFinite(Date.parse(value)) &&
  new Date(value).toISOString() === value

export function validExistingVisualReview(value: unknown): value is ExistingVisualReview {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const v = value as ExistingVisualReview
  return Object.keys(v).length === 6 &&
    typeof v.hostSlideId === 'string' && v.hostSlideId.length > 0 && v.hostSlideId.length <= 256 &&
    !Array.from(v.hostSlideId).some((c) => c.charCodeAt(0) < 32) &&
    typeof v.screenshotDigest === 'string' && /^[a-f0-9]{64}$/.test(v.screenshotDigest) &&
    timestamp(v.capturedAt) && timestamp(v.reviewedAt) && v.reviewedAt >= v.capturedAt &&
    (v.status === 'pass' || v.status === 'fail') &&
    typeof v.notes === 'string' && bytes(v) <= 8192
}
