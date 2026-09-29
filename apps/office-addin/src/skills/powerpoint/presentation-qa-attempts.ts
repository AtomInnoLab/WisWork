/** Durable bookkeeping only; a recorded screenshot attempt is not host or visual acceptance. */
export interface PresentationQaAttempt {
  version: 1
  id: string
  source?: 'production'
  documentId: string
  projectId: string
  requestId: string
  artifactDigest: string
  pageId: string
  hostSlideId: string
  startedAt: string
  status: 'started' | 'recorded' | 'waiting' | 'failed' | 'cancelled'
  finishedAt?: string
  errorCode?:
    | 'screenshot_unavailable'
    | 'inspection_failed'
    | 'cancelled'
    | 'publication_failed'
    | 'state_changed'
}
// Reserve the longest terminal status and fields before accepting a started record.
export const PRESENTATION_QA_ATTEMPT_TERMINAL_RESERVE_BYTES =
  2 +
  ',"finishedAt":"0000-00-00T00:00:00.000Z"'.length +
  ',"errorCode":"screenshot_unavailable"'.length
const iso = (value: unknown): value is string =>
  typeof value === 'string' &&
  /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(value) &&
  Number.isFinite(Date.parse(value)) &&
  new Date(value).toISOString() === value
const text = (value: unknown, max: number): value is string =>
  typeof value === 'string' && value.length > 0 && value.length <= max
const id = (value: unknown, max: number): value is string =>
  text(value, max) && /^[A-Za-z0-9_-]+$/.test(value)
export function parsePresentationQaAttempt(value: unknown): PresentationQaAttempt {
  const invalid = (): never => {
    throw Error('presentation_qa_attempt_state_invalid')
  }
  if (!value || typeof value !== 'object' || Array.isArray(value)) invalid()
  const v = value as Record<string, unknown>
  if (
    Object.keys(v).some(
      (key) =>
        ![
          'version',
          'id',
          'source',
          'documentId',
          'projectId',
          'requestId',
          'artifactDigest',
          'pageId',
          'hostSlideId',
          'startedAt',
          'status',
          'finishedAt',
          'errorCode',
        ].includes(key),
    ) ||
    v.version !== 1 ||
    typeof v.id !== 'string' ||
    !/^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/.test(v.id) ||
    (Object.hasOwn(v, 'source') && v.source !== 'production') ||
    !text(v.documentId, 4096) ||
    !id(v.projectId, 80) ||
    !text(v.requestId, 128) ||
    !id(v.pageId, 80) ||
    !text(v.hostSlideId, 256) ||
    typeof v.artifactDigest !== 'string' ||
    !/^[a-f0-9]{64}$/.test(v.artifactDigest) ||
    !iso(v.startedAt) ||
    !['started', 'recorded', 'waiting', 'failed', 'cancelled'].includes(String(v.status))
  )
    invalid()
  if (v.status === 'started') {
    if (Object.hasOwn(v, 'finishedAt') || Object.hasOwn(v, 'errorCode')) invalid()
  } else {
    if (!iso(v.finishedAt) || v.finishedAt < (v.startedAt as string)) invalid()
    if (v.status === 'recorded') {
      if (Object.hasOwn(v, 'errorCode')) invalid()
    } else if (
      v.status === 'waiting'
        ? v.errorCode !== 'screenshot_unavailable'
        : v.status === 'cancelled'
          ? v.errorCode !== 'cancelled'
          : !['inspection_failed', 'publication_failed', 'state_changed'].includes(
              String(v.errorCode),
            )
    )
      invalid()
  }
  return structuredClone(value) as PresentationQaAttempt
}
export function validatePresentationQaAttempt(value: unknown): value is PresentationQaAttempt {
  try {
    parsePresentationQaAttempt(value)
    return true
  } catch {
    return false
  }
}
