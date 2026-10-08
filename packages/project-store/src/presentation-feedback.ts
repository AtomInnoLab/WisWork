/** User-reported editing needs, never QA, source verification or host acceptance. */
export type PresentationProductionFeedbackStatus =
  'needs_correction' | 'no_correction' | 'not_evaluated'
export interface PresentationProductionFeedbackPage {
  pageId: string
  status: PresentationProductionFeedbackStatus
  note?: string
}
export interface PresentationProductionFeedbackSnapshot {
  revision: number
  recordedAt: string
  pages: PresentationProductionFeedbackPage[]
}
export interface PresentationProductionFeedbackLedger {
  version: 1
  source: 'user_reported'
  projectId: string
  documentId: string
  requestId: string
  inputDigest: string
  planDigest: string
  planRevision: number
  pageIds: string[]
  revision: number
  snapshots: PresentationProductionFeedbackSnapshot[]
}
export const MAX_PRESENTATION_PRODUCTION_FEEDBACK_BYTES = 5 * 1024 * 1024
export const MAX_PRESENTATION_PRODUCTION_FEEDBACK_REVISIONS = 64
function reject(error: string): never {
  throw new Error(error)
}
function object(value: unknown, keys: string[], error: string): Record<string, unknown> {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.getPrototypeOf(value) !== Object.prototype
  )
    reject(error)
  const record = value as Record<string, unknown>
  if (
    Object.keys(record).some((k) => !keys.includes(k)) ||
    Object.values(Object.getOwnPropertyDescriptors(record)).some((d) => !('value' in d))
  )
    reject(error)
  return record
}
const id = (value: unknown): value is string =>
  typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
const positive = (value: unknown) => Number.isSafeInteger(value) && Number(value) > 0
function pages(value: unknown, error: string): PresentationProductionFeedbackPage[] {
  if (!Array.isArray(value) || !value.length || value.length > 32) reject(error)
  const seen = new Set<string>()
  for (const item of value) {
    const page = object(item, ['pageId', 'status', 'note'], error)
    if (
      !id(page.pageId) ||
      seen.has(page.pageId) ||
      !['needs_correction', 'no_correction', 'not_evaluated'].includes(page.status as string) ||
      (Object.hasOwn(page, 'note') &&
        (typeof page.note !== 'string' || new TextEncoder().encode(page.note).byteLength > 2000))
    )
      reject(error)
    seen.add(page.pageId)
  }
  return structuredClone(value) as PresentationProductionFeedbackPage[]
}
export function parsePresentationProductionFeedbackPages(
  value: unknown,
): PresentationProductionFeedbackPage[] {
  return pages(value, 'invalid_request')
}
export function parsePresentationProductionFeedbackLedger(
  value: unknown,
): PresentationProductionFeedbackLedger {
  const error = 'invalid_state',
    record = object(
      value,
      [
        'version',
        'source',
        'projectId',
        'documentId',
        'requestId',
        'inputDigest',
        'planDigest',
        'planRevision',
        'pageIds',
        'revision',
        'snapshots',
      ],
      error,
    )
  if (
    record.version !== 1 ||
    record.source !== 'user_reported' ||
    !id(record.projectId) ||
    !id(record.requestId) ||
    typeof record.documentId !== 'string' ||
    !record.documentId.trim() ||
    record.documentId.length > 2048 ||
    !positive(record.planRevision) ||
    !positive(record.revision) ||
    Number(record.revision) > MAX_PRESENTATION_PRODUCTION_FEEDBACK_REVISIONS ||
    [record.inputDigest, record.planDigest].some(
      (v) => typeof v !== 'string' || !/^[a-f0-9]{64}$/.test(v),
    ) ||
    !Array.isArray(record.pageIds) ||
    !record.pageIds.length ||
    record.pageIds.length > 32 ||
    record.pageIds.some((v) => !id(v)) ||
    new Set(record.pageIds).size !== record.pageIds.length ||
    !Array.isArray(record.snapshots) ||
    record.snapshots.length !== record.revision
  )
    reject(error)
  let previousTime = ''
  for (const [index, item] of (record.snapshots as unknown[]).entries()) {
    const snapshot = object(item, ['revision', 'recordedAt', 'pages'], error)
    if (
      snapshot.revision !== index + 1 ||
      typeof snapshot.recordedAt !== 'string' ||
      !/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(snapshot.recordedAt) ||
      !Number.isFinite(Date.parse(snapshot.recordedAt)) ||
      new Date(snapshot.recordedAt).toISOString() !== snapshot.recordedAt ||
      snapshot.recordedAt < previousTime
    )
      reject(error)
    previousTime = snapshot.recordedAt
    const parsed = pages(snapshot.pages, error)
    if (
      parsed.length !== (record.pageIds as string[]).length ||
      parsed.some((page, i) => page.pageId !== (record.pageIds as string[])[i])
    )
      reject(error)
  }
  if (
    new TextEncoder().encode(JSON.stringify(value)).byteLength >
    MAX_PRESENTATION_PRODUCTION_FEEDBACK_BYTES
  )
    reject(error)
  return structuredClone(value) as PresentationProductionFeedbackLedger
}
