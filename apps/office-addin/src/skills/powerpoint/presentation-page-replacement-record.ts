import { validSourceSlideId } from './presentation-page-delivery.js'

export interface PresentationPageReplacement {
  version: 1
  changeId: string
  documentId: string
  projectId: string
  parentRequestId: string
  requestId: string
  pageId: string
  backupId: string
  parentArtifactDigest: string
  backupDigest: string
  originalPackageDigest: string
  replacementPackageDigest: string
  sourceSlideId: string
  oldSlideId: string
  beforeSlideIds: string[]
  state: 'pending' | 'inserted' | 'staged' | 'discard_pending' | 'discarded'
  newSlideId?: string
}
const id = (value: unknown, max: number) =>
  typeof value === 'string' && new RegExp(`^[A-Za-z0-9_-]{1,${max}}$`).test(value)
const hostId = (value: unknown): value is string =>
  typeof value === 'string' &&
  value.length > 0 &&
  value.length <= 256 &&
  !Array.from(value).some(
    (char) => char.charCodeAt(0) < 32 || (char.charCodeAt(0) >= 127 && char.charCodeAt(0) <= 159),
  )
const digest = (value: unknown) => typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
export function validatePresentationPageReplacement(
  value: unknown,
): value is PresentationPageReplacement {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const r = value as PresentationPageReplacement
  return (
    Object.keys(r).every((key) =>
      [
        'version',
        'changeId',
        'documentId',
        'projectId',
        'parentRequestId',
        'requestId',
        'pageId',
        'backupId',
        'parentArtifactDigest',
        'backupDigest',
        'originalPackageDigest',
        'replacementPackageDigest',
        'sourceSlideId',
        'oldSlideId',
        'beforeSlideIds',
        'state',
        'newSlideId',
      ].includes(key),
    ) &&
    r.version === 1 &&
    id(r.changeId, 128) &&
    id(r.projectId, 80) &&
    id(r.parentRequestId, 128) &&
    id(r.requestId, 128) &&
    r.parentRequestId !== r.requestId &&
    id(r.pageId, 80) &&
    id(r.backupId, 128) &&
    typeof r.documentId === 'string' &&
    r.documentId.length > 0 &&
    r.documentId.length <= 4096 &&
    digest(r.parentArtifactDigest) &&
    digest(r.backupDigest) &&
    digest(r.originalPackageDigest) &&
    digest(r.replacementPackageDigest) &&
    validSourceSlideId(r.sourceSlideId) &&
    hostId(r.oldSlideId) &&
    Array.isArray(r.beforeSlideIds) &&
    r.beforeSlideIds.length > 0 &&
    r.beforeSlideIds.length <= 512 &&
    Array.from(r.beforeSlideIds).every(hostId) &&
    new Set(r.beforeSlideIds).size === r.beforeSlideIds.length &&
    r.beforeSlideIds.includes(r.oldSlideId) &&
    ['pending', 'inserted', 'staged', 'discard_pending', 'discarded'].includes(r.state) &&
    (r.state === 'pending'
      ? r.newSlideId === undefined
      : hostId(r.newSlideId) && !r.beforeSlideIds.includes(r.newSlideId)) &&
    // Reserve the largest UTF-8 host ID before insertion so recording it cannot exceed the limit.
    new TextEncoder().encode(
      JSON.stringify({ ...r, newSlideId: '界'.repeat(256), state: 'discard_pending' }),
    ).byteLength <=
      192 * 1024
  )
}
