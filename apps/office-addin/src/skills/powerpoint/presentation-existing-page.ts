import { validSourceSlideId } from './presentation-page-delivery.js'
import { validExistingVisualCapture, validExistingVisualReview, type ExistingVisualCapture, type ExistingVisualReview } from './presentation-existing-visual-review.js'

export interface PresentationExistingPageChange {
  version: 1
  changeId: string
  documentId: string
  baselineId: string
  baselineDigest: string
  scope: { slideIds: string[] }
  oldSlideId: string
  beforeSlideIds: string[]
  originalPackageDigest: string
  replacementPackageDigest: string
  sourceSlideId: string
  backup: { backupId: string; sha256: string; sizeBytes: number }
  state:
    | 'pending'
    | 'inserted'
    | 'staged'
    | 'discard_pending'
    | 'discarded'
    | 'commit_pending'
    | 'applied'
    | 'undo_pending'
    | 'restore_inserted'
    | 'undone'
  newSlideId?: string
  restoredSlideId?: string
  captures?: ExistingVisualCapture[]
  reviews?: ExistingVisualReview[]
  backupReleasedAt?: string
}

const id = (value: unknown) => typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
const digest = (value: unknown) => typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
const hostId = (value: unknown): value is string =>
  typeof value === 'string' && value.length > 0 && value.length <= 256 &&
  !Array.from(value).some((c) => c.charCodeAt(0) < 32 || (c.charCodeAt(0) >= 127 && c.charCodeAt(0) <= 159))
const ids = (value: unknown): value is string[] =>
  Array.isArray(value) && value.length > 0 && value.length <= 512 &&
  value.every(hostId) && new Set(value).size === value.length
const bytes = (value: unknown) => new TextEncoder().encode(JSON.stringify(value)).byteLength

export const existingPageReservedBytes = (r: PresentationExistingPageChange) =>
  bytes({
    ...r,
    state: 'restore_inserted',
    newSlideId: '\uffff'.repeat(256),
    restoredSlideId: '\uffff'.repeat(256),
    backupReleasedAt: '2026-09-24T00:00:00.000Z',
  }) - bytes(r) + Math.max(0, 2 * 8202 + 12 - (r.reviews ? bytes(r.reviews) + ',"reviews":'.length : 0)) +
  Math.max(0, 2 * 524 + 13 - (r.captures ? bytes(r.captures) + ',"captures":'.length : 0))

export function validatePresentationExistingPageChange(value: unknown): value is PresentationExistingPageChange {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const r = value as PresentationExistingPageChange
  const keys = [
    'version', 'changeId', 'documentId', 'baselineId', 'baselineDigest', 'scope',
    'oldSlideId', 'beforeSlideIds', 'originalPackageDigest', 'replacementPackageDigest',
    'sourceSlideId', 'backup', 'state', 'newSlideId', 'restoredSlideId', 'captures', 'reviews', 'backupReleasedAt',
  ]
  if (
    Object.keys(r).some((key) => !keys.includes(key)) ||
    r.version !== 1 || !id(r.changeId) || !id(r.baselineId) ||
    typeof r.documentId !== 'string' || !r.documentId || r.documentId.length > 4096 ||
    !digest(r.baselineDigest) || !digest(r.originalPackageDigest) ||
    !digest(r.replacementPackageDigest) || !validSourceSlideId(r.sourceSlideId) ||
    !hostId(r.oldSlideId) || !ids(r.beforeSlideIds) || !r.beforeSlideIds.includes(r.oldSlideId) ||
    !r.scope || typeof r.scope !== 'object' || Array.isArray(r.scope) ||
    Object.keys(r.scope).length !== 1 || !ids(r.scope.slideIds) ||
    !r.scope.slideIds.includes(r.oldSlideId) ||
    !r.scope.slideIds.every((slideId) => r.beforeSlideIds.includes(slideId)) ||
    !r.backup || typeof r.backup !== 'object' || Array.isArray(r.backup) ||
    Object.keys(r.backup).length !== 3 || !id(r.backup.backupId) ||
    !digest(r.backup.sha256) || !Number.isSafeInteger(r.backup.sizeBytes) ||
    r.backup.sizeBytes < 1 || r.backup.sizeBytes > 100 * 1024 * 1024 ||
    !['pending', 'inserted', 'staged', 'discard_pending', 'discarded', 'commit_pending', 'applied', 'undo_pending', 'restore_inserted', 'undone'].includes(r.state)
  ) return false
  if (r.backupReleasedAt !== undefined &&
    (!['discarded', 'undone'].includes(r.state) || typeof r.backupReleasedAt !== 'string' ||
      !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(r.backupReleasedAt) ||
      !Number.isFinite(Date.parse(r.backupReleasedAt)) ||
      new Date(r.backupReleasedAt).toISOString() !== r.backupReleasedAt)) return false
  if (r.state === 'pending') {
    if (r.newSlideId !== undefined) return false
  } else if (!hostId(r.newSlideId) || r.beforeSlideIds.includes(r.newSlideId)) return false
  if (['restore_inserted', 'undone'].includes(r.state)) {
    if (!hostId(r.restoredSlideId) || r.restoredSlideId === r.newSlideId || r.beforeSlideIds.includes(r.restoredSlideId)) return false
  } else if (r.restoredSlideId !== undefined) return false
  const targets = r.state === 'staged' ? [r.oldSlideId, r.newSlideId] :
      r.state === 'applied' ? [r.newSlideId] :
      r.state === 'discarded' ? [r.oldSlideId] :
      r.state === 'undone' ? [r.restoredSlideId] : []
  if (r.captures !== undefined &&
    (!Array.isArray(r.captures) || r.captures.length > targets.length ||
      !r.captures.every(validExistingVisualCapture) ||
      r.captures.some((capture) => !targets.includes(capture.hostSlideId)) ||
      new Set(r.captures.map((capture) => capture.hostSlideId)).size !== r.captures.length)) return false
  if (r.reviews !== undefined) {
    if (!Array.isArray(r.reviews) || r.reviews.length > targets.length ||
      !r.reviews.every(validExistingVisualReview) ||
      r.reviews.some((review) => !targets.includes(review.hostSlideId)) ||
      new Set(r.reviews.map((review) => review.hostSlideId)).size !== r.reviews.length ||
      r.reviews.some((review) => { const capture = r.captures?.find((c) => c.hostSlideId === review.hostSlideId); return capture !== undefined && (capture.screenshotDigest !== review.screenshotDigest || capture.capturedAt !== review.capturedAt) })) return false
  }
  return bytes(r) + existingPageReservedBytes(r) <= 192 * 1024
}

export function validExistingPageTransition(
  before: PresentationExistingPageChange | undefined,
  after: PresentationExistingPageChange,
): boolean {
  if (!validatePresentationExistingPageChange(after)) return false
  if (!before) return after.state === 'pending'
  if (!validatePresentationExistingPageChange(before)) return false
  const identity = (r: PresentationExistingPageChange) => JSON.stringify({
    ...r, state: undefined, newSlideId: undefined, restoredSlideId: undefined, captures: undefined, reviews: undefined, backupReleasedAt: undefined,
  })
  const next: Record<PresentationExistingPageChange['state'], PresentationExistingPageChange['state'][]> = {
    pending: ['inserted'], inserted: ['staged'], staged: ['discard_pending', 'commit_pending'],
    discard_pending: ['discarded'], discarded: [],
    commit_pending: ['applied'], applied: ['undo_pending'], undo_pending: ['restore_inserted'],
    restore_inserted: ['undone'], undone: [],
  }
  const terminal = before.state === after.state && ['staged', 'applied', 'discarded', 'undone'].includes(after.state)
  const captureOnly = terminal && JSON.stringify(before.reviews) === JSON.stringify(after.reviews) &&
    (after.captures?.length ?? 0) >= (before.captures?.length ?? 0) &&
    (after.captures?.length ?? 0) <= (before.captures?.length ?? 0) + 2 &&
    after.captures?.some((capture) => !before.captures?.some((prior) => JSON.stringify(prior) === JSON.stringify(capture))) === true &&
    (before.captures ?? []).every((capture) => after.captures?.some((nextCapture) => nextCapture.hostSlideId === capture.hostSlideId))
  const reviewOnly = terminal && JSON.stringify(before.captures) === JSON.stringify(after.captures) &&
    (after.reviews?.length ?? 0) === (before.reviews?.length ?? 0) + 1 &&
    (before.reviews ?? []).every((review) => after.reviews?.some((nextReview) => JSON.stringify(nextReview) === JSON.stringify(review)))
  const released = before.state === after.state && ['discarded', 'undone'].includes(after.state) &&
    before.backupReleasedAt === undefined && after.backupReleasedAt !== undefined &&
    before.newSlideId === after.newSlideId && before.restoredSlideId === after.restoredSlideId &&
    JSON.stringify(before.captures) === JSON.stringify(after.captures) &&
    JSON.stringify(before.reviews) === JSON.stringify(after.reviews)
  return identity(before) === identity(after) && (released ||
    (before.backupReleasedAt === after.backupReleasedAt && (captureOnly || reviewOnly ||
    (before.backupReleasedAt === undefined && next[before.state].includes(after.state) &&
      after.captures === undefined && after.reviews === undefined)))) &&
    (before.newSlideId === undefined || before.newSlideId === after.newSlideId) &&
    (before.restoredSlideId === undefined || before.restoredSlideId === after.restoredSlideId)
}
