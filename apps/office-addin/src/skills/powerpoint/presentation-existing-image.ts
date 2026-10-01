import type { PictureSnapshot } from './browser-presentation-image-adapter.js'
import {
  validPresentationImageBackupMetadata,
  type PresentationImageBackupMetadata,
} from './presentation-image-backup.js'
import {
  validExistingVisualCapture,
  validExistingVisualReview,
  type ExistingVisualCapture,
  type ExistingVisualReview,
} from './presentation-existing-visual-review.js'

export interface PresentationExistingImageChange {
  version: 1
  changeId: string
  documentId: string
  baselineId: string
  baselineDigest: string
  scope: { slideIds: string[]; shapeIds?: string[] }
  hostSlideId: string
  oldShapeId: string
  assetDigest: string
  original: PictureSnapshot
  sourceBackup?: PresentationImageBackupMetadata
  reapplies?: string
  backup: PresentationImageBackupMetadata
  state: 'pending' | 'complete' | 'undo_pending' | 'undone'
  insertedShapeId?: string
  after?: PictureSnapshot
  undoBaseline?: PictureSnapshot
  restoredShapeId?: string
  capture?: ExistingVisualCapture
  review?: ExistingVisualReview
}

const bytes = (value: unknown) => new TextEncoder().encode(JSON.stringify(value)).byteLength
const id = (value: unknown) => typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
const digest = (value: unknown) => typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
const hostId = (value: unknown): value is string =>
  typeof value === 'string' &&
  value.length > 0 &&
  value.length <= 256 &&
  !Array.from(value).some(
    (char) => char.charCodeAt(0) < 32 || (char.charCodeAt(0) >= 127 && char.charCodeAt(0) <= 159),
  )
const ids = (value: unknown): value is string[] =>
  Array.isArray(value) &&
  value.length > 0 &&
  value.length <= 500 &&
  value.every(hostId) &&
  new Set(value).size === value.length

function snapshot(value: unknown): value is PictureSnapshot {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const s = value as PictureSnapshot
  const keys = [
    'slideId',
    'shapeId',
    'geometry',
    'rotation',
    'name',
    'altTextTitle',
    'altTextDescription',
    'zOrderPosition',
    'shapeIds',
    'pictureFingerprint',
    'mediaDigest',
  ]
  const g = s.geometry
  return (
    Object.keys(s).length === keys.length &&
    Object.keys(s).every((key) => keys.includes(key)) &&
    hostId(s.slideId) &&
    hostId(s.shapeId) &&
    !!g &&
    typeof g === 'object' &&
    !Array.isArray(g) &&
    Object.keys(g).length === 4 &&
    (['left', 'top', 'width', 'height'] as const).every(
      (key) => typeof g[key] === 'number' && Number.isFinite(g[key]) && Math.abs(g[key]) <= 100000,
    ) &&
    g.width >= 0 &&
    g.height >= 0 &&
    typeof s.rotation === 'number' &&
    Number.isFinite(s.rotation) &&
    Math.abs(s.rotation) <= 360 &&
    [s.name, s.altTextTitle, s.altTextDescription].every(
      (text) => typeof text === 'string' && text.length <= 12000,
    ) &&
    Number.isSafeInteger(s.zOrderPosition) &&
    s.zOrderPosition >= 0 &&
    Array.isArray(s.shapeIds) &&
    s.shapeIds.length > 0 &&
    s.shapeIds.length <= 1_000 &&
    s.shapeIds.every(hostId) &&
    new Set(s.shapeIds).size === s.shapeIds.length &&
    s.shapeIds[s.zOrderPosition] === s.shapeId &&
    digest(s.pictureFingerprint) &&
    digest(s.mediaDigest)
  )
}

// Reserve both future snapshots, IDs, and the longest state before any native write.
export const existingImageReservedBytes = (r: PresentationExistingImageChange) =>
  (r.after ? 0 : bytes(r.original) + 3300) +
  (r.undoBaseline ? 0 : bytes(r.original) + 3300) +
  (r.insertedShapeId ? 0 : 1700) +
  (r.restoredShapeId ? 0 : 1700) +
  (r.capture ? 0 : 524) +
  (r.review ? 0 : 8202) +
  32

export function validatePresentationExistingImageChange(
  value: unknown,
): value is PresentationExistingImageChange {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const r = value as PresentationExistingImageChange
  const allowed = [
    'version',
    'changeId',
    'documentId',
    'baselineId',
    'baselineDigest',
    'scope',
    'hostSlideId',
    'oldShapeId',
    'assetDigest',
    'original',
    'backup',
    'sourceBackup',
    'reapplies',
    'state',
    'insertedShapeId',
    'after',
    'undoBaseline',
    'restoredShapeId',
    'capture',
    'review',
  ]
  if (
    Object.keys(r).some((key) => !allowed.includes(key)) ||
    r.version !== 1 ||
    !id(r.changeId) ||
    !id(r.baselineId) ||
    !digest(r.baselineDigest) ||
    !digest(r.assetDigest) ||
    typeof r.documentId !== 'string' ||
    !r.documentId.length ||
    r.documentId.length > 4096 ||
    !hostId(r.hostSlideId) ||
    !hostId(r.oldShapeId) ||
    !r.scope ||
    typeof r.scope !== 'object' ||
    Array.isArray(r.scope) ||
    Object.keys(r.scope).some((key) => !['slideIds', 'shapeIds'].includes(key)) ||
    !ids(r.scope.slideIds) ||
    !r.scope.slideIds.includes(r.hostSlideId) ||
    (r.scope.shapeIds !== undefined &&
      (!ids(r.scope.shapeIds) || !r.scope.shapeIds.includes(r.oldShapeId))) ||
    !snapshot(r.original) ||
    r.original.slideId !== r.hostSlideId ||
    r.original.shapeId !== r.oldShapeId ||
    !validPresentationImageBackupMetadata(r.backup) ||
    r.backup.attachmentId !== r.original.mediaDigest ||
    !['pending', 'complete', 'undo_pending', 'undone'].includes(r.state)
  )
    return false
  if (
    r.reapplies !== undefined &&
    (!id(r.reapplies) || r.reapplies === r.changeId || !r.sourceBackup)
  )
    return false
  if (
    r.sourceBackup !== undefined &&
    (!validPresentationImageBackupMetadata(r.sourceBackup) ||
      r.sourceBackup.attachmentId !== r.assetDigest)
  )
    return false
  if (
    r.insertedShapeId !== undefined &&
    (!hostId(r.insertedShapeId) || r.original.shapeIds.includes(r.insertedShapeId))
  )
    return false
  if (
    r.after !== undefined &&
    (!r.insertedShapeId ||
      !snapshot(r.after) ||
      r.after.slideId !== r.hostSlideId ||
      r.after.shapeId !== r.insertedShapeId ||
      r.after.mediaDigest !== r.assetDigest)
  )
    return false
  if (
    r.undoBaseline !== undefined &&
    (!r.after ||
      !snapshot(r.undoBaseline) ||
      JSON.stringify(r.undoBaseline) !== JSON.stringify(r.after))
  )
    return false
  if (
    r.restoredShapeId !== undefined &&
    (!r.undoBaseline ||
      !hostId(r.restoredShapeId) ||
      r.undoBaseline.shapeIds.includes(r.restoredShapeId))
  )
    return false
  if (r.state === 'pending' && (r.after || r.undoBaseline || r.restoredShapeId)) return false
  if (
    r.state === 'complete' &&
    (!r.insertedShapeId || !r.after || r.undoBaseline || r.restoredShapeId)
  )
    return false
  if (r.state === 'undo_pending' && (!r.after || !r.undoBaseline)) return false
  if (r.state === 'undone' && (!r.after || !r.undoBaseline || !r.restoredShapeId)) return false
  if (
    r.capture !== undefined &&
    (!['complete', 'undone'].includes(r.state) ||
      !validExistingVisualCapture(r.capture) ||
      r.capture.hostSlideId !== r.hostSlideId)
  )
    return false
  if (
    r.review !== undefined &&
    (!['complete', 'undone'].includes(r.state) ||
      !validExistingVisualReview(r.review) ||
      r.review.hostSlideId !== r.hostSlideId ||
      (r.capture !== undefined &&
        (r.capture.screenshotDigest !== r.review.screenshotDigest ||
          r.capture.capturedAt !== r.review.capturedAt)))
  )
    return false
  return bytes(r) + existingImageReservedBytes(r) <= 192 * 1024
}

export function validExistingImageTransition(
  before: PresentationExistingImageChange | undefined,
  after: PresentationExistingImageChange,
): boolean {
  if (!validatePresentationExistingImageChange(after)) return false
  if (!before) return after.state === 'pending' && after.insertedShapeId === undefined
  if (!validatePresentationExistingImageChange(before)) return false
  const core = (r: PresentationExistingImageChange) =>
    JSON.stringify({
      ...r,
      state: undefined,
      insertedShapeId: undefined,
      after: undefined,
      undoBaseline: undefined,
      restoredShapeId: undefined,
      capture: undefined,
      review: undefined,
    })
  if (core(before) !== core(after)) return false
  if (before.state === 'pending')
    return (
      (after.state === 'pending' && !before.insertedShapeId && !!after.insertedShapeId) ||
      (after.state === 'complete' &&
        !!after.insertedShapeId &&
        (!before.insertedShapeId || before.insertedShapeId === after.insertedShapeId))
    )
  if (before.state === 'complete')
    return (
      (after.state === 'complete' &&
        before.review === undefined &&
        ((after.capture !== undefined &&
          JSON.stringify(before.capture) !== JSON.stringify(after.capture) &&
          after.review === undefined) ||
          (!!after.review && JSON.stringify(before.capture) === JSON.stringify(after.capture))) &&
        before.insertedShapeId === after.insertedShapeId &&
        JSON.stringify(before.after) === JSON.stringify(after.after)) ||
      (after.state === 'undo_pending' &&
        after.review === undefined &&
        after.capture === undefined &&
        before.insertedShapeId === after.insertedShapeId &&
        JSON.stringify(before.after) === JSON.stringify(after.after) &&
        after.restoredShapeId === undefined)
    )
  if (before.state === 'undo_pending')
    return (
      before.insertedShapeId === after.insertedShapeId &&
      JSON.stringify(before.after) === JSON.stringify(after.after) &&
      JSON.stringify(before.undoBaseline) === JSON.stringify(after.undoBaseline) &&
      ((after.state === 'undo_pending' && !before.restoredShapeId && !!after.restoredShapeId) ||
        (after.state === 'undone' &&
          !!after.restoredShapeId &&
          (!before.restoredShapeId || before.restoredShapeId === after.restoredShapeId)))
    )
  if (before.state === 'undone')
    return (
      after.state === 'undone' &&
      before.review === undefined &&
      ((after.capture !== undefined &&
        JSON.stringify(before.capture) !== JSON.stringify(after.capture) &&
        after.review === undefined) ||
        (!!after.review && JSON.stringify(before.capture) === JSON.stringify(after.capture))) &&
      before.restoredShapeId === after.restoredShapeId
    )
  return false
}
