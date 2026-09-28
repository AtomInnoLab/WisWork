import type { PresentationImportRecord } from './presentation-delivery.js'
import { validPresentationImportRecord, validSourceSlideId } from './presentation-page-delivery.js'

export interface PresentationPageReplacement {
  version: 1 | 2
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
  /** Package digests of imported pages outside the revision target, in host order. */
  untouchedSlideDigests?: { slideId: string; digest: string }[]
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
  parentReceipt?: PresentationImportRecord
  childReceipt?: PresentationImportRecord
  restoredSlideId?: string
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
        'untouchedSlideDigests',
        'state',
        'newSlideId',
        'parentReceipt',
        'childReceipt',
        'restoredSlideId',
      ].includes(key),
    ) &&
    (r.version === 1 || r.version === 2) &&
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
    (r.version === 1
      ? r.untouchedSlideDigests === undefined
      : Array.isArray(r.untouchedSlideDigests) &&
        r.untouchedSlideDigests.length <= 31 &&
        r.untouchedSlideDigests.every(
          (item, index) =>
            item &&
            Object.keys(item).sort().join(',') === 'digest,slideId' &&
            hostId(item.slideId) &&
            item.slideId !== r.oldSlideId &&
            digest(item.digest) &&
            r.beforeSlideIds.indexOf(item.slideId) >
              (index ? r.beforeSlideIds.indexOf(r.untouchedSlideDigests![index - 1]!.slideId) : -1),
        )) &&
    [
      'pending',
      'inserted',
      'staged',
      'discard_pending',
      'discarded',
      'commit_pending',
      'applied',
      'undo_pending',
      'restore_inserted',
      'undone',
    ].includes(r.state) &&
    (r.state === 'pending'
      ? r.newSlideId === undefined
      : hostId(r.newSlideId) && !r.beforeSlideIds.includes(r.newSlideId)) &&
    (['restore_inserted', 'undone'].includes(r.state)
      ? hostId(r.restoredSlideId) &&
        r.restoredSlideId !== r.newSlideId &&
        !r.beforeSlideIds.includes(r.restoredSlideId)
      : r.restoredSlideId === undefined) &&
    validReceipts(r) &&
    // Reserve the largest UTF-8 host ID before insertion so recording it cannot exceed the limit.
    new TextEncoder().encode(
      JSON.stringify({
        ...r,
        newSlideId: '\uffff'.repeat(256),
        restoredSlideId: '\uffff'.repeat(256),
        state: 'restore_inserted',
      }),
    ).byteLength <=
      192 * 1024
  )
}

function validReceipts(r: PresentationPageReplacement): boolean {
  if (
    !['commit_pending', 'applied', 'undo_pending', 'restore_inserted', 'undone'].includes(r.state)
  )
    return r.parentReceipt === undefined && r.childReceipt === undefined
  const parent = r.parentReceipt,
    child = r.childReceipt
  if (
    !validPresentationImportRecord(parent) ||
    !validPresentationImportRecord(child) ||
    parent.state !== 'complete' ||
    child.state !== 'complete' ||
    parent.documentId !== r.documentId ||
    child.documentId !== r.documentId ||
    parent.checkpoint?.version !== 2 ||
    child.checkpoint?.version !== 2
  )
    return false
  const p = parent.checkpoint,
    c = child.checkpoint
  const index = p.pageIds!.indexOf(r.pageId)
  const otherSlideIds = p.completed
    .filter((_, position) => position !== index)
    .map((page) => page.slideId)
  return (
    index >= 0 &&
    p.artifactDigest === r.parentArtifactDigest &&
    JSON.stringify(p.pageIds) === JSON.stringify(c.pageIds) &&
    JSON.stringify(p.baselineSlideIds) === JSON.stringify(c.baselineSlideIds) &&
    p.completed[index].slideId === r.oldSlideId &&
    c.completed[index].slideId === r.newSlideId &&
    c.sourceSlideIds[index] === r.sourceSlideId &&
    (r.version === 1 ||
      (r.untouchedSlideDigests!.length === otherSlideIds.length &&
        otherSlideIds.every((id) =>
          r.untouchedSlideDigests!.some((item) => item.slideId === id),
        ))) &&
    p.completed.every(
      (page, i) =>
        hostId(page.slideId) &&
        r.beforeSlideIds.includes(page.slideId) &&
        (i === index ||
          (page.slideId === c.completed[i].slideId &&
            page.sourceSlideId === c.completed[i].sourceSlideId)),
    )
  )
}
