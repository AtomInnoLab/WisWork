import type { PowerPointMasterOperation } from './browser-powerpoint-adapter.js'
import { masterOperationKey, MASTER_PATTERN_TYPES } from './presentation-master-program.js'

export interface MasterBackupRef {
  key: string
  sha256: string
  sizeBytes: number
}
export type StoredMasterOperation =
  | Exclude<PowerPointMasterOperation, { op: 'set_master_background' }>
  | {
      op: 'set_master_background'
      master_id: string
      fill:
        | Exclude<
            Extract<PowerPointMasterOperation, { op: 'set_master_background' }>['fill'],
            { type: 'picture_or_texture' }
          >
        | { type: 'picture_or_texture'; imageRef: MasterBackupRef; transparency: number }
    }
export interface PresentationNativeMasterChange {
  version: 1
  kind: 'native_master'
  changeId: string
  documentId: string
  intent: string
  snapshotRef: MasterBackupRef
  operations: StoredMasterOperation[]
  inverseOperations: StoredMasterOperation[]
  scope: { masterIds: string[]; affectedPageCount: number }
  nextIndex: number
  state: 'applying' | 'applied' | 'undoing' | 'undone'
  currentProofRef: MasterBackupRef
  pending?: {
    direction: 'forward' | 'undo'
    index: number
    beforeProofRef: MasterBackupRef
    afterProofRef?: MasterBackupRef
  }
  receipts: Array<{ direction: 'forward' | 'undo'; index: number; proofRef: MasterBackupRef }>
  reviews: Array<{ slideId: string; reviewRef: MasterBackupRef }>
}
const bytes = (v: unknown) => new TextEncoder().encode(JSON.stringify(v)).byteLength
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const obj = (v: unknown): v is Record<string, unknown> =>
  !!v && typeof v === 'object' && !Array.isArray(v)
const keys = (v: unknown, required: string[], optional: string[] = []) =>
  obj(v) &&
  required.every((k) => Object.hasOwn(v, k)) &&
  Object.keys(v).every((k) => [...required, ...optional].includes(k))
const id = (v: unknown) => typeof v === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(v)
export const validMasterHostId = (v: unknown): v is string =>
  typeof v === 'string' &&
  !!v.trim() &&
  v.length <= 256 &&
  ![...v].some((c) => c.charCodeAt(0) < 32 || (c.charCodeAt(0) >= 127 && c.charCodeAt(0) <= 159))
export const validMasterBackupRef = (v: unknown): v is MasterBackupRef =>
  obj(v) &&
  keys(v, ['key', 'sha256', 'sizeBytes']) &&
  typeof v.key === 'string' &&
  /^(snapshot|(?:page|image|receipt)-[0-9]+)$/.test(v.key) &&
  typeof v.sha256 === 'string' &&
  /^[a-f0-9]{64}$/.test(v.sha256) &&
  Number.isSafeInteger(v.sizeBytes) &&
  Number(v.sizeBytes) > 0 &&
  Number(v.sizeBytes) <= 8 * 1024 * 1024
const color = (v: unknown) => typeof v === 'string' && /^#[0-9A-Fa-f]{6}$/.test(v)
const ratio = (v: unknown) => typeof v === 'number' && Number.isFinite(v) && v >= 0 && v <= 1
const themeSlots = [
  'Accent1',
  'Accent2',
  'Accent3',
  'Accent4',
  'Accent5',
  'Accent6',
  'Dark1',
  'Dark2',
  'Light1',
  'Light2',
  'Hyperlink',
  'FollowedHyperlink',
]
export function validStoredMasterOperation(v: unknown): v is StoredMasterOperation {
  if (!obj(v) || !validMasterHostId(v.master_id)) return false
  if (v.op === 'set_master_theme_color')
    return (
      keys(v, ['op', 'master_id', 'theme_color', 'color']) &&
      typeof v.theme_color === 'string' &&
      themeSlots.includes(v.theme_color) &&
      color(v.color)
    )
  if (v.op === 'set_layout_background_following')
    return (
      keys(v, ['op', 'master_id', 'layout_id', 'follow_master', 'show_master_graphics']) &&
      validMasterHostId(v.layout_id) &&
      typeof v.follow_master === 'boolean' &&
      typeof v.show_master_graphics === 'boolean'
    )
  if (v.op !== 'set_master_background' || !keys(v, ['op', 'master_id', 'fill']) || !obj(v.fill))
    return false
  const f = v.fill
  if (f.type === 'solid')
    return keys(f, ['type', 'color', 'transparency']) && color(f.color) && ratio(f.transparency)
  if (f.type === 'gradient')
    return (
      keys(f, ['type', 'gradient_type']) &&
      typeof f.gradient_type === 'string' &&
      ['Linear', 'Radial', 'Rectangular', 'Path', 'ShadeFromTitle'].includes(f.gradient_type)
    )
  if (f.type === 'pattern')
    return (
      keys(f, ['type', 'pattern', 'foreground_color', 'background_color']) &&
      typeof f.pattern === 'string' &&
      (MASTER_PATTERN_TYPES as readonly string[]).includes(f.pattern) &&
      color(f.foreground_color) &&
      color(f.background_color)
    )
  return (
    f.type === 'picture_or_texture' &&
    keys(f, ['type', 'imageRef', 'transparency']) &&
    validMasterBackupRef(f.imageRef) &&
    ratio(f.transparency)
  )
}
export function nativeMasterReservedBytes(r: PresentationNativeMasterChange): number {
  // Reserve every future host receipt and pending proof before authorizing any host mutation.
  const ref = { key: 'x'.repeat(128), sha256: 'f'.repeat(64), sizeBytes: 8 * 1024 * 1024 }
  return Math.max(
    0,
    bytes({
      ...r,
      receipts: Array.from({ length: r.operations.length * 2 }, (_, index) => ({
        direction: 'forward',
        index,
        proofRef: ref,
      })),
      pending: { direction: 'forward', index: 31, beforeProofRef: ref, afterProofRef: ref },
    }) - bytes(r),
  )
}
export function validatePresentationNativeMasterChange(
  v: unknown,
): v is PresentationNativeMasterChange {
  if (
    !keys(
      v,
      [
        'version',
        'kind',
        'changeId',
        'documentId',
        'intent',
        'snapshotRef',
        'operations',
        'inverseOperations',
        'scope',
        'nextIndex',
        'state',
        'currentProofRef',
        'receipts',
        'reviews',
      ],
      ['pending'],
    )
  )
    return false
  const r = v as unknown as PresentationNativeMasterChange
  if (
    r.version !== 1 ||
    r.kind !== 'native_master' ||
    !id(r.changeId) ||
    typeof r.documentId !== 'string' ||
    !r.documentId ||
    r.documentId.length > 2048 ||
    typeof r.intent !== 'string' ||
    !r.intent.trim() ||
    r.intent.length > 300 ||
    !validMasterBackupRef(r.snapshotRef) ||
    !validMasterBackupRef(r.currentProofRef) ||
    !Array.isArray(r.operations) ||
    !r.operations.length ||
    r.operations.length > 32 ||
    !r.operations.every(validStoredMasterOperation) ||
    !Array.isArray(r.inverseOperations) ||
    r.inverseOperations.length !== r.operations.length ||
    !r.inverseOperations.every(validStoredMasterOperation)
  )
    return false
  const opKey = (o: StoredMasterOperation) => masterOperationKey(o as PowerPointMasterOperation)
  if (
    new Set(r.operations.map(opKey)).size !== r.operations.length ||
    r.inverseOperations.some((o, i) => opKey(o) !== opKey(r.operations[i]!))
  )
    return false
  if (
    !keys(r.scope, ['masterIds', 'affectedPageCount']) ||
    !Array.isArray(r.scope.masterIds) ||
    !r.scope.masterIds.length ||
    !r.scope.masterIds.every(validMasterHostId) ||
    new Set(r.scope.masterIds).size !== r.scope.masterIds.length ||
    !same(
      [...r.scope.masterIds].sort(),
      [...new Set(r.operations.map((o) => o.master_id))].sort(),
    ) ||
    !Number.isSafeInteger(r.scope.affectedPageCount) ||
    r.scope.affectedPageCount < 0 ||
    !Number.isSafeInteger(r.nextIndex) ||
    r.nextIndex < 0 ||
    r.nextIndex > r.operations.length ||
    !['applying', 'applied', 'undoing', 'undone'].includes(r.state)
  )
    return false
  if (
    (r.state === 'applied' && r.nextIndex !== r.operations.length) ||
    (r.state === 'undone' && (r.nextIndex !== 0 || r.pending)) ||
    (r.state === 'applying' && r.nextIndex === r.operations.length) ||
    (r.state === 'undoing' && r.nextIndex === 0)
  )
    return false
  if (
    r.pending &&
    (!keys(r.pending, ['direction', 'index', 'beforeProofRef'], ['afterProofRef']) ||
      !['forward', 'undo'].includes(r.pending.direction) ||
      r.pending.index !== (r.pending.direction === 'forward' ? r.nextIndex : r.nextIndex - 1) ||
      r.pending.index < 0 ||
      r.pending.index >= r.operations.length ||
      !validMasterBackupRef(r.pending.beforeProofRef) ||
      !same(r.pending.beforeProofRef, r.currentProofRef) ||
      (r.pending.afterProofRef !== undefined && !validMasterBackupRef(r.pending.afterProofRef)) ||
      r.state !== (r.pending.direction === 'forward' ? 'applying' : 'undoing'))
  )
    return false
  if (
    !Array.isArray(r.receipts) ||
    r.receipts.length > r.operations.length * 2 ||
    !r.receipts.every(
      (x) =>
        keys(x, ['direction', 'index', 'proofRef']) &&
        ['forward', 'undo'].includes(x.direction) &&
        Number.isSafeInteger(x.index) &&
        x.index >= 0 &&
        x.index < r.operations.length &&
        validMasterBackupRef(x.proofRef),
    )
  )
    return false
  let cursor = 0,
    undo = false
  for (const receipt of r.receipts) {
    if (receipt.direction === 'forward') {
      if (undo || receipt.index !== cursor) return false
      cursor++
    } else {
      undo = true
      if (receipt.index !== cursor - 1) return false
      cursor--
    }
  }
  if (
    cursor !== r.nextIndex ||
    (r.receipts.length && !same(r.currentProofRef, r.receipts.at(-1)!.proofRef))
  )
    return false
  if (
    !Array.isArray(r.reviews) ||
    new Set(r.reviews.map((x) => x.slideId)).size !== r.reviews.length ||
    !r.reviews.every(
      (x) =>
        keys(x, ['slideId', 'reviewRef']) &&
        validMasterHostId(x.slideId) &&
        validMasterBackupRef(x.reviewRef),
    ) ||
    bytes(r) > 192 * 1024 ||
    bytes(r) + nativeMasterReservedBytes(r) > 192 * 1024
  )
    return false
  return true
}
export function validNativeMasterTransition(
  before: PresentationNativeMasterChange | undefined,
  after: PresentationNativeMasterChange,
): boolean {
  if (
    !validatePresentationNativeMasterChange(after) ||
    (before && !validatePresentationNativeMasterChange(before))
  )
    return false
  if (!before)
    return (
      after.state === 'applying' &&
      after.nextIndex === 0 &&
      !after.pending &&
      !after.receipts.length &&
      !after.reviews.length
    )
  if (same(before, after)) return true
  const core = (r: PresentationNativeMasterChange) => ({
    version: r.version,
    kind: r.kind,
    changeId: r.changeId,
    documentId: r.documentId,
    intent: r.intent,
    snapshotRef: r.snapshotRef,
    operations: r.operations,
    inverseOperations: r.inverseOperations,
    scope: r.scope,
  })
  if (!same(core(before), core(after))) return false
  const execution = (r: PresentationNativeMasterChange) => ({
    nextIndex: r.nextIndex,
    state: r.state,
    currentProofRef: r.currentProofRef,
    pending: r.pending,
    receipts: r.receipts,
  })
  if (same(execution(before), execution(after))) {
    const changed = after.reviews.filter((x) => !before.reviews.some((y) => same(x, y)))
    return (
      !before.pending &&
      ['applied', 'undone'].includes(before.state) &&
      changed.length === 1 &&
      before.reviews.every((x) => after.reviews.some((y) => y.slideId === x.slideId)) &&
      after.reviews.length <= before.reviews.length + 1
    )
  }
  if (after.reviews.length || (before.reviews.length && after.reviews.length)) return false
  if (!same(before.currentProofRef, after.currentProofRef) && !before.pending?.afterProofRef)
    return false
  if (!before.pending && after.pending)
    return (
      same(before.receipts, after.receipts) &&
      before.nextIndex === after.nextIndex &&
      before.state === after.state &&
      same(after.pending.beforeProofRef, before.currentProofRef) &&
      !after.pending.afterProofRef
    )
  if (before.pending && after.pending)
    return (
      same(before.receipts, after.receipts) &&
      before.nextIndex === after.nextIndex &&
      before.state === after.state &&
      same(
        { ...before.pending, afterProofRef: undefined },
        { ...after.pending, afterProofRef: undefined },
      ) &&
      !before.pending.afterProofRef &&
      !!after.pending.afterProofRef
    )
  if (before.pending && !after.pending) {
    if (before.pending.afterProofRef) {
      const direction = before.pending.direction,
        index = before.pending.index
      return (
        after.nextIndex === before.nextIndex + (direction === 'forward' ? 1 : -1) &&
        same(after.currentProofRef, before.pending.afterProofRef) &&
        same(after.receipts, [
          ...before.receipts,
          { direction, index, proofRef: before.pending.afterProofRef },
        ]) &&
        after.state ===
          (direction === 'forward'
            ? after.nextIndex === after.operations.length
              ? 'applied'
              : 'applying'
            : after.nextIndex === 0
              ? 'undone'
              : 'undoing')
      )
    }
    // Explicit reconciliation of an exact before proof cancels forward replay.
    return (
      after.nextIndex === before.nextIndex &&
      same(after.receipts, before.receipts) &&
      same(after.currentProofRef, before.currentProofRef) &&
      after.state === (after.nextIndex === 0 ? 'undone' : 'undoing')
    )
  }
  return (
    !before.pending &&
    !after.pending &&
    same(before.receipts, after.receipts) &&
    same(before.currentProofRef, after.currentProofRef) &&
    before.nextIndex === after.nextIndex &&
    after.state === (after.nextIndex === 0 ? 'undone' : 'undoing') &&
    ['applying', 'applied'].includes(before.state)
  )
}
