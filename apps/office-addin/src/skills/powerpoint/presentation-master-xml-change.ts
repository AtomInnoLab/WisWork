import { validPackageBackupRef, validPackageHostId } from './presentation-package-change.js'
import type { PackageBackupRef } from './presentation-package-backup.js'
export const MASTER_XML_ACTIONS = [
  'original_probe_stage',
  'original_probe',
  'original_probe_reset',
  'original_probe_delete',
  'stage',
  'stage_master',
  'delete_stage_master',
  'import_probe_stage',
  'import_probe',
  'import_probe_reset',
  'import_probe_delete',
  'forward_page',
  'delete_source',
  'restore_stage',
  'restore_stage_master',
  'restore_probe_stage',
  'restore_probe',
  'restore_probe_reset',
  'restore_probe_delete',
  'restore_page',
  'restore_page_stage',
  'restore_page_associate',
  'restore_page_delete',
  'delete_applied_source',
  'discard',
  'verify',
] as const
export type MasterXmlAction = (typeof MASTER_XML_ACTIONS)[number]
export type MasterXmlState =
  | 'prepared'
  | 'probing_original'
  | 'staged'
  | 'probing_imported'
  | 'applying'
  | 'applied'
  | 'recovery_required'
  | 'restoring'
  | 'probing_restored'
  | 'restoring_pages'
  | 'undone'
  | 'discarded'
export interface MasterXmlNativeSource {
  slideId: string
  masterId: string
  layoutId: string
}
export interface MasterXmlCursor {
  phase: MasterXmlAction
  index: number
  substep: number
}
export interface PresentationMasterXmlChange {
  version: 1
  kind: 'master_xml'
  changeId: string
  documentId: string
  intent: string
  sourceSlideId: string
  packageSourceSlideId: string
  snapshotRef: PackageBackupRef
  preparedRef: PackageBackupRef
  currentProofRef: PackageBackupRef
  state: MasterXmlState
  cursor: MasterXmlCursor
  reviewSequence?: number
  receiptCount: number
  scope: {
    originalMasterId: string
    affectedPageCount: number
    originalLayoutCount: number
    affectedMasterCount?: number
  }
  introducedMasterCount: number
  inventoryCleanupVerified: false
  stagedSource?: MasterXmlNativeSource
  restoredSource?: MasterXmlNativeSource
  originalProbeSource?: MasterXmlNativeSource
  importedProbeSource?: MasterXmlNativeSource
  restoredProbeSource?: MasterXmlNativeSource
  pending?: {
    action: MasterXmlAction
    index: number
    targetSlideId?: string
    targetMasterId?: string
    targetLayoutId?: string
    beforeProofRef: PackageBackupRef
    inserted?: MasterXmlNativeSource
    afterProofRef?: PackageBackupRef
    currentBackupRef?: PackageBackupRef
    recoveryPreimageRef?: PackageBackupRef
  }
  reviews: { slideId: string; reviewRef: PackageBackupRef }[]
}
const exact = (v: unknown, req: string[], opt: string[] = []): v is Record<string, any> =>
  !!v &&
  typeof v === 'object' &&
  !Array.isArray(v) &&
  req.every((k) => Object.hasOwn(v, k)) &&
  Object.keys(v).every((k) => req.includes(k) || opt.includes(k))
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const count = (v: unknown): v is number => Number.isSafeInteger(v) && Number(v) >= 0
const source = (v: unknown): v is MasterXmlNativeSource =>
  exact(v, ['slideId', 'masterId', 'layoutId']) &&
  [v.slideId, v.masterId, v.layoutId].every(validPackageHostId)
const cursor = (v: unknown): v is MasterXmlCursor =>
  exact(v, ['phase', 'index', 'substep']) &&
  MASTER_XML_ACTIONS.includes(v.phase) &&
  count(v.index) &&
  count(v.substep) &&
  v.substep <= 3
const states: MasterXmlState[] = [
  'prepared',
  'probing_original',
  'staged',
  'probing_imported',
  'applying',
  'applied',
  'recovery_required',
  'restoring',
  'probing_restored',
  'restoring_pages',
  'undone',
  'discarded',
]
const stageSlots = {
  original_probe_stage: 'originalProbeSource',
  stage: 'stagedSource',
  import_probe_stage: 'importedProbeSource',
  restore_stage: 'restoredSource',
  restore_probe_stage: 'restoredProbeSource',
} as const
const slots = Object.values(stageSlots)
const completedStates: Record<MasterXmlAction, readonly MasterXmlState[]> = {
  original_probe_stage: ['probing_original'],
  original_probe: ['probing_original'],
  original_probe_reset: ['probing_original'],
  original_probe_delete: ['probing_original', 'prepared'],
  stage: ['staged'],
  stage_master: ['staged'],
  delete_stage_master: ['applying', 'applied'],
  import_probe_stage: ['probing_imported'],
  import_probe: ['probing_imported'],
  import_probe_reset: ['probing_imported'],
  import_probe_delete: ['probing_imported', 'staged'],
  forward_page: ['applying'],
  delete_source: ['applied'],
  restore_stage: ['restoring'],
  restore_stage_master: ['restoring'],
  restore_probe_stage: ['probing_restored'],
  restore_probe: ['probing_restored'],
  restore_probe_reset: ['probing_restored'],
  restore_probe_delete: ['restoring', 'probing_restored'],
  restore_page: ['restoring_pages'],
  restore_page_stage: ['restoring_pages'],
  restore_page_associate: ['restoring_pages'],
  restore_page_delete: ['restoring_pages'],
  delete_applied_source: ['restoring_pages', 'undone'],
  discard: [
    'prepared',
    'probing_original',
    'staged',
    'probing_imported',
    'recovery_required',
    'discarded',
  ],
  verify: [
    'prepared',
    'probing_original',
    'staged',
    'probing_imported',
    'applying',
    'applied',
    'restoring',
    'probing_restored',
    'restoring_pages',
    'undone',
    'discarded',
  ],
}

const allowedActions: Record<MasterXmlState, readonly MasterXmlAction[]> = {
  prepared: ['original_probe_stage', 'stage', 'discard', 'verify'],
  probing_original: [
    'original_probe',
    'original_probe_reset',
    'original_probe_delete',
    'stage',
    'discard',
    'verify',
  ],
  staged: [
    'delete_stage_master',
    'stage_master',
    'import_probe_stage',
    'forward_page',
    'delete_source',
    'discard',
    'verify',
  ],
  probing_imported: [
    'import_probe',
    'import_probe_reset',
    'import_probe_delete',
    'forward_page',
    'discard',
    'verify',
  ],
  applying: [
    'delete_stage_master',
    'forward_page',
    'delete_source',
    'restore_stage',
    'restore_page',
    'discard',
    'verify',
  ],
  applied: ['restore_stage', 'restore_page', 'verify'],
  recovery_required: [
    'discard',
    'restore_stage',
    'restore_page',
    'restore_page_stage',
    'restore_page_associate',
    'restore_page_delete',
    'delete_applied_source',
    'verify',
  ],
  restoring: ['restore_stage_master', 'restore_probe_stage', 'restore_page', 'verify'],
  probing_restored: [
    'restore_probe',
    'restore_probe_reset',
    'restore_probe_delete',
    'restore_page',
    'verify',
  ],
  restoring_pages: [
    'restore_page',
    'restore_page_stage',
    'restore_page_associate',
    'restore_page_delete',
    'delete_applied_source',
    'verify',
  ],
  undone: [],
  discarded: [],
}
export function masterXmlReservedBytes(r: PresentationMasterXmlChange): number {
  const id = '界'.repeat(256),
    native = { slideId: id, masterId: id, layoutId: id },
    ref = { key: 'receipt-' + '9'.repeat(120), sha256: 'f'.repeat(64), sizeBytes: 8388608 }
  const maximum = {
    ...r,
    cursor: { phase: 'restore_page_associate', index: Number.MAX_SAFE_INTEGER, substep: 3 },
    receiptCount: Number.MAX_SAFE_INTEGER,
    reviewSequence: Number.MAX_SAFE_INTEGER,
    introducedMasterCount: Number.MAX_SAFE_INTEGER,
    currentProofRef: ref,
    stagedSource: native,
    restoredSource: native,
    originalProbeSource: native,
    importedProbeSource: native,
    restoredProbeSource: native,
    pending: {
      action: 'restore_page_stage',
      index: Number.MAX_SAFE_INTEGER,
      targetSlideId: id,
      targetMasterId: id,
      targetLayoutId: id,
      beforeProofRef: ref,
      inserted: native,
      afterProofRef: ref,
      currentBackupRef: ref,
      recoveryPreimageRef: ref,
    },
    reviews: [...r.reviews, { slideId: id, reviewRef: ref }],
  }
  return Math.max(
    0,
    new TextEncoder().encode(JSON.stringify(maximum)).length -
      new TextEncoder().encode(JSON.stringify(r)).length,
  )
}
export const presentationMasterXmlReservedBytes = masterXmlReservedBytes
export function validatePresentationMasterXmlChange(v: unknown): v is PresentationMasterXmlChange {
  if (
    !exact(
      v,
      [
        'version',
        'kind',
        'changeId',
        'documentId',
        'intent',
        'sourceSlideId',
        'packageSourceSlideId',
        'snapshotRef',
        'preparedRef',
        'currentProofRef',
        'state',
        'cursor',
        'receiptCount',
        'scope',
        'introducedMasterCount',
        'inventoryCleanupVerified',
        'reviews',
      ],
      [...slots, 'pending', 'reviewSequence'],
    )
  )
    return false
  if (
    v.version !== 1 ||
    v.kind !== 'master_xml' ||
    typeof v.changeId !== 'string' ||
    !/^[A-Za-z0-9_-]{1,128}$/.test(v.changeId) ||
    typeof v.documentId !== 'string' ||
    !v.documentId.trim() ||
    v.documentId.length > 2048 ||
    typeof v.intent !== 'string' ||
    v.intent.length > 8192 ||
    !validPackageHostId(v.sourceSlideId) ||
    typeof v.packageSourceSlideId !== 'string' ||
    !/^[1-9][0-9]{0,9}#$/.test(v.packageSourceSlideId) ||
    Number(v.packageSourceSlideId.slice(0, -1)) < 256 ||
    Number(v.packageSourceSlideId.slice(0, -1)) > 4294967295
  )
    return false
  if (
    ![v.snapshotRef, v.preparedRef, v.currentProofRef].every(validPackageBackupRef) ||
    v.snapshotRef.key !== 'snapshot' ||
    v.preparedRef.key !== 'page-0' ||
    v.currentProofRef.key !== `receipt-${v.receiptCount}` ||
    !states.includes(v.state) ||
    !cursor(v.cursor) ||
    !count(v.receiptCount) ||
    !count(v.introducedMasterCount) ||
    (v.reviewSequence !== undefined && !count(v.reviewSequence)) ||
    v.inventoryCleanupVerified !== false
  )
    return false
  if (
    !exact(
      v.scope,
      ['originalMasterId', 'affectedPageCount', 'originalLayoutCount'],
      ['affectedMasterCount'],
    ) ||
    !validPackageHostId(v.scope.originalMasterId) ||
    !count(v.scope.affectedPageCount) ||
    v.scope.affectedPageCount < 1 ||
    !count(v.scope.originalLayoutCount) ||
    v.scope.originalLayoutCount < 1 ||
    (v.scope.affectedMasterCount !== undefined &&
      (!count(v.scope.affectedMasterCount) || v.scope.affectedMasterCount < 1))
  )
    return false
  if (slots.some((k) => v[k] !== undefined && !source(v[k]))) return false
  const nativeIds = slots.flatMap((k) => (v[k] ? [v[k].slideId] : []))
  if (new Set(nativeIds).size !== nativeIds.length || nativeIds.includes(v.sourceSlideId))
    return false
  if (
    [
      'staged',
      'probing_imported',
      'applying',
      'applied',
      'restoring',
      'probing_restored',
      'restoring_pages',
      'undone',
    ].includes(v.state) &&
    !v.stagedSource
  )
    return false
  if (['restoring', 'probing_restored'].includes(v.state) && !v.restoredSource) return false
  if (
    !Array.isArray(v.reviews) ||
    new Set(v.reviews.map((r: any) => r.slideId)).size !== v.reviews.length ||
    !v.reviews.every(
      (r: any) =>
        exact(r, ['slideId', 'reviewRef']) &&
        validPackageHostId(r.slideId) &&
        validPackageBackupRef(r.reviewRef),
    ) ||
    (v.reviews.length && !['applied', 'undone'].includes(v.state))
  )
    return false
  if (v.pending) {
    const p = v.pending
    if (
      !exact(
        p,
        ['action', 'index', 'beforeProofRef'],
        [
          'targetSlideId',
          'targetMasterId',
          'targetLayoutId',
          'inserted',
          'afterProofRef',
          'currentBackupRef',
          'recoveryPreimageRef',
        ],
      ) ||
      !MASTER_XML_ACTIONS.includes(p.action) ||
      !allowedActions[v.state as MasterXmlState].includes(p.action) ||
      ([
        'original_probe',
        'original_probe_reset',
        'import_probe',
        'import_probe_reset',
        'forward_page',
        'restore_probe',
        'restore_probe_reset',
        'restore_page',
        'restore_page_associate',
      ].includes(p.action) &&
        ![p.targetSlideId, p.targetMasterId, p.targetLayoutId].every(validPackageHostId)) ||
      ([
        'original_probe_delete',
        'import_probe_delete',
        'delete_source',
        'delete_stage_master',
        'restore_probe_delete',
        'restore_page_delete',
        'delete_applied_source',
        'discard',
      ].includes(p.action) &&
        !validPackageHostId(p.targetSlideId)) ||
      p.action !== v.cursor.phase ||
      p.index !== v.cursor.index ||
      !same(p.beforeProofRef, v.currentProofRef) ||
      !validPackageBackupRef(p.beforeProofRef) ||
      ['targetSlideId', 'targetMasterId', 'targetLayoutId'].some(
        (k) => p[k] !== undefined && !validPackageHostId(p[k]),
      ) ||
      (p.inserted !== undefined && !source(p.inserted)) ||
      ['currentBackupRef', 'recoveryPreimageRef'].some(
        (k) => p[k] !== undefined && !validPackageBackupRef(p[k]),
      ) ||
      (p.recoveryPreimageRef !== undefined &&
        !['forward_page', 'restore_page', 'restore_page_associate'].includes(p.action)) ||
      (p.currentBackupRef !== undefined &&
        [
          'original_probe_stage',
          'stage',
          'stage_master',
          'import_probe_stage',
          'restore_stage',
          'restore_stage_master',
          'restore_probe_stage',
          'restore_page_stage',
          'verify',
        ].includes(p.action)) ||
      (p.afterProofRef !== undefined &&
        (!validPackageBackupRef(p.afterProofRef) ||
          p.afterProofRef.key !== `receipt-${v.receiptCount + 1}`)) ||
      (p.inserted !== undefined &&
        !Object.hasOwn(stageSlots, p.action) &&
        !['restore_page_stage', 'stage_master', 'restore_stage_master'].includes(p.action)) ||
      (p.afterProofRef !== undefined &&
        (Object.hasOwn(stageSlots, p.action) ||
          ['restore_page_stage', 'stage_master', 'restore_stage_master'].includes(p.action)) &&
        !p.inserted)
    )
      return false
  }
  const r = v as unknown as PresentationMasterXmlChange
  return (
    new TextEncoder().encode(JSON.stringify(r)).length + masterXmlReservedBytes(r) <= 192 * 1024
  )
}
export function validMasterXmlTransition(
  before: PresentationMasterXmlChange | undefined,
  next: PresentationMasterXmlChange,
): boolean {
  if (!validatePresentationMasterXmlChange(next)) return false
  if (!before)
    return (
      next.state === 'prepared' &&
      next.cursor.phase === 'original_probe_stage' &&
      next.cursor.index === 0 &&
      next.cursor.substep === 0 &&
      next.receiptCount === 0 &&
      (next.reviewSequence === undefined || next.reviewSequence === 0) &&
      next.introducedMasterCount === 0 &&
      !next.pending &&
      !next.reviews.length &&
      slots.every((k) => !next[k])
    )
  if (!validatePresentationMasterXmlChange(before)) return false
  if (before.pending && next.reviewSequence !== before.reviewSequence) return false
  const immutable = [
    'version',
    'kind',
    'changeId',
    'documentId',
    'intent',
    'sourceSlideId',
    'packageSourceSlideId',
    'snapshotRef',
    'preparedRef',
    'scope',
    'inventoryCleanupVerified',
  ] as const
  if (
    immutable.some((k) => !same(before[k], next[k])) ||
    slots.some((k) => before[k] && !same(before[k], next[k]))
  )
    return false
  if (!before.pending) {
    if (
      !next.pending &&
      next.cursor.phase === 'verify' &&
      next.cursor.index === 0 &&
      next.cursor.substep === 0 &&
      same({ ...before, cursor: next.cursor, reviews: [] }, { ...next, reviews: [] }) &&
      !next.reviews.length
    )
      return true
    if (next.pending)
      return (
        same({ ...before, reviews: [] }, { ...next, pending: undefined }) &&
        !next.reviews.length &&
        !next.pending.inserted &&
        !next.pending.afterProofRef
      )
    if (before.state === 'prepared' && next.state === 'discarded')
      return same({ ...before, state: 'discarded' }, next)
    return (
      same(
        { ...before, reviews: [], reviewSequence: undefined },
        { ...next, reviews: [], reviewSequence: undefined },
      ) &&
      before.reviews.every((r) => next.reviews.some((n) => n.slideId === r.slideId)) &&
      next.reviews.filter((r) => !before.reviews.some((n) => same(n, r))).length === 1 &&
      next.reviewSequence === (before.reviewSequence ?? 0) + 1
    )
  }
  if (next.pending)
    return (
      same({ ...before, pending: undefined }, { ...next, pending: undefined }) &&
      same(
        {
          ...before.pending,
          inserted: undefined,
          afterProofRef: undefined,
          recoveryPreimageRef: undefined,
        },
        {
          ...next.pending,
          inserted: undefined,
          afterProofRef: undefined,
          recoveryPreimageRef: undefined,
        },
      ) &&
      (!before.pending.recoveryPreimageRef ||
        same(before.pending.recoveryPreimageRef, next.pending.recoveryPreimageRef)) &&
      (!before.pending.inserted || same(before.pending.inserted, next.pending.inserted)) &&
      (!before.pending.afterProofRef ||
        same(before.pending.afterProofRef, next.pending.afterProofRef))
    )
  if (!before.pending.afterProofRef)
    return (
      ['discarded', 'recovery_required'].includes(next.state) &&
      (before.state === 'prepared' || next.state === 'recovery_required') &&
      same({ ...before, pending: undefined, state: next.state }, next)
    )
  if (
    next.receiptCount !== before.receiptCount + 1 ||
    !same(next.currentProofRef, before.pending.afterProofRef) ||
    next.introducedMasterCount < before.introducedMasterCount ||
    next.reviews.length
  )
    return false
  if (
    next.state !== 'recovery_required' &&
    !completedStates[before.pending.action].includes(next.state)
  )
    return false
  if (next.state === 'recovery_required' && !before.pending.targetSlideId) return false
  if (
    !['undone', 'discarded'].includes(next.state) &&
    !allowedActions[next.state].includes(next.cursor.phase)
  )
    return false
  const slot = stageSlots[before.pending.action as keyof typeof stageSlots]
  if (slot && !before[slot] && !same(next[slot], before.pending.inserted)) return false
  if (slots.some((k) => !before[k] && next[k] && k !== slot)) return false
  return true
}
export const validPresentationMasterXmlTransition = validMasterXmlTransition
