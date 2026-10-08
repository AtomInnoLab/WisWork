import type { PackageBackupRef } from './presentation-package-backup.js'
export type PackageAction = 'import' | 'delete_source' | 'restore' | 'delete_applied' | 'discard'
export interface PresentationPackageChange {
  version: 1
  kind: 'package_xml'
  changeId: string
  documentId: string
  intent: string
  sourceKind: 'slide' | 'chart'
  sourceSlideId: string
  packageSourceSlideId: string
  snapshotRef: PackageBackupRef
  originalRef: PackageBackupRef
  preparedRef: PackageBackupRef
  currentProofRef: PackageBackupRef
  state: 'prepared' | 'staged' | 'applied' | 'restore_staged' | 'undone' | 'discarded'
  replacementSlideId?: string
  restoredSlideId?: string
  pending?: {
    action: PackageAction
    beforeProofRef: PackageBackupRef
    insertedSlideId?: string
    afterProofRef?: PackageBackupRef
  }
  receipts: { action: PackageAction; proofRef: PackageBackupRef; slideId?: string }[]
  reviews: { slideId: string; reviewRef: PackageBackupRef }[]
}
const exact = (v: unknown, required: string[], optional: string[] = []): v is Record<string, any> =>
  !!v &&
  typeof v === 'object' &&
  !Array.isArray(v) &&
  required.every((k) => Object.hasOwn(v, k)) &&
  Object.keys(v).every((k) => required.includes(k) || optional.includes(k))
export const validPackageHostId = (v: unknown): v is string =>
  typeof v === 'string' &&
  !!v.trim() &&
  v.length <= 256 &&
  ![...v].some((c) => {
    const code = c.charCodeAt(0)
    return code <= 31 || (code >= 127 && code <= 159)
  })
export const validPackageBackupRef = (v: unknown): v is PackageBackupRef =>
  exact(v, ['key', 'sha256', 'sizeBytes']) &&
  typeof v.key === 'string' &&
  v.key.length <= 128 &&
  /^(snapshot|(?:page|image|receipt)-[0-9]+)$/.test(v.key) &&
  typeof v.sha256 === 'string' &&
  /^[a-f0-9]{64}$/.test(v.sha256) &&
  Number.isSafeInteger(v.sizeBytes) &&
  v.sizeBytes > 0 &&
  v.sizeBytes <= 8 * 1024 * 1024
const actions = ['import', 'delete_source', 'restore', 'delete_applied', 'discard']
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
// Reserve additional space for up to five phase receipts, host IDs, pending proof and historical review.
export function packageChangeReservedBytes(r: PresentationPackageChange): number {
  const max = {
    ...r,
    replacementSlideId: 'x'.repeat(2048),
    restoredSlideId: 'x'.repeat(2048),
    pending: {
      action: 'delete_applied',
      beforeProofRef: r.currentProofRef,
      insertedSlideId: 'x'.repeat(2048),
      afterProofRef: r.currentProofRef,
    },
    receipts: Array.from({ length: 5 }, () => ({
      action: 'delete_applied',
      proofRef: { key: 'receipt-99999999999999999999', sha256: 'f'.repeat(64), sizeBytes: 8388608 },
      slideId: 'x'.repeat(2048),
    })),
    reviews: [
      {
        slideId: 'x'.repeat(2048),
        reviewRef: {
          key: 'receipt-99999999999999999999',
          sha256: 'f'.repeat(64),
          sizeBytes: 8388608,
        },
      },
    ],
  }
  return Math.max(
    0,
    new TextEncoder().encode(JSON.stringify(max)).length -
      new TextEncoder().encode(JSON.stringify(r)).length,
  )
}
export const presentationPackageReservedBytes = packageChangeReservedBytes
export function validatePresentationPackageChange(v: unknown): v is PresentationPackageChange {
  if (
    !exact(
      v,
      [
        'version',
        'kind',
        'changeId',
        'documentId',
        'intent',
        'sourceKind',
        'sourceSlideId',
        'packageSourceSlideId',
        'snapshotRef',
        'originalRef',
        'preparedRef',
        'currentProofRef',
        'state',
        'receipts',
        'reviews',
      ],
      ['replacementSlideId', 'restoredSlideId', 'pending'],
    )
  )
    return false
  if (
    v.version !== 1 ||
    v.kind !== 'package_xml' ||
    typeof v.documentId !== 'string' ||
    !v.documentId.trim() ||
    v.documentId.length > 2048 ||
    !validPackageHostId(v.sourceSlideId) ||
    typeof v.packageSourceSlideId !== 'string' ||
    !/^[1-9][0-9]{0,9}#$/.test(v.packageSourceSlideId) ||
    Number(v.packageSourceSlideId.slice(0, -1)) < 256 ||
    Number(v.packageSourceSlideId.slice(0, -1)) > 4294967295 ||
    typeof v.changeId !== 'string' ||
    !/^[A-Za-z0-9_-]{1,128}$/.test(v.changeId) ||
    typeof v.intent !== 'string' ||
    v.intent.length > 8192 ||
    !['slide', 'chart'].includes(v.sourceKind) ||
    !['prepared', 'staged', 'applied', 'restore_staged', 'undone', 'discarded'].includes(v.state)
  )
    return false
  if (
    ![v.snapshotRef, v.originalRef, v.preparedRef, v.currentProofRef].every(
      validPackageBackupRef,
    ) ||
    v.snapshotRef.key !== 'snapshot' ||
    v.originalRef.key !== 'page-0' ||
    v.preparedRef.key !== 'page-1'
  )
    return false
  if (
    (v.replacementSlideId !== undefined && !validPackageHostId(v.replacementSlideId)) ||
    (v.restoredSlideId !== undefined && !validPackageHostId(v.restoredSlideId))
  )
    return false
  if (['staged', 'applied', 'restore_staged', 'undone'].includes(v.state) && !v.replacementSlideId)
    return false
  if (['restore_staged', 'undone'].includes(v.state) && !v.restoredSlideId) return false
  if (
    !Array.isArray(v.receipts) ||
    v.receipts.length > 5 ||
    !v.receipts.every(
      (r: any) =>
        exact(r, ['action', 'proofRef'], ['slideId']) &&
        actions.includes(r.action) &&
        validPackageBackupRef(r.proofRef) &&
        (r.slideId === undefined || validPackageHostId(r.slideId)),
    ) ||
    new Set(v.receipts.map((r: any) => r.action)).size !== v.receipts.length
  )
    return false
  if (
    !Array.isArray(v.reviews) ||
    !v.reviews.every(
      (r: any) =>
        exact(r, ['slideId', 'reviewRef']) &&
        validPackageHostId(r.slideId) &&
        validPackageBackupRef(r.reviewRef),
    ) ||
    new Set(v.reviews.map((r: any) => r.slideId)).size !== v.reviews.length
  )
    return false
  if (
    v.pending &&
    (!exact(v.pending, ['action', 'beforeProofRef'], ['insertedSlideId', 'afterProofRef']) ||
      !actions.includes(v.pending.action) ||
      !same(v.pending.beforeProofRef, v.currentProofRef) ||
      !validPackageBackupRef(v.pending.beforeProofRef) ||
      (v.pending.afterProofRef !== undefined && !validPackageBackupRef(v.pending.afterProofRef)) ||
      (v.pending.insertedSlideId !== undefined && !validPackageHostId(v.pending.insertedSlideId)))
  )
    return false
  const sequences: Record<string, string[][]> = {
    prepared: [[]],
    staged: [['import']],
    applied: [['import', 'delete_source']],
    restore_staged: [['import', 'delete_source', 'restore']],
    undone: [['import', 'delete_source', 'restore', 'delete_applied']],
    discarded: [[], ['import', 'discard']],
  }
  if (
    !sequences[v.state]!.some((seq) =>
      same(
        seq,
        v.receipts.map((r: any) => r.action),
      ),
    )
  )
    return false
  const latest = v.receipts.at(-1)
  if (latest ? !same(latest.proofRef, v.currentProofRef) : v.currentProofRef.key !== 'receipt-0')
    return false
  const imported = v.receipts.find((r: any) => r.action === 'import'),
    restored = v.receipts.find((r: any) => r.action === 'restore')
  if (imported?.slideId !== v.replacementSlideId || restored?.slideId !== v.restoredSlideId)
    return false
  if (
    v.replacementSlideId === v.sourceSlideId ||
    v.restoredSlideId === v.sourceSlideId ||
    (v.restoredSlideId !== undefined && v.restoredSlideId === v.replacementSlideId)
  )
    return false
  const allowed: Record<string, string[]> = {
    prepared: ['import'],
    staged: ['delete_source', 'discard'],
    applied: ['restore'],
    restore_staged: ['delete_applied'],
    undone: [],
    discarded: [],
  }
  if (
    v.pending &&
    (!allowed[v.state]!.includes(v.pending.action) ||
      (v.pending.insertedSlideId !== undefined &&
        !['import', 'restore'].includes(v.pending.action)) ||
      (v.pending.afterProofRef !== undefined &&
        v.pending.afterProofRef.key !== `receipt-${v.receipts.length + 1}`))
  )
    return false
  if (
    v.pending?.insertedSlideId &&
    [v.sourceSlideId, v.replacementSlideId, v.restoredSlideId].includes(v.pending.insertedSlideId)
  )
    return false
  if (
    v.receipts.some(
      (receipt: any, i: number) =>
        receipt.proofRef.key !== `receipt-${i + 1}` ||
        (['import', 'restore'].includes(receipt.action)
          ? !receipt.slideId
          : receipt.slideId !== undefined),
    )
  )
    return false
  if (v.reviews.length && !['applied', 'undone'].includes(v.state)) return false
  const r = v as unknown as PresentationPackageChange
  return (
    new TextEncoder().encode(JSON.stringify(r)).length + packageChangeReservedBytes(r) <= 192 * 1024
  )
}
export function validPackageTransition(
  before: PresentationPackageChange | undefined,
  next: PresentationPackageChange,
): boolean {
  if (!validatePresentationPackageChange(next)) return false
  if (!before)
    return (
      next.state === 'prepared' &&
      !next.pending &&
      !next.receipts.length &&
      !next.replacementSlideId &&
      !next.restoredSlideId &&
      !next.reviews.length
    )
  if (!validatePresentationPackageChange(before)) return false
  const immutable = [
    'version',
    'kind',
    'changeId',
    'documentId',
    'intent',
    'sourceKind',
    'sourceSlideId',
    'packageSourceSlideId',
    'snapshotRef',
    'originalRef',
    'preparedRef',
  ] as const
  if (immutable.some((k) => !same(before[k], next[k]))) return false
  if (
    (before.replacementSlideId && next.replacementSlideId !== before.replacementSlideId) ||
    (before.restoredSlideId && next.restoredSlideId !== before.restoredSlideId)
  )
    return false
  if (!before.pending) {
    if (
      before.state === 'prepared' &&
      next.state === 'discarded' &&
      same({ ...before, state: 'discarded' }, next)
    )
      return true
    if (next.pending) {
      const expectedMap: Record<string, PackageAction[]> = {
        prepared: ['import'],
        staged: ['delete_source', 'discard'],
        applied: ['restore'],
        restore_staged: ['delete_applied'],
        undone: [],
        discarded: [],
      }
      const expected = expectedMap[before.state]!
      return (
        expected.includes(next.pending.action) &&
        !next.reviews.length &&
        same(
          { ...next, pending: undefined, reviews: before.reviews },
          { ...before, pending: undefined },
        ) &&
        !next.pending.insertedSlideId &&
        !next.pending.afterProofRef
      )
    }
    return (
      same({ ...before, reviews: [] }, { ...next, reviews: [] }) &&
      before.reviews.every((r) => next.reviews.some((n) => n.slideId === r.slideId)) &&
      next.reviews.length <= before.reviews.length + 1 &&
      next.reviews.filter((n) => !before.reviews.some((r) => same(r, n))).length <= 1
    )
  }
  if (next.pending) {
    return (
      next.state === before.state &&
      same({ ...before, pending: undefined }, { ...next, pending: undefined }) &&
      next.pending.action === before.pending.action &&
      same(next.pending.beforeProofRef, before.pending.beforeProofRef) &&
      (!before.pending.insertedSlideId ||
        next.pending.insertedSlideId === before.pending.insertedSlideId) &&
      (!before.pending.afterProofRef ||
        same(next.pending.afterProofRef, before.pending.afterProofRef))
    )
  }
  if (!before.pending.afterProofRef) {
    const closed =
      before.pending.action === 'import' ? next.state === 'discarded' : next.state === before.state
    return (
      closed &&
      same({ ...before, pending: undefined, state: next.state }, { ...next, pending: undefined })
    )
  }
  const states = {
    import: 'staged',
    delete_source: 'applied',
    restore: 'restore_staged',
    delete_applied: 'undone',
    discard: 'discarded',
  } as const
  return (
    next.state === states[before.pending.action] &&
    same(next.currentProofRef, before.pending.afterProofRef) &&
    next.receipts.length === before.receipts.length + 1 &&
    before.receipts.every((r, i) => same(r, next.receipts[i])) &&
    same(next.receipts.at(-1), {
      action: before.pending.action,
      proofRef: before.pending.afterProofRef,
      ...(before.pending.insertedSlideId ? { slideId: before.pending.insertedSlideId } : {}),
    }) &&
    !next.reviews.length
  )
}
export const validPresentationPackageTransition = validPackageTransition
