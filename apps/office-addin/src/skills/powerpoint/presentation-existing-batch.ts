import type { PresentationPageGeometry } from './browser-powerpoint-adapter.js'

export type ExistingBatchOperation = {
  hostSlideId: string
  shapeId: string
  shapeType: string
} & (
  | { kind: 'text'; before: string; after: string }
  | {
      kind: 'table_cell'
      rowIndex: number
      columnIndex: number
      tableStructureDigest: string
      before: string
      after: string
    }
  | { kind: 'geometry'; before: PresentationPageGeometry; after: PresentationPageGeometry }
)

export interface PresentationExistingBatch {
  version: 1
  changeId: string
  documentId: string
  baselineId: string
  baselineDigest: string
  beforeSlideIds?: string[]
  backups?: {
    hostSlideId: string
    backupId: string
    sha256: string
    sizeBytes: number
    packageDigest: string
  }[]
  backupReleasedAt?: string
  scope: { slideIds: string[]; shapeIds?: string[] }
  intent: string
  preserved: string[]
  validation: string[]
  /** Captured baseline fields of non-target shapes on affected pages. */
  preservedPageDigests?: Record<string, string>
  /** Captured fields of target shapes outside the planned operation kinds. */
  preservedTargetDigests?: Record<string, string>
  risk: 'medium' | 'high'
  operations: ExistingBatchOperation[]
  state: 'applying' | 'applied' | 'undoing' | 'undone'
  /** Number of operations with durable host readback, in forward order. */
  cursor: number
  /** New records reserve room for one historical review per affected page. */
  reviewCapacity?: true
  reviews?: {
    hostSlideId: string
    screenshotDigest: string
    capturedAt: string
    reviewedAt: string
    status: 'pass' | 'fail'
    notes: string
  }[]
}

const bytes = (v: unknown) => new TextEncoder().encode(JSON.stringify(v)).byteLength
const id = (v: unknown) => typeof v === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(v)
const hostId = (v: unknown): v is string =>
  typeof v === 'string' &&
  v.length > 0 &&
  v.length <= 256 &&
  !Array.from(v).some(
    (c) => c.charCodeAt(0) < 32 || (c.charCodeAt(0) >= 127 && c.charCodeAt(0) <= 159),
  )
const uniqueIds = (v: unknown, max: number): v is string[] =>
  Array.isArray(v) &&
  v.length > 0 &&
  v.length <= max &&
  v.every(hostId) &&
  new Set(v).size === v.length
const labels = (v: unknown, max: number): v is string[] =>
  Array.isArray(v) &&
  v.length <= max &&
  v.every((x) => typeof x === 'string' && x.length > 0 && x.length <= 300)
const geometry = (v: unknown): v is PresentationPageGeometry => {
  if (!v || typeof v !== 'object' || Array.isArray(v)) return false
  const g = v as PresentationPageGeometry
  return (
    Object.keys(g).length === 4 &&
    (['left', 'top', 'width', 'height'] as const).every(
      (k) => typeof g[k] === 'number' && Number.isFinite(g[k]) && Math.abs(g[k]) <= 100000,
    ) &&
    g.width >= 0 &&
    g.height >= 0
  )
}
export function validatePresentationExistingBatch(v: unknown): v is PresentationExistingBatch {
  if (!v || typeof v !== 'object' || Array.isArray(v)) return false
  const r = v as PresentationExistingBatch
  if (
    Object.keys(r).some(
      (k) =>
        ![
          'version',
          'changeId',
          'documentId',
          'baselineId',
          'baselineDigest',
          'beforeSlideIds',
          'backups',
          'backupReleasedAt',
          'scope',
          'intent',
          'preserved',
          'validation',
          'preservedPageDigests',
          'preservedTargetDigests',
          'risk',
          'operations',
          'state',
          'cursor',
          'reviewCapacity',
          'reviews',
        ].includes(k),
    ) ||
    r.version !== 1 ||
    !id(r.changeId) ||
    !id(r.baselineId) ||
    typeof r.documentId !== 'string' ||
    !r.documentId.length ||
    r.documentId.length > 4096 ||
    typeof r.baselineDigest !== 'string' ||
    !/^[a-f0-9]{64}$/.test(r.baselineDigest) ||
    (r.beforeSlideIds !== undefined && !uniqueIds(r.beforeSlideIds, 500)) ||
    (r.backups !== undefined &&
      (!r.beforeSlideIds ||
        !Array.isArray(r.backups) ||
        r.backups.length > 8 ||
        new Set(r.backups.map((backup) => backup?.hostSlideId)).size !== r.backups.length ||
        new Set(r.backups.map((backup) => backup?.backupId)).size !== r.backups.length ||
        r.backups.some(
          (backup) =>
            !backup ||
            typeof backup !== 'object' ||
            Array.isArray(backup) ||
            Object.keys(backup).sort().join(',') !==
              'backupId,hostSlideId,packageDigest,sha256,sizeBytes' ||
            !hostId(backup.hostSlideId) ||
            !id(backup.backupId) ||
            typeof backup.sha256 !== 'string' ||
            !/^[a-f0-9]{64}$/.test(backup.sha256) ||
            typeof backup.packageDigest !== 'string' ||
            !/^[a-f0-9]{64}$/.test(backup.packageDigest) ||
            !Number.isSafeInteger(backup.sizeBytes) ||
            backup.sizeBytes < 1 ||
            backup.sizeBytes > 8 * 1024 * 1024 ||
            !r.beforeSlideIds!.includes(backup.hostSlideId),
        ))) ||
    (r.backupReleasedAt !== undefined &&
      (!r.backups || r.state !== 'undone' || !timestamp(r.backupReleasedAt))) ||
    !r.scope ||
    typeof r.scope !== 'object' ||
    Array.isArray(r.scope) ||
    Object.keys(r.scope).some((k) => !['slideIds', 'shapeIds'].includes(k)) ||
    !uniqueIds(r.scope.slideIds, 20) ||
    (r.scope.shapeIds !== undefined && !uniqueIds(r.scope.shapeIds, 100)) ||
    typeof r.intent !== 'string' ||
    !r.intent.length ||
    r.intent.length > 300 ||
    !labels(r.preserved, 20) ||
    !labels(r.validation, 20) ||
    (r.preservedPageDigests !== undefined &&
      (typeof r.preservedPageDigests !== 'object' ||
        Array.isArray(r.preservedPageDigests) ||
        Object.keys(r.preservedPageDigests).length > 20 ||
        Object.entries(r.preservedPageDigests).some(
          ([slideId, value]) =>
            !r.scope.slideIds.includes(slideId) ||
            typeof value !== 'string' ||
            !/^[a-f0-9]{64}$/.test(value),
        ))) ||
    (r.preservedTargetDigests !== undefined &&
      (typeof r.preservedTargetDigests !== 'object' ||
        Array.isArray(r.preservedTargetDigests) ||
        Object.keys(r.preservedTargetDigests).length > 8 ||
        Object.entries(r.preservedTargetDigests).some(
          ([, value]) => typeof value !== 'string' || !/^[a-f0-9]{64}$/.test(value),
        ))) ||
    !['medium', 'high'].includes(r.risk) ||
    !Array.isArray(r.operations) ||
    r.operations.length < 2 ||
    r.operations.length > 8 ||
    !['applying', 'applied', 'undoing', 'undone'].includes(r.state) ||
    !Number.isInteger(r.cursor) ||
    (r.reviewCapacity !== undefined && r.reviewCapacity !== true) ||
    r.cursor < 0 ||
    r.cursor > r.operations.length ||
    (r.state === 'applied' && r.cursor !== r.operations.length) ||
    (r.state === 'undone' && r.cursor !== 0) ||
    (r.state === 'applying' && r.cursor === r.operations.length) ||
    (r.state === 'undoing' && r.cursor === 0)
  )
    return false
  const keys = new Set<string>()
  for (const op of r.operations) {
    if (
      !op ||
      typeof op !== 'object' ||
      Array.isArray(op) ||
      Object.keys(op).some(
        (k) =>
          ![
            'hostSlideId',
            'shapeId',
            'shapeType',
            'kind',
            'rowIndex',
            'columnIndex',
            'tableStructureDigest',
            'before',
            'after',
          ].includes(k),
      ) ||
      !hostId(op.hostSlideId) ||
      !hostId(op.shapeId) ||
      !hostId(op.shapeType) ||
      !r.scope.slideIds.includes(op.hostSlideId) ||
      (r.scope.shapeIds !== undefined && !r.scope.shapeIds.includes(op.shapeId)) ||
      (op.kind === 'text' || op.kind === 'table_cell'
        ? typeof op.before !== 'string' ||
          op.before.length > (op.kind === 'table_cell' ? 128 : 12000) ||
          typeof op.after !== 'string' ||
          op.after.length > (op.kind === 'table_cell' ? 128 : 12000) ||
          op.before === op.after
        : op.kind !== 'geometry' ||
          !geometry(op.before) ||
          !geometry(op.after) ||
          JSON.stringify(op.before) === JSON.stringify(op.after))
    )
      return false
    if (op.kind === 'table_cell') {
      if (
        op.shapeType !== 'Table' ||
        !Number.isSafeInteger(op.rowIndex) ||
        !Number.isSafeInteger(op.columnIndex) ||
        op.rowIndex < 0 ||
        op.rowIndex > 19 ||
        op.columnIndex < 0 ||
        op.columnIndex > 11 ||
        !/^[a-f0-9]{64}$/.test(op.tableStructureDigest)
      )
        return false
    } else if ('rowIndex' in op || 'columnIndex' in op || 'tableStructureDigest' in op) return false
    const key = JSON.stringify([
      op.hostSlideId,
      op.shapeId,
      op.kind,
      ...(op.kind === 'table_cell' ? [op.rowIndex, op.columnIndex] : []),
    ])
    if (keys.has(key)) return false
    keys.add(key)
  }
  const tableOps = r.operations.filter(
    (op): op is Extract<ExistingBatchOperation, { kind: 'table_cell' }> => op.kind === 'table_cell',
  )
  if (r.backups) {
    const affected = new Set(r.operations.map((op) => op.hostSlideId))
    if (
      r.backups.length !== affected.size ||
      r.backups.some((backup) => !affected.has(backup.hostSlideId))
    )
      return false
  }
  if (r.preservedPageDigests) {
    const affected = new Set(r.operations.map((op) => op.hostSlideId))
    if (
      Object.keys(r.preservedPageDigests).length !== affected.size ||
      ![...affected].every((slideId) => Object.hasOwn(r.preservedPageDigests!, slideId))
    )
      return false
  }
  if (r.preservedTargetDigests) {
    const targets = new Set(r.operations.map((op) => JSON.stringify([op.hostSlideId, op.shapeId])))
    if (
      Object.keys(r.preservedTargetDigests).length !== targets.size ||
      ![...targets].every((key) => Object.hasOwn(r.preservedTargetDigests!, key))
    )
      return false
  }
  if (
    tableOps.length &&
    (tableOps.length !== r.operations.length ||
      tableOps.some(
        (op) =>
          op.hostSlideId !== tableOps[0]!.hostSlideId ||
          op.shapeId !== tableOps[0]!.shapeId ||
          op.tableStructureDigest !== tableOps[0]!.tableStructureDigest,
      ))
  )
    return false
  if (r.reviews !== undefined) {
    const affected = new Set(r.operations.map((op) => op.hostSlideId))
    if (
      !['applied', 'undone'].includes(r.state) ||
      !Array.isArray(r.reviews) ||
      r.reviews.length > affected.size ||
      new Set(r.reviews.map((v) => v?.hostSlideId)).size !== r.reviews.length ||
      r.reviews.some(
        (v) =>
          !v ||
          typeof v !== 'object' ||
          Array.isArray(v) ||
          Object.keys(v).length !== 6 ||
          !affected.has(v.hostSlideId) ||
          typeof v.screenshotDigest !== 'string' ||
          !/^[a-f0-9]{64}$/.test(v.screenshotDigest) ||
          !timestamp(v.capturedAt) ||
          !timestamp(v.reviewedAt) ||
          v.reviewedAt < v.capturedAt ||
          !['pass', 'fail'].includes(v.status) ||
          typeof v.notes !== 'string' ||
          bytes(v) > 8192,
      )
    )
      return false
  }
  return bytes(r) + existingBatchReservedBytes(r) <= 192 * 1024
}

const timestamp = (v: unknown): v is string =>
  typeof v === 'string' &&
  v.length <= 40 &&
  Number.isFinite(Date.parse(v)) &&
  new Date(v).toISOString() === v

export const existingBatchReservedBytes = (r: PresentationExistingBatch) => {
  const stateBytes =
    Math.max('applying'.length, 'applied'.length, 'undoing'.length, 'undone'.length) -
    r.state.length +
    1
  const releaseBytes = r.backups && !r.backupReleasedAt ? 68 : 0
  if (!r.reviewCapacity) return stateBytes + releaseBytes
  const pages = new Set(r.operations.map((op) => op.hostSlideId)).size
  const reviews = r.reviews ?? []
  return (
    stateBytes +
    releaseBytes +
    (r.reviews === undefined ? ',"reviews":[]'.length : 0) +
    pages * (8192 + 2) -
    reviews.reduce((sum, review) => sum + bytes(review), 0)
  )
}

export function validExistingBatchTransition(
  before: PresentationExistingBatch | undefined,
  after: PresentationExistingBatch,
): boolean {
  if (!before) return after.state === 'applying' && after.cursor === 0
  const core = (r: PresentationExistingBatch) =>
    JSON.stringify({
      ...r,
      state: undefined,
      cursor: undefined,
      reviews: undefined,
      backupReleasedAt: undefined,
    })
  if (core(before) !== core(after)) {
    if (
      before.state !== 'applying' ||
      after.state !== 'applying' ||
      before.cursor !== 0 ||
      after.cursor !== 0 ||
      !before.backups ||
      !after.backups ||
      before.backups.length !== after.backups.length ||
      before.backupReleasedAt !== undefined ||
      after.backupReleasedAt !== undefined ||
      JSON.stringify(before.reviews) !== JSON.stringify(after.reviews)
    )
      return false
    const stableCore = (r: PresentationExistingBatch) =>
      JSON.stringify({
        ...r,
        backups: r.backups?.map(({ hostSlideId, packageDigest }) => ({
          hostSlideId,
          packageDigest,
        })),
      })
    return (
      stableCore(before) === stableCore(after) &&
      before.backups.filter((backup, index) => {
        const next = after.backups![index]!
        return (
          backup.backupId !== next.backupId &&
          (backup.sha256 !== next.sha256 || backup.sizeBytes !== next.sizeBytes)
        )
      }).length === 1 &&
      before.backups.filter(
        (backup, index) => JSON.stringify(backup) !== JSON.stringify(after.backups![index]),
      ).length === 1
    )
  }
  if (before.backupReleasedAt !== after.backupReleasedAt)
    return (
      before.state === 'undone' &&
      after.state === 'undone' &&
      before.cursor === 0 &&
      after.cursor === 0 &&
      before.backupReleasedAt === undefined &&
      timestamp(after.backupReleasedAt) &&
      JSON.stringify(before.reviews) === JSON.stringify(after.reviews)
    )
  if (before.state !== after.state && after.reviews !== undefined) return false
  if (
    before.state === after.state &&
    before.cursor === after.cursor &&
    ['applied', 'undone'].includes(before.state)
  )
    return (before.reviews ?? []).every((review) =>
      after.reviews?.some((next) => next.hostSlideId === review.hostSlideId),
    )
  if (before.state === 'applying')
    return (
      (after.state === 'applying' &&
        after.cursor === before.cursor + 1 &&
        after.cursor < after.operations.length) ||
      (after.state === 'applied' &&
        after.cursor === after.operations.length &&
        before.cursor === after.cursor - 1)
    )
  if (before.state === 'applied')
    return (
      after.state === 'undoing' && after.cursor === before.cursor && after.reviews === undefined
    )
  if (before.state === 'undoing')
    return (
      (after.state === 'undoing' && after.cursor === before.cursor - 1 && after.cursor > 0) ||
      (after.state === 'undone' && after.cursor === 0 && before.cursor === 1)
    )
  return false
}
