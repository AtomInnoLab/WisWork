import type {
  PowerPointDeclarativeOperation,
  PresentationPageGeometry,
} from './browser-powerpoint-adapter.js'

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

export interface PresentationTargetBatch {
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

export type NativeAddOperation = Extract<
  PowerPointDeclarativeOperation,
  {
    op: 'add_text_box' | 'add_geometric_shape' | 'add_native_table'
  }
>
export interface PresentationNativeAddBatch {
  version: 2
  kind: 'native_page_add'
  changeId: string
  documentId: string
  baselineId: string
  baselineDigest: string
  hostSlideId: string
  slideIndex: number
  beforeSlideIds: string[]
  scope: { slideIds: string[] }
  intent: string
  preserved: string[]
  validation: string[]
  risk: 'high'
  backups: NonNullable<PresentationTargetBatch['backups']>
  operations: NativeAddOperation[]
  createdShapeIds: string[]
  nextIndex: number
  inFlightIndex?: number
  state: 'applying' | 'applied' | 'undoing' | 'undone'
  restoredSlideId?: string
  backupReleasedAt?: string
}
export type NativeModifyOperation = Extract<
  PowerPointDeclarativeOperation,
  { op: 'set_shape_text' | 'set_shape_geometry' | 'delete_shape' }
>
export interface PresentationNativeModifyBatch {
  version: 3
  kind: 'native_page_modify'
  changeId: string
  documentId: string
  baselineId: string
  baselineDigest: string
  beforeSlideIds: string[]
  scope: { slideIds: string[] }
  intent: string
  preserved: string[]
  validation: string[]
  risk: 'high'
  backups: NonNullable<PresentationTargetBatch['backups']>
  operations: NativeModifyOperation[]
  pages: { hostSlideId: string; slideIndex: number; expectedPackageDigest: string }[]
  nextIndex: number
  inFlightIndex?: number
  state: 'applying' | 'applied' | 'undoing' | 'undone'
  restoredSlideIds?: Record<string, string>
  reviews?: PresentationTargetBatch['reviews']
  backupReleasedAt?: string
}
export interface PresentationSlideDuplicationBatch {
  version: 4
  kind: 'native_slide_duplicate'
  changeId: string
  documentId: string
  baselineId: string
  baselineDigest: string
  hostSlideId: string
  slideIndex: number
  sourceSlideId: string
  beforeSlideIds: string[]
  scope: { slideIds: string[] }
  intent: string
  preserved: string[]
  validation: string[]
  risk: 'high'
  backups: NonNullable<PresentationTargetBatch['backups']>
  operations: [{ op: 'duplicate_slide'; slide_index: number }]
  nextIndex: 0 | 1
  inFlightIndex?: 0
  insertedSlideId?: string
  state: 'applying' | 'applied' | 'undoing' | 'undone'
  reviews?: PresentationTargetBatch['reviews']
  backupReleasedAt?: string
}
export type PresentationExistingBatch =
  | PresentationTargetBatch
  | PresentationNativeAddBatch
  | PresentationNativeModifyBatch
  | PresentationSlideDuplicationBatch

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
function validateTargetBatch(v: unknown): v is PresentationTargetBatch {
  if (!v || typeof v !== 'object' || Array.isArray(v)) return false
  const r = v as PresentationTargetBatch
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

const targetBatchReservedBytes = (r: PresentationTargetBatch) => {
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

function validTargetBatchTransition(
  before: PresentationTargetBatch | undefined,
  after: PresentationTargetBatch,
): boolean {
  if (!before) return after.state === 'applying' && after.cursor === 0
  const core = (r: PresentationTargetBatch) =>
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
    const stableCore = (r: PresentationTargetBatch) =>
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
        before.cursor === after.cursor - 1) ||
      (after.state === 'undoing' &&
        after.cursor === before.cursor &&
        before.cursor > 0 &&
        after.reviews === undefined) ||
      (after.state === 'undone' &&
        before.cursor === 0 &&
        after.cursor === 0 &&
        after.reviews === undefined)
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
  if (before.state === 'undone')
    return (
      after.state === 'applying' &&
      after.cursor === 0 &&
      before.cursor === 0 &&
      !!before.backups?.length &&
      before.backupReleasedAt === undefined &&
      after.reviews === undefined
    )
  return false
}

const record = (v: unknown): v is Record<string, unknown> =>
  !!v && typeof v === 'object' && !Array.isArray(v)
const keysOnly = (v: Record<string, unknown>, keys: string[]) =>
  Object.keys(v).every((k) => keys.includes(k))
const boundedText = (v: unknown, max: number) =>
  typeof v === 'string' && v.length > 0 && v.length <= max
const color = (v: unknown) => typeof v === 'string' && /^[A-Fa-f0-9]{6}$/.test(v)
const range = (v: unknown, min: number, max: number): v is number =>
  typeof v === 'number' && Number.isFinite(v) && v >= min && v <= max
const optional = (v: Record<string, unknown>, key: string, valid: (x: unknown) => boolean) =>
  !Object.hasOwn(v, key) || valid(v[key])
// The declarative parser is private to powerpoint-skill; validate the same bounded add subset here.
function validNativeAddOperation(v: unknown, slideIndex: number): v is NativeAddOperation {
  if (
    !record(v) ||
    v.slide_index !== slideIndex ||
    !boundedText(v.name, 256) ||
    !['left', 'top', 'width', 'height'].every((k) => range(v[k], -100000, 100000)) ||
    Number(v.width) <= 0 ||
    Number(v.height) <= 0
  )
    return false
  const common = ['op', 'slide_index', 'name', 'left', 'top', 'width', 'height']
  if (v.op === 'add_geometric_shape')
    return (
      keysOnly(v, [...common, 'shape', 'fill', 'lineColor']) &&
      ['rect', 'ellipse', 'roundRect'].includes(String(v.shape)) &&
      color(v.fill) &&
      color(v.lineColor)
    )
  if (v.op === 'add_text_box')
    return (
      keysOnly(v, [
        ...common,
        'text',
        'fontFace',
        'fontSize',
        'color',
        'bold',
        'align',
        'margin',
        'verticalAlignment',
      ]) &&
      typeof v.text === 'string' &&
      v.text.length <= 12000 &&
      optional(v, 'fontFace', (x) => boundedText(x, 128)) &&
      optional(v, 'fontSize', (x) => range(x, 6, 96)) &&
      optional(v, 'color', color) &&
      optional(v, 'bold', (x) => typeof x === 'boolean') &&
      optional(v, 'align', (x) => ['left', 'center', 'right'].includes(String(x))) &&
      optional(v, 'margin', (x) => range(x, 0, 72)) &&
      optional(v, 'verticalAlignment', (x) => ['top', 'middle', 'bottom'].includes(String(x)))
    )
  if (
    v.op !== 'add_native_table' ||
    !keysOnly(v, [
      ...common,
      'rows',
      'fontFace',
      'fontSize',
      'color',
      'borderColor',
      'cellMargin',
    ]) ||
    !boundedText(v.fontFace, 128) ||
    !range(v.fontSize, 6, 48) ||
    !color(v.color) ||
    !optional(v, 'borderColor', color) ||
    !optional(v, 'cellMargin', (x) => range(x, 0, 36))
  )
    return false
  const rows = v.rows
  return (
    Array.isArray(rows) &&
    rows.length >= 1 &&
    rows.length <= 20 &&
    Array.isArray(rows[0]) &&
    rows[0].length >= 1 &&
    rows[0].length <= 12 &&
    rows.length * rows[0].length <= 128 &&
    Array.from(rows).every(
      (row) =>
        Array.isArray(row) &&
        row.length === rows[0].length &&
        Array.from(row).every((cell) => typeof cell === 'string' && cell.length <= 256),
    ) &&
    JSON.stringify(rows).length <= 12000
  )
}
function validateNativeAddBatch(v: unknown): v is PresentationNativeAddBatch {
  if (
    !record(v) ||
    !keysOnly(v, [
      'version',
      'kind',
      'changeId',
      'documentId',
      'baselineId',
      'baselineDigest',
      'hostSlideId',
      'slideIndex',
      'beforeSlideIds',
      'scope',
      'intent',
      'preserved',
      'validation',
      'risk',
      'backups',
      'operations',
      'createdShapeIds',
      'nextIndex',
      'inFlightIndex',
      'state',
      'restoredSlideId',
      'backupReleasedAt',
    ]) ||
    v.version !== 2 ||
    v.kind !== 'native_page_add' ||
    !id(v.changeId) ||
    !id(v.baselineId) ||
    !boundedText(v.documentId, 4096) ||
    !hostId(v.hostSlideId) ||
    !Number.isSafeInteger(v.slideIndex) ||
    !range(v.slideIndex, 0, 100000) ||
    !uniqueIds(v.beforeSlideIds, 500) ||
    !Array.from(v.beforeSlideIds).every(hostId) ||
    v.beforeSlideIds[v.slideIndex] !== v.hostSlideId ||
    !record(v.scope) ||
    Object.keys(v.scope).join(',') !== 'slideIds' ||
    !Array.isArray(v.scope.slideIds) ||
    v.scope.slideIds.length !== 1 ||
    v.scope.slideIds[0] !== v.hostSlideId ||
    !boundedText(v.intent, 300) ||
    !labels(v.preserved, 20) ||
    !labels(v.validation, 20) ||
    !Array.from(v.preserved as unknown[]).every((x) => boundedText(x, 300)) ||
    !Array.from(v.validation as unknown[]).every((x) => boundedText(x, 300)) ||
    v.risk !== 'high' ||
    !Array.isArray(v.backups) ||
    v.backups.length !== 1 ||
    !record(v.backups[0]) ||
    Object.keys(v.backups[0]).sort().join(',') !==
      'backupId,hostSlideId,packageDigest,sha256,sizeBytes'
  )
    return false
  const backup = v.backups[0]
  const digest = (x: unknown) => typeof x === 'string' && /^[a-f0-9]{64}$/.test(x)
  if (
    backup.hostSlideId !== v.hostSlideId ||
    !id(backup.backupId) ||
    !digest(backup.sha256) ||
    !digest(backup.packageDigest) ||
    backup.packageDigest !== v.baselineDigest ||
    !Number.isSafeInteger(backup.sizeBytes) ||
    !range(backup.sizeBytes, 1, 8 * 1024 * 1024) ||
    !Array.isArray(v.operations) ||
    v.operations.length < 1 ||
    v.operations.length > 32 ||
    !Array.from(v.operations).every((op) => validNativeAddOperation(op, v.slideIndex as number)) ||
    new Set(v.operations.map((op) => op.name)).size !== v.operations.length ||
    !Array.isArray(v.createdShapeIds) ||
    !Number.isSafeInteger(v.nextIndex) ||
    !range(v.nextIndex, 0, v.operations.length) ||
    v.createdShapeIds.length !== v.nextIndex ||
    !Array.from(v.createdShapeIds).every(hostId) ||
    new Set(v.createdShapeIds).size !== v.createdShapeIds.length ||
    !['applying', 'applied', 'undoing', 'undone'].includes(String(v.state))
  )
    return false
  const operationCount = v.operations.length
  if (
    !optional(
      v,
      'inFlightIndex',
      (x) => v.state === 'applying' && x === v.nextIndex && Number(v.nextIndex) < operationCount,
    ) ||
    (v.state === 'applying' && v.nextIndex === v.operations.length) ||
    (v.state === 'applied' && v.nextIndex !== v.operations.length) ||
    !optional(v, 'restoredSlideId', (x) => v.state === 'undone' && hostId(x)) ||
    (v.state === 'undone' && !hostId(v.restoredSlideId)) ||
    !optional(v, 'backupReleasedAt', (x) => v.state === 'undone' && timestamp(x))
  )
    return false
  return bytes(v) + nativeAddReservedBytes(v as unknown as PresentationNativeAddBatch) <= 192 * 1024
}
function nativeAddReservedBytes(r: PresentationNativeAddBatch) {
  // Reserve JSON-escaped lone surrogates (six bytes per code unit), as well as future control fields.
  const maximal = {
    ...r,
    state: 'applying',
    nextIndex: r.operations.length,
    inFlightIndex: 31,
    createdShapeIds: r.operations.map(
      (_, i) => String.fromCharCode(0xd800 + i) + '\ud800'.repeat(255),
    ),
    restoredSlideId: '\ud800'.repeat(256),
    backupReleasedAt: '2026-09-29T00:00:00.000Z',
  }
  return Math.max(0, bytes(maximal) - bytes(r))
}
export function validatePresentationExistingBatch(v: unknown): v is PresentationExistingBatch {
  return (
    record(v) &&
    (v.version === 4
      ? validateSlideDuplicationBatch(v)
      : v.version === 3
        ? validateNativeModifyBatch(v)
        : v.version === 2
          ? validateNativeAddBatch(v)
          : validateTargetBatch(v))
  )
}
export const existingBatchReservedBytes = (r: PresentationExistingBatch) =>
  r.version === 4
    ? slideDuplicationReservedBytes(r)
    : r.version === 3
      ? nativeModifyReservedBytes(r)
      : r.version === 2
        ? nativeAddReservedBytes(r)
        : targetBatchReservedBytes(r)
export function validExistingBatchTransition(
  before: PresentationExistingBatch | undefined,
  after: PresentationExistingBatch,
): boolean {
  if (
    !validatePresentationExistingBatch(after) ||
    (before && !validatePresentationExistingBatch(before))
  )
    return false
  if (after.version === 4) return validSlideDuplicationTransition(before, after)
  if (before?.version === 4) return false
  if (after.version === 3) return validNativeModifyTransition(before, after)
  if (before?.version === 3) return false
  if (after.version === 1)
    return (!before || before.version === 1) && validTargetBatchTransition(before, after)
  if (!before)
    return after.state === 'applying' && after.nextIndex === 0 && after.inFlightIndex === undefined
  if (before.version !== 2) return false
  const core = (r: PresentationNativeAddBatch) =>
    JSON.stringify({
      ...r,
      state: undefined,
      nextIndex: undefined,
      createdShapeIds: undefined,
      inFlightIndex: undefined,
      restoredSlideId: undefined,
      backupReleasedAt: undefined,
    })
  if (
    core(before) !== core(after) ||
    before.createdShapeIds.some((id, i) => after.createdShapeIds[i] !== id) ||
    (before.restoredSlideId !== undefined && before.restoredSlideId !== after.restoredSlideId) ||
    (before.backupReleasedAt !== undefined && before.backupReleasedAt !== after.backupReleasedAt)
  )
    return false
  if (
    before.state === 'applying' &&
    after.state === 'applying' &&
    after.nextIndex === before.nextIndex &&
    JSON.stringify(before.createdShapeIds) === JSON.stringify(after.createdShapeIds)
  )
    return (
      (before.inFlightIndex === undefined && after.inFlightIndex === before.nextIndex) ||
      (before.inFlightIndex === before.nextIndex && after.inFlightIndex === undefined)
    )
  if (
    before.state === 'applying' &&
    ['applying', 'applied'].includes(after.state) &&
    before.inFlightIndex === before.nextIndex &&
    after.inFlightIndex === undefined &&
    after.nextIndex === before.nextIndex + 1
  )
    return true
  if (
    after.nextIndex !== before.nextIndex ||
    JSON.stringify(after.createdShapeIds) !== JSON.stringify(before.createdShapeIds)
  )
    return false
  if (['applying', 'applied'].includes(before.state) && after.state === 'undoing')
    return after.inFlightIndex === undefined
  if (before.state === 'undoing' && after.state === 'undone') return hostId(after.restoredSlideId)
  return (
    before.state === 'undone' &&
    after.state === 'undone' &&
    before.backupReleasedAt === undefined &&
    timestamp(after.backupReleasedAt)
  )
}

function validateNativeModifyBatch(value: Record<string, unknown>): boolean {
  const r = value as unknown as PresentationNativeModifyBatch
  const hash = (v: unknown) => typeof v === 'string' && /^[a-f0-9]{64}$/.test(v)
  const allowed = [
    'version',
    'kind',
    'changeId',
    'documentId',
    'baselineId',
    'baselineDigest',
    'beforeSlideIds',
    'scope',
    'intent',
    'preserved',
    'validation',
    'risk',
    'backups',
    'operations',
    'pages',
    'nextIndex',
    'inFlightIndex',
    'state',
    'restoredSlideIds',
    'backupReleasedAt',
    'reviews',
  ]
  if (
    Object.keys(value).some((k) => !allowed.includes(k)) ||
    r.version !== 3 ||
    r.kind !== 'native_page_modify' ||
    !id(r.changeId) ||
    !id(r.baselineId) ||
    !boundedText(r.documentId, 4096) ||
    !hash(r.baselineDigest) ||
    !uniqueIds(r.beforeSlideIds, 512) ||
    !record(r.scope) ||
    Object.keys(r.scope).join(',') !== 'slideIds' ||
    !uniqueIds(r.scope.slideIds, 8) ||
    !boundedText(r.intent, 300) ||
    !labels(r.preserved, 20) ||
    !labels(r.validation, 20) ||
    r.risk !== 'high' ||
    !Array.isArray(r.pages) ||
    r.pages.length !== r.scope.slideIds.length ||
    !Array.isArray(r.backups) ||
    r.backups.length !== r.pages.length ||
    !Array.isArray(r.operations) ||
    r.operations.length < 1 ||
    r.operations.length > 32 ||
    bytes({ version: 1, operations: r.operations }) > 32768 ||
    !Number.isInteger(r.nextIndex) ||
    r.nextIndex < 0 ||
    r.nextIndex > r.operations.length ||
    !['applying', 'applied', 'undoing', 'undone'].includes(r.state) ||
    (r.state === 'applied' && r.nextIndex !== r.operations.length) ||
    (r.state === 'applying' && r.nextIndex === r.operations.length) ||
    (r.inFlightIndex !== undefined && (r.state !== 'applying' || r.inFlightIndex !== r.nextIndex))
  )
    return false
  if (
    new Set(r.pages.map((p) => p?.hostSlideId)).size !== r.pages.length ||
    new Set(r.backups.map((b) => b?.backupId)).size !== r.backups.length
  )
    return false
  for (const p of r.pages) {
    if (
      !record(p) ||
      Object.keys(p).sort().join(',') !== 'expectedPackageDigest,hostSlideId,slideIndex' ||
      !hostId(p.hostSlideId) ||
      !Number.isInteger(p.slideIndex) ||
      p.slideIndex < 0 ||
      p.slideIndex >= r.beforeSlideIds.length ||
      r.beforeSlideIds[p.slideIndex] !== p.hostSlideId ||
      !r.scope.slideIds.includes(p.hostSlideId) ||
      !hash(p.expectedPackageDigest)
    )
      return false
    const b = r.backups.find((b) => b.hostSlideId === p.hostSlideId)
    if (
      !b ||
      !record(b) ||
      Object.keys(b).sort().join(',') !== 'backupId,hostSlideId,packageDigest,sha256,sizeBytes' ||
      !id(b.backupId) ||
      !hash(b.sha256) ||
      !hash(b.packageDigest) ||
      !Number.isSafeInteger(b.sizeBytes) ||
      b.sizeBytes < 1 ||
      b.sizeBytes > 8 * 1024 * 1024
    )
      return false
  }
  for (const op of r.operations) {
    if (
      !record(op) ||
      !r.pages.some((p) => p.slideIndex === op.slide_index) ||
      !hostId(op.shape_id)
    )
      return false
    if (op.op === 'set_shape_text') {
      if (
        Object.keys(op).sort().join(',') !== 'op,shape_id,slide_index,text' ||
        typeof op.text !== 'string' ||
        op.text.length > 12000
      )
        return false
    } else if (op.op === 'set_shape_geometry') {
      if (
        Object.keys(op).sort().join(',') !== 'height,left,op,shape_id,slide_index,top,width' ||
        ![op.left, op.top, op.width, op.height].every(
          (v) => typeof v === 'number' && Number.isFinite(v),
        ) ||
        op.width <= 0 ||
        op.height <= 0
      )
        return false
    } else if (
      op.op !== 'delete_shape' ||
      Object.keys(op).sort().join(',') !== 'op,shape_id,slide_index'
    )
      return false
  }
  if (
    r.restoredSlideIds !== undefined &&
    (!['undoing', 'undone'].includes(r.state) ||
      !record(r.restoredSlideIds) ||
      Object.entries(r.restoredSlideIds).some(
        ([k, v]) =>
          !r.scope.slideIds.includes(k) ||
          !hostId(v) ||
          r.beforeSlideIds.filter((id) => !r.scope.slideIds.includes(id)).includes(v),
      ) ||
      new Set(Object.values(r.restoredSlideIds)).size !== Object.keys(r.restoredSlideIds).length)
  )
    return false
  const targets = new Map<string, string[]>()
  for (const op of r.operations) {
    const key = JSON.stringify([op.slide_index, op.shape_id])
    const prior = targets.get(key) ?? []
    if (
      prior.includes(op.op) ||
      prior.includes('delete_shape') ||
      (op.op === 'delete_shape' && prior.length)
    )
      return false
    targets.set(key, [...prior, op.op])
  }
  if (r.baselineDigest !== r.backups[0]?.packageDigest) return false
  if (
    r.state === 'undone' &&
    (!r.restoredSlideIds ||
      Object.keys(r.restoredSlideIds).length !== r.pages.length ||
      r.pages.some(
        (p) =>
          !Object.hasOwn(r.restoredSlideIds!, p.hostSlideId) ||
          !hostId(r.restoredSlideIds![p.hostSlideId]),
      ))
  )
    return false
  if (r.backupReleasedAt !== undefined && (r.state !== 'undone' || !timestamp(r.backupReleasedAt)))
    return false
  if (
    r.reviews !== undefined &&
    (!Array.isArray(r.reviews) ||
      r.reviews.length > r.pages.length ||
      new Set(r.reviews.map((review) => review?.hostSlideId)).size !== r.reviews.length ||
      r.reviews.some(
        (review) =>
          !record(review) ||
          Object.keys(review).sort().join(',') !==
            'capturedAt,hostSlideId,notes,reviewedAt,screenshotDigest,status' ||
          !r.scope.slideIds.includes(review.hostSlideId) ||
          !hash(review.screenshotDigest) ||
          !timestamp(review.capturedAt) ||
          !timestamp(review.reviewedAt) ||
          Date.parse(review.capturedAt) > Date.parse(review.reviewedAt) ||
          !['pass', 'fail'].includes(review.status) ||
          typeof review.notes !== 'string' ||
          review.notes.length > 2000,
      ))
  )
    return false
  return bytes(r) + nativeModifyReservedBytes(r) <= 192 * 1024
}
function validNativeModifyTransition(
  before: PresentationExistingBatch | undefined,
  after: PresentationNativeModifyBatch,
) {
  if (!before)
    return (
      after.state === 'applying' &&
      after.reviews === undefined &&
      after.nextIndex === 0 &&
      after.inFlightIndex === undefined &&
      after.restoredSlideIds === undefined &&
      after.backupReleasedAt === undefined &&
      after.pages.every(
        (p) =>
          p.expectedPackageDigest ===
          after.backups.find((b) => b.hostSlideId === p.hostSlideId)?.packageDigest,
      )
    )
  if (before.version !== 3) return false
  const core = (r: PresentationNativeModifyBatch) =>
    JSON.stringify({
      ...r,
      state: undefined,
      nextIndex: undefined,
      inFlightIndex: undefined,
      restoredSlideIds: undefined,
      backupReleasedAt: undefined,
      reviews: undefined,
      pages: r.pages.map((p) => ({ ...p, expectedPackageDigest: undefined })),
    })
  if (core(before) !== core(after)) return false
  const reviewsSame = JSON.stringify(before.reviews) === JSON.stringify(after.reviews)
  if (!reviewsSame) {
    const without = (r: PresentationNativeModifyBatch) =>
      JSON.stringify({ ...r, reviews: undefined })
    if (
      before.state !== 'applied' ||
      after.state !== 'applied' ||
      without(before) !== without(after) ||
      !after.reviews ||
      after.reviews.length < (before.reviews?.length ?? 0)
    )
      return false
    const changed = after.reviews.filter(
      (review) =>
        JSON.stringify(review) !==
        JSON.stringify(before.reviews?.find((prior) => prior.hostSlideId === review.hostSlideId)),
    )
    return (
      changed.length === 1 &&
      (before.reviews ?? []).every((prior) =>
        after.reviews!.some((review) => review.hostSlideId === prior.hostSlideId),
      )
    )
  }
  const pagesSame = JSON.stringify(before.pages) === JSON.stringify(after.pages)
  const restoredSame =
    JSON.stringify(before.restoredSlideIds) === JSON.stringify(after.restoredSlideIds)
  if (
    before.state === 'applying' &&
    after.state === 'applying' &&
    before.nextIndex === after.nextIndex &&
    pagesSame &&
    restoredSame &&
    before.backupReleasedAt === after.backupReleasedAt
  )
    return before.inFlightIndex === undefined && after.inFlightIndex === before.nextIndex
  if (
    before.state === 'applying' &&
    ['applying', 'applied'].includes(after.state) &&
    before.inFlightIndex === before.nextIndex &&
    after.inFlightIndex === undefined &&
    after.nextIndex === before.nextIndex + 1 &&
    restoredSame &&
    before.backupReleasedAt === after.backupReleasedAt
  )
    return before.pages.every(
      (p, i) =>
        p.slideIndex === before.operations[before.nextIndex].slide_index ||
        JSON.stringify(p) === JSON.stringify(after.pages[i]),
    )
  if (before.nextIndex !== after.nextIndex || !pagesSame) return false
  if (
    ['applying', 'applied', 'undoing'].includes(before.state) &&
    ['undoing', 'undone'].includes(after.state) &&
    (after.state !== 'undone' || before.state === 'undoing') &&
    after.inFlightIndex === undefined &&
    before.backupReleasedAt === after.backupReleasedAt
  )
    return Object.entries(before.restoredSlideIds ?? {}).every(
      ([k, v]) => after.restoredSlideIds?.[k] === v,
    )
  return (
    before.state === 'undone' &&
    after.state === 'undone' &&
    restoredSame &&
    before.backupReleasedAt === undefined &&
    timestamp(after.backupReleasedAt)
  )
}

function nativeModifyReservedBytes(r: PresentationNativeModifyBatch) {
  return (
    bytes({
      restoredSlideIds: Object.fromEntries(
        r.scope.slideIds.map((id) => [id, '\u0001'.repeat(256)]),
      ),
      backupReleasedAt: '9999-12-31T23:59:59.999Z',
      reviews: r.scope.slideIds.map((hostSlideId) => ({
        hostSlideId,
        screenshotDigest: 'f'.repeat(64),
        capturedAt: '9999-12-31T23:59:59.999Z',
        reviewedAt: '9999-12-31T23:59:59.999Z',
        status: 'pass',
        notes: '\u0001'.repeat(2000),
      })),
    }) +
    256 -
    bytes(r.reviews ?? []) -
    bytes(r.restoredSlideIds ?? {})
  )
}

function validateSlideDuplicationBatch(value: Record<string, unknown>): boolean {
  const r = value as unknown as PresentationSlideDuplicationBatch
  const hash = (value: unknown) => typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
  const allowed = [
    'version',
    'kind',
    'changeId',
    'documentId',
    'baselineId',
    'baselineDigest',
    'hostSlideId',
    'slideIndex',
    'sourceSlideId',
    'beforeSlideIds',
    'scope',
    'intent',
    'preserved',
    'validation',
    'risk',
    'backups',
    'operations',
    'nextIndex',
    'inFlightIndex',
    'insertedSlideId',
    'state',
    'reviews',
    'backupReleasedAt',
  ]
  if (
    Object.keys(value).some((key) => !allowed.includes(key)) ||
    r.version !== 4 ||
    r.kind !== 'native_slide_duplicate' ||
    !id(r.changeId) ||
    !id(r.baselineId) ||
    !boundedText(r.documentId, 4096) ||
    !hash(r.baselineDigest) ||
    !hostId(r.hostSlideId) ||
    !Number.isInteger(r.slideIndex) ||
    !uniqueIds(r.beforeSlideIds, 512) ||
    r.slideIndex < 0 ||
    r.beforeSlideIds[r.slideIndex] !== r.hostSlideId ||
    typeof r.sourceSlideId !== 'string' ||
    !/^([1-9][0-9]{0,9})#$/.test(r.sourceSlideId) ||
    Number(r.sourceSlideId.slice(0, -1)) < 256 ||
    Number(r.sourceSlideId.slice(0, -1)) > 4294967295 ||
    !record(r.scope) ||
    Object.keys(r.scope).join(',') !== 'slideIds' ||
    !Array.isArray(r.scope.slideIds) ||
    r.scope.slideIds.length !== 1 ||
    r.scope.slideIds[0] !== r.hostSlideId ||
    !boundedText(r.intent, 300) ||
    !labels(r.preserved, 20) ||
    !labels(r.validation, 20) ||
    r.risk !== 'high' ||
    !Array.isArray(r.backups) ||
    r.backups.length !== 1 ||
    !Array.isArray(r.operations) ||
    r.operations.length !== 1 ||
    !record(r.operations[0]) ||
    Object.keys(r.operations[0]).sort().join(',') !== 'op,slide_index' ||
    r.operations[0].op !== 'duplicate_slide' ||
    r.operations[0].slide_index !== r.slideIndex ||
    ![0, 1].includes(r.nextIndex) ||
    !['applying', 'applied', 'undoing', 'undone'].includes(r.state)
  )
    return false
  const b = r.backups[0]
  if (
    !record(b) ||
    Object.keys(b).sort().join(',') !== 'backupId,hostSlideId,packageDigest,sha256,sizeBytes' ||
    b.hostSlideId !== r.hostSlideId ||
    !id(b.backupId) ||
    !hash(b.sha256) ||
    b.packageDigest !== r.baselineDigest ||
    !Number.isSafeInteger(b.sizeBytes) ||
    b.sizeBytes < 1 ||
    b.sizeBytes > 8 * 1024 * 1024
  )
    return false
  if (
    r.inFlightIndex !== undefined &&
    (r.inFlightIndex !== 0 || r.state !== 'applying' || r.nextIndex !== 0)
  )
    return false
  if (
    r.insertedSlideId !== undefined &&
    (!hostId(r.insertedSlideId) || r.beforeSlideIds.includes(r.insertedSlideId))
  )
    return false
  if (
    r.state === 'applying' &&
    (r.nextIndex !== 0 || (r.insertedSlideId !== undefined && r.inFlightIndex !== 0))
  )
    return false
  if (
    r.state !== 'applying' &&
    !(r.state === 'undone' && r.nextIndex === 0 && r.insertedSlideId === undefined) &&
    (r.nextIndex !== 1 || r.inFlightIndex !== undefined || !hostId(r.insertedSlideId))
  )
    return false
  if (r.backupReleasedAt !== undefined && (r.state !== 'undone' || !timestamp(r.backupReleasedAt)))
    return false
  if (
    r.reviews !== undefined &&
    (!Array.isArray(r.reviews) ||
      r.reviews.length !== 1 ||
      r.reviews.some(
        (review) =>
          !record(review) ||
          Object.keys(review).sort().join(',') !==
            'capturedAt,hostSlideId,notes,reviewedAt,screenshotDigest,status' ||
          review.hostSlideId !== r.insertedSlideId ||
          !hash(review.screenshotDigest) ||
          !timestamp(review.capturedAt) ||
          !timestamp(review.reviewedAt) ||
          Date.parse(review.capturedAt) > Date.parse(review.reviewedAt) ||
          !['pass', 'fail'].includes(review.status) ||
          typeof review.notes !== 'string' ||
          review.notes.length > 2000,
      ))
  )
    return false
  return bytes(r) + slideDuplicationReservedBytes(r) <= 192 * 1024
}
function slideDuplicationReservedBytes(r: PresentationSlideDuplicationBatch) {
  const maximal = {
    ...r,
    state: 'applying',
    nextIndex: 1,
    inFlightIndex: 0,
    insertedSlideId: '\ud800'.repeat(256),
    backupReleasedAt: '9999-12-31T23:59:59.999Z',
    reviews: [
      {
        hostSlideId: '\ud800'.repeat(256),
        screenshotDigest: 'f'.repeat(64),
        capturedAt: '9999-12-31T23:59:59.999Z',
        reviewedAt: '9999-12-31T23:59:59.999Z',
        status: 'pass',
        notes: '\u0001'.repeat(2000),
      },
    ],
  }
  return Math.max(0, bytes(maximal) - bytes(r))
}
function validSlideDuplicationTransition(
  before: PresentationExistingBatch | undefined,
  after: PresentationSlideDuplicationBatch,
) {
  if (!before)
    return (
      after.state === 'applying' &&
      after.nextIndex === 0 &&
      after.inFlightIndex === undefined &&
      after.insertedSlideId === undefined &&
      after.reviews === undefined &&
      after.backupReleasedAt === undefined
    )
  if (before.version !== 4) return false
  const core = (r: PresentationSlideDuplicationBatch) =>
    JSON.stringify({
      ...r,
      state: undefined,
      nextIndex: undefined,
      inFlightIndex: undefined,
      insertedSlideId: undefined,
      reviews: undefined,
      backupReleasedAt: undefined,
    })
  if (
    core(before) !== core(after) ||
    (before.insertedSlideId !== undefined && before.insertedSlideId !== after.insertedSlideId)
  )
    return false
  const reviewsSame = JSON.stringify(before.reviews) === JSON.stringify(after.reviews)
  if (!reviewsSame)
    return (
      before.state === 'applied' &&
      after.state === 'applied' &&
      before.nextIndex === after.nextIndex &&
      before.insertedSlideId === after.insertedSlideId &&
      before.inFlightIndex === after.inFlightIndex &&
      before.backupReleasedAt === after.backupReleasedAt &&
      after.reviews?.length === 1
    )
  if (before.backupReleasedAt !== after.backupReleasedAt)
    return (
      before.state === 'undone' &&
      after.state === 'undone' &&
      before.backupReleasedAt === undefined &&
      timestamp(after.backupReleasedAt) &&
      before.nextIndex === after.nextIndex &&
      before.insertedSlideId === after.insertedSlideId
    )
  if (
    before.state === 'applying' &&
    (before.inFlightIndex === 0 || before.inFlightIndex === undefined) &&
    before.nextIndex === 0 &&
    before.insertedSlideId === undefined &&
    after.state === 'undone' &&
    after.nextIndex === 0 &&
    after.insertedSlideId === undefined &&
    after.inFlightIndex === undefined
  )
    return true
  if (
    before.state === 'applying' &&
    after.state === 'applying' &&
    before.nextIndex === after.nextIndex
  ) {
    if (before.inFlightIndex === undefined && after.inFlightIndex === 0)
      return before.insertedSlideId === undefined && after.insertedSlideId === undefined
    return (
      before.inFlightIndex === 0 &&
      after.inFlightIndex === 0 &&
      before.insertedSlideId === undefined &&
      hostId(after.insertedSlideId)
    )
  }
  if (
    before.state === 'applying' &&
    before.inFlightIndex === 0 &&
    after.state === 'applied' &&
    after.nextIndex === 1 &&
    after.inFlightIndex === undefined
  )
    return hostId(after.insertedSlideId)
  if (
    before.nextIndex !== after.nextIndex ||
    before.insertedSlideId !== after.insertedSlideId ||
    before.inFlightIndex !== after.inFlightIndex
  )
    return false
  return (
    (before.state === 'applied' && after.state === 'undoing') ||
    (before.state === 'undoing' && after.state === 'undone')
  )
}
