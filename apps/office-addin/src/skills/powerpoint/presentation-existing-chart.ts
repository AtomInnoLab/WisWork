export interface PresentationExistingChartChange {
  version: 1
  changeId: string
  documentId: string
  oldSlideId: string
  shapeId: string
  slideIndex: number
  beforeSlideIds: string[]
  beforePackageDigest: string
  afterPackageDigest: string
  backup: { backupId: string; sha256: string; sizeBytes: number }
  state: 'pending' | 'write_pending' | 'applied' | 'undo_pending' | 'undone' | 'cancelled'
  newSlideId?: string
  restoredSlideId?: string
  values?: string[][]
  reapplies?: string
  backupReleasedAt?: string
}

const id = (v: unknown): v is string => typeof v === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(v)
const digest = (v: unknown): v is string => typeof v === 'string' && /^[a-f0-9]{64}$/.test(v)
const hostId = (v: unknown): v is string =>
  typeof v === 'string' &&
  v.length > 0 &&
  v.length <= 256 &&
  !Array.from(v).some(
    (c) => c.charCodeAt(0) < 32 || (c.charCodeAt(0) >= 127 && c.charCodeAt(0) <= 159),
  )
const bytes = (v: unknown) => new TextEncoder().encode(JSON.stringify(v)).byteLength

export const existingChartReservedBytes = (r: PresentationExistingChartChange) =>
  Math.max(
    0,
    bytes({
      ...r,
      state: 'write_pending',
      newSlideId: '\uffff'.repeat(256),
      restoredSlideId: '\uffff'.repeat(256),
      backupReleasedAt: '2026-09-24T00:00:00.000Z',
    }) - bytes(r),
  )

export function validatePresentationExistingChartChange(
  value: unknown,
): value is PresentationExistingChartChange {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const r = value as PresentationExistingChartChange
  if (
    Object.keys(r).some(
      (k) =>
        ![
          'version',
          'changeId',
          'documentId',
          'oldSlideId',
          'shapeId',
          'slideIndex',
          'beforeSlideIds',
          'beforePackageDigest',
          'afterPackageDigest',
          'backup',
          'state',
          'newSlideId',
          'restoredSlideId',
          'backupReleasedAt',
          'values',
          'reapplies',
        ].includes(k),
    )
  )
    return false
  if (
    r.version !== 1 ||
    !id(r.changeId) ||
    typeof r.documentId !== 'string' ||
    r.documentId.length < 1 ||
    r.documentId.length > 4096 ||
    !hostId(r.oldSlideId) ||
    !/^[1-9]\d{0,9}$/.test(r.shapeId) ||
    !Number.isSafeInteger(r.slideIndex) ||
    r.slideIndex < 0 ||
    r.slideIndex > 511 ||
    !Array.isArray(r.beforeSlideIds) ||
    r.beforeSlideIds.length < 1 ||
    r.beforeSlideIds.length > 512 ||
    !r.beforeSlideIds.every(hostId) ||
    new Set(r.beforeSlideIds).size !== r.beforeSlideIds.length ||
    r.beforeSlideIds[r.slideIndex] !== r.oldSlideId ||
    !digest(r.beforePackageDigest) ||
    !digest(r.afterPackageDigest) ||
    r.beforePackageDigest === r.afterPackageDigest ||
    !r.backup ||
    typeof r.backup !== 'object' ||
    Array.isArray(r.backup) ||
    Object.keys(r.backup).length !== 3 ||
    !id(r.backup.backupId) ||
    !digest(r.backup.sha256) ||
    !Number.isSafeInteger(r.backup.sizeBytes) ||
    r.backup.sizeBytes < 1 ||
    r.backup.sizeBytes > 100 * 1024 * 1024 ||
    !['pending', 'write_pending', 'applied', 'undo_pending', 'undone', 'cancelled'].includes(
      r.state,
    )
  )
    return false
  if (
    r.values !== undefined &&
    (!Array.isArray(r.values) ||
      r.values.length < 1 ||
      r.values.length > 8 ||
      r.values.some(
        (series) =>
          !Array.isArray(series) ||
          series.length < 1 ||
          series.length > 32 ||
          series.some(
            (value) =>
              typeof value !== 'string' ||
              value.length > 32 ||
              !/^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(value),
          ),
      ))
  )
    return false
  if (r.reapplies !== undefined && (!id(r.reapplies) || r.reapplies === r.changeId || !r.values))
    return false
  if (
    r.backupReleasedAt !== undefined &&
    (!(r.state === 'undone' || r.state === 'cancelled') ||
      typeof r.backupReleasedAt !== 'string' ||
      !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(r.backupReleasedAt) ||
      !Number.isFinite(Date.parse(r.backupReleasedAt)) ||
      new Date(r.backupReleasedAt).toISOString() !== r.backupReleasedAt)
  )
    return false
  if (r.state === 'pending' || r.state === 'write_pending' || r.state === 'cancelled') {
    if (r.newSlideId !== undefined || r.restoredSlideId !== undefined) return false
  } else {
    if (
      !hostId(r.newSlideId) ||
      (r.newSlideId !== r.oldSlideId && r.beforeSlideIds.includes(r.newSlideId))
    )
      return false
    if (r.state === 'undone') {
      if (
        !hostId(r.restoredSlideId) ||
        (r.restoredSlideId !== r.oldSlideId && r.beforeSlideIds.includes(r.restoredSlideId))
      )
        return false
    } else if (r.restoredSlideId !== undefined) return false
  }
  return bytes(r) + existingChartReservedBytes(r) <= 32 * 1024
}

export function validExistingChartTransition(
  before: PresentationExistingChartChange | undefined,
  after: PresentationExistingChartChange,
): boolean {
  if (!validatePresentationExistingChartChange(after)) return false
  if (!before) return after.state === 'pending'
  if (!validatePresentationExistingChartChange(before)) return false
  const core = (r: PresentationExistingChartChange) =>
    JSON.stringify({
      ...r,
      state: undefined,
      newSlideId: undefined,
      restoredSlideId: undefined,
      backupReleasedAt: undefined,
    })
  const next: Record<
    PresentationExistingChartChange['state'],
    PresentationExistingChartChange['state'][]
  > = {
    pending: ['write_pending', 'cancelled'],
    write_pending: ['applied', 'cancelled'],
    applied: ['undo_pending'],
    undo_pending: ['undone'],
    undone: [],
    cancelled: [],
  }
  const release =
    before.state === after.state &&
    (after.state === 'undone' || after.state === 'cancelled') &&
    before.backupReleasedAt === undefined &&
    after.backupReleasedAt !== undefined &&
    before.newSlideId === after.newSlideId &&
    before.restoredSlideId === after.restoredSlideId
  return (
    core(before) === core(after) &&
    (release ||
      (after.backupReleasedAt === undefined && next[before.state].includes(after.state))) &&
    (before.newSlideId === undefined || before.newSlideId === after.newSlideId)
  )
}
