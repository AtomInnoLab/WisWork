import type {
  PresentationPageGeometry,
  PresentationTextRangeSnapshot,
} from './browser-powerpoint-adapter.js'

interface ExistingChangeBase {
  version: 1
  changeId: string
  documentId: string
  baselineId: string
  baselineDigest: string
  scope: { slideIds: string[]; shapeIds?: string[] }
  hostSlideId: string
  shapeId: string
  shapeType: string
  state: 'pending' | 'applied' | 'undo_pending' | 'undone'
  beforeSlideIds?: string[]
  backup?: {
    hostSlideId: string
    backupId: string
    sha256: string
    sizeBytes: number
    packageDigest: string
  }
  backupReleasedAt?: string
  review?: {
    screenshotDigest: string
    capturedAt: string
    reviewedAt: string
    status: 'pass' | 'fail'
    notes: string
  }
}
export type PresentationExistingChange = ExistingChangeBase &
  (
    | { kind: 'text'; before: string; after: string }
    | {
        kind: 'text_range'
        start: number
        length: number
        font: PresentationTextRangeSnapshot['font']
        /** Present on run-verified records; absent on older records. */
        runStructureDigest?: string
        before: string
        after: string
      }
    | {
        kind: 'table_cell'
        rowIndex: number
        columnIndex: number
        cellStructureDigest: string
        before: string
        after: string
      }
    | { kind: 'geometry'; before: PresentationPageGeometry; after: PresentationPageGeometry }
  )
const bytes = (value: unknown) => new TextEncoder().encode(JSON.stringify(value)).byteLength
const hostId = (value: unknown): value is string =>
  typeof value === 'string' &&
  value.length > 0 &&
  value.length <= 256 &&
  !Array.from(value).some((char) => {
    const code = char.charCodeAt(0)
    return code < 32 || (code >= 127 && code <= 159)
  })
const id = (value: unknown) => typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
const digest = (value: unknown) => typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
const timestamp = (s: unknown): s is string =>
  typeof s === 'string' &&
  s.length <= 40 &&
  Number.isFinite(Date.parse(s)) &&
  new Date(s).toISOString() === s
const ids = (value: unknown): value is string[] =>
  Array.isArray(value) &&
  value.length > 0 &&
  value.length <= 500 &&
  value.every(hostId) &&
  new Set(value).size === value.length
function geometry(value: unknown): boolean {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const g = value as PresentationPageGeometry
  return (
    Object.keys(g).length === 4 &&
    (['left', 'top', 'width', 'height'] as const).every(
      (key) => typeof g[key] === 'number' && Number.isFinite(g[key]) && Math.abs(g[key]) <= 100000,
    ) &&
    g.width >= 0 &&
    g.height >= 0
  )
}
export function validatePresentationExistingChange(
  value: unknown,
): value is PresentationExistingChange {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const r = value as PresentationExistingChange
  if (
    !Object.keys(r).every((key) =>
      [
        'version',
        'changeId',
        'documentId',
        'baselineId',
        'baselineDigest',
        'scope',
        'hostSlideId',
        'shapeId',
        'shapeType',
        'kind',
        'rowIndex',
        'columnIndex',
        'cellStructureDigest',
        'start',
        'length',
        'font',
        'runStructureDigest',
        'before',
        'after',
        'state',
        'beforeSlideIds',
        'backup',
        'backupReleasedAt',
        'review',
      ].includes(key),
    ) ||
    r.version !== 1 ||
    !id(r.changeId) ||
    !id(r.baselineId) ||
    !digest(r.baselineDigest) ||
    typeof r.documentId !== 'string' ||
    !r.documentId.length ||
    r.documentId.length > 4096 ||
    !hostId(r.hostSlideId) ||
    !hostId(r.shapeId) ||
    !hostId(r.shapeType) ||
    !['pending', 'applied', 'undo_pending', 'undone'].includes(r.state) ||
    !r.scope ||
    typeof r.scope !== 'object' ||
    Array.isArray(r.scope) ||
    Object.keys(r.scope).some((key) => !['slideIds', 'shapeIds'].includes(key)) ||
    !ids(r.scope.slideIds) ||
    !r.scope.slideIds.includes(r.hostSlideId) ||
    (r.scope.shapeIds !== undefined &&
      (!ids(r.scope.shapeIds) || !r.scope.shapeIds.includes(r.shapeId)))
  )
    return false
  if (
    r.backupReleasedAt !== undefined &&
    (!r.backup || r.state !== 'undone' || !timestamp(r.backupReleasedAt))
  )
    return false
  if (
    (r.beforeSlideIds !== undefined) !== (r.backup !== undefined) ||
    (r.beforeSlideIds !== undefined && !ids(r.beforeSlideIds)) ||
    (r.backup !== undefined &&
      (!r.beforeSlideIds ||
        !r.beforeSlideIds.includes(r.hostSlideId) ||
        !r.backup ||
        typeof r.backup !== 'object' ||
        Array.isArray(r.backup) ||
        Object.keys(r.backup).sort().join(',') !==
          'backupId,hostSlideId,packageDigest,sha256,sizeBytes' ||
        r.backup.hostSlideId !== r.hostSlideId ||
        !id(r.backup.backupId) ||
        !digest(r.backup.sha256) ||
        !digest(r.backup.packageDigest) ||
        !Number.isSafeInteger(r.backup.sizeBytes) ||
        r.backup.sizeBytes < 1 ||
        r.backup.sizeBytes > 8 * 1024 * 1024))
  )
    return false
  const cell = r.kind === 'table_cell'
  if (
    cell &&
    (r.shapeType !== 'Table' ||
      !digest(r.cellStructureDigest) ||
      !Number.isSafeInteger(r.rowIndex) ||
      !Number.isSafeInteger(r.columnIndex) ||
      r.rowIndex < 0 ||
      r.rowIndex > 1000 ||
      r.columnIndex < 0 ||
      r.columnIndex > 1000)
  )
    return false
  if (
    r.kind === 'text' || r.kind === 'text_range' || cell
      ? typeof r.before !== 'string' ||
        r.before.length > 12000 ||
        typeof r.after !== 'string' ||
        r.after.length > (cell ? 128 : 12000)
      : r.kind !== 'geometry' || !geometry(r.before) || !geometry(r.after)
  )
    return false
  if (!cell && ('rowIndex' in r || 'columnIndex' in r || 'cellStructureDigest' in r)) return false
  if (r.kind === 'text_range') {
    if (
      (r.runStructureDigest !== undefined && !digest(r.runStructureDigest)) ||
      !Number.isSafeInteger(r.start) ||
      r.start < 0 ||
      !Number.isSafeInteger(r.length) ||
      r.length < 1 ||
      r.length > 128 ||
      r.before.length !== r.after.length ||
      r.start + r.length > r.before.length ||
      /[\uD800-\uDFFF]/u.test(r.before) ||
      /[\uD800-\uDFFF]/u.test(r.after) ||
      /[\uD800-\uDFFF]/u.test(r.before.slice(r.start, r.start + r.length)) ||
      /[\uD800-\uDFFF]/u.test(r.after.slice(r.start, r.start + r.length)) ||
      /[\r\n]/.test(r.before.slice(r.start, r.start + r.length)) ||
      /[\r\n]/.test(r.after.slice(r.start, r.start + r.length)) ||
      r.before.slice(0, r.start) !== r.after.slice(0, r.start) ||
      r.before.slice(r.start + r.length) !== r.after.slice(r.start + r.length) ||
      r.before.slice(r.start, r.start + r.length) === r.after.slice(r.start, r.start + r.length) ||
      !r.font ||
      typeof r.font !== 'object' ||
      Array.isArray(r.font) ||
      Object.keys(r.font).sort().join(',') !== 'bold,color,italic,name,size,underline' ||
      typeof r.font.name !== 'string' ||
      r.font.name.length > 256 ||
      typeof r.font.size !== 'number' ||
      !Number.isFinite(r.font.size) ||
      r.font.size < 0 ||
      typeof r.font.color !== 'string' ||
      r.font.color.length > 256 ||
      typeof r.font.bold !== 'boolean' ||
      typeof r.font.italic !== 'boolean' ||
      typeof r.font.underline !== 'string' ||
      r.font.underline.length > 64
    )
      return false
  } else if ('start' in r || 'length' in r || 'font' in r || 'runStructureDigest' in r) return false
  if (r.review !== undefined) {
    const v = r.review
    if (
      !['applied', 'undone'].includes(r.state) ||
      !v ||
      typeof v !== 'object' ||
      Array.isArray(v) ||
      Object.keys(v).length !== 5 ||
      !digest(v.screenshotDigest) ||
      !timestamp(v.capturedAt) ||
      !timestamp(v.reviewedAt) ||
      v.reviewedAt < v.capturedAt ||
      !['pass', 'fail'].includes(v.status) ||
      typeof v.notes !== 'string' ||
      bytes(v) > 8192
    )
      return false
  }
  return bytes(r) + existingChangeReservedBytes(r) <= 192 * 1024
}
// Reserve both the terminal review and longest recovery state before host mutation.
export const existingChangeReservedBytes = (record: PresentationExistingChange) =>
  'undo_pending'.length -
  record.state.length +
  8192 +
  ',"review":'.length -
  (record.review === undefined ? 0 : bytes(record.review) + ',"review":'.length) +
  (record.backup && !record.backupReleasedAt ? 68 : 0)
