import type { PresentationPageGeometry } from './browser-powerpoint-adapter.js'

export type ExistingBatchOperation = {
  hostSlideId: string
  shapeId: string
  shapeType: string
} & (
  | { kind: 'text'; before: string; after: string }
  | { kind: 'geometry'; before: PresentationPageGeometry; after: PresentationPageGeometry }
)

export interface PresentationExistingBatch {
  version: 1
  changeId: string
  documentId: string
  baselineId: string
  baselineDigest: string
  scope: { slideIds: string[]; shapeIds?: string[] }
  intent: string
  preserved: string[]
  validation: string[]
  risk: 'medium' | 'high'
  operations: ExistingBatchOperation[]
  state: 'applying' | 'applied' | 'undoing' | 'undone'
  /** Number of operations with durable host readback, in forward order. */
  cursor: number
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
          'scope',
          'intent',
          'preserved',
          'validation',
          'risk',
          'operations',
          'state',
          'cursor',
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
    !['medium', 'high'].includes(r.risk) ||
    !Array.isArray(r.operations) ||
    r.operations.length < 2 ||
    r.operations.length > 8 ||
    !['applying', 'applied', 'undoing', 'undone'].includes(r.state) ||
    !Number.isInteger(r.cursor) ||
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
        (k) => !['hostSlideId', 'shapeId', 'shapeType', 'kind', 'before', 'after'].includes(k),
      ) ||
      !hostId(op.hostSlideId) ||
      !hostId(op.shapeId) ||
      !hostId(op.shapeType) ||
      !r.scope.slideIds.includes(op.hostSlideId) ||
      (r.scope.shapeIds !== undefined && !r.scope.shapeIds.includes(op.shapeId)) ||
      (op.kind === 'text'
        ? typeof op.before !== 'string' ||
          op.before.length > 12000 ||
          typeof op.after !== 'string' ||
          op.after.length > 12000 ||
          op.before === op.after
        : op.kind !== 'geometry' ||
          !geometry(op.before) ||
          !geometry(op.after) ||
          JSON.stringify(op.before) === JSON.stringify(op.after))
    )
      return false
    const key = JSON.stringify([op.hostSlideId, op.shapeId, op.kind])
    if (keys.has(key)) return false
    keys.add(key)
  }
  return bytes(r) + existingBatchReservedBytes(r) <= 192 * 1024
}

export const existingBatchReservedBytes = (r: PresentationExistingBatch) =>
  Math.max('applying'.length, 'applied'.length, 'undoing'.length, 'undone'.length) -
  r.state.length +
  1

export function validExistingBatchTransition(
  before: PresentationExistingBatch | undefined,
  after: PresentationExistingBatch,
): boolean {
  if (!before) return after.state === 'applying' && after.cursor === 0
  const core = (r: PresentationExistingBatch) =>
    JSON.stringify({ ...r, state: undefined, cursor: undefined })
  if (core(before) !== core(after)) return false
  if (before.state === 'applying')
    return (
      (after.state === 'applying' &&
        after.cursor === before.cursor + 1 &&
        after.cursor < after.operations.length) ||
      (after.state === 'applied' &&
        after.cursor === after.operations.length &&
        before.cursor === after.cursor - 1)
    )
  if (before.state === 'applied') return after.state === 'undoing' && after.cursor === before.cursor
  if (before.state === 'undoing')
    return (
      (after.state === 'undoing' && after.cursor === before.cursor - 1 && after.cursor > 0) ||
      (after.state === 'undone' && after.cursor === 0 && before.cursor === 1)
    )
  return false
}
