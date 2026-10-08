export interface PresentationTextChange {
  version: 1
  changeId: string
  documentId: string
  projectId: string
  requestId: string
  source?: 'production'
  artifactDigest: string
  pageId: string
  hostSlideId: string
  shapeId: string
  before: string
  after: string
  state: 'pending' | 'applied' | 'undo_pending' | 'undone'
}
const id = (value: unknown, max: number) =>
  typeof value === 'string' && new RegExp(`^[A-Za-z0-9_-]{1,${max}}$`).test(value)
const hostId = (value: unknown) =>
  typeof value === 'string' &&
  value.length > 0 &&
  value.length <= 256 &&
  !Array.from(value).some(
    (char) => char.charCodeAt(0) < 32 || (char.charCodeAt(0) >= 127 && char.charCodeAt(0) <= 159),
  )
export function validatePresentationTextChange(value: unknown): value is PresentationTextChange {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const r = value as PresentationTextChange
  return (
    Object.keys(r).every((key) =>
      [
        'version',
        'changeId',
        'documentId',
        'projectId',
        'requestId',
        'source',
        'artifactDigest',
        'pageId',
        'hostSlideId',
        'shapeId',
        'before',
        'after',
        'state',
      ].includes(key),
    ) &&
    r.version === 1 &&
    id(r.changeId, 128) &&
    typeof r.documentId === 'string' &&
    r.documentId.length > 0 &&
    r.documentId.length <= 4096 &&
    id(r.projectId, 80) &&
    id(r.requestId, 128) &&
    id(r.pageId, 80) &&
    (r.source === undefined || r.source === 'production') &&
    typeof r.artifactDigest === 'string' &&
    /^[a-f0-9]{64}$/.test(r.artifactDigest) &&
    hostId(r.hostSlideId) &&
    hostId(r.shapeId) &&
    typeof r.before === 'string' &&
    r.before.length <= 12000 &&
    typeof r.after === 'string' &&
    r.after.length <= 12000 &&
    ['pending', 'applied', 'undo_pending', 'undone'].includes(r.state) &&
    new TextEncoder().encode(JSON.stringify(r)).byteLength <= 192 * 1024
  )
}
