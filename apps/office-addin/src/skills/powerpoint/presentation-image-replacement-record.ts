export interface ImageReplacementRecord {
  version: 1
  documentId: string
  projectId: string
  requestId: string
  pageId: string
  hostSlideId: string
  oldShapeId: string
  assetDigest: string
  state: 'pending' | 'complete'
  newShapeId?: string
}
const id = (value: unknown, max: number) =>
  typeof value === 'string' && new RegExp(`^[A-Za-z0-9_-]{1,${max}}$`).test(value)
const hostId = (value: unknown) =>
  typeof value === 'string' &&
  value.length > 0 &&
  value.length <= 256 &&
  !Array.from(value).some((c) => c.charCodeAt(0) < 32)
// Reserve the maximum escaped newShapeId and the longer completed-state label before insertion.
export function imageReplacementReservedBytes(record: ImageReplacementRecord): number {
  return record.state === 'pending' ? (record.newShapeId ? 1 : 1600) : 0
}
export function validateImageReplacementRecord(value: unknown): value is ImageReplacementRecord {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const r = value as ImageReplacementRecord
  return (
    Object.keys(r).every((key) =>
      [
        'version',
        'documentId',
        'projectId',
        'requestId',
        'pageId',
        'hostSlideId',
        'oldShapeId',
        'assetDigest',
        'state',
        'newShapeId',
      ].includes(key),
    ) &&
    r.version === 1 &&
    typeof r.documentId === 'string' &&
    r.documentId.length > 0 &&
    r.documentId.length <= 4096 &&
    id(r.projectId, 80) &&
    id(r.requestId, 128) &&
    id(r.pageId, 80) &&
    hostId(r.hostSlideId) &&
    hostId(r.oldShapeId) &&
    typeof r.assetDigest === 'string' &&
    /^[a-f0-9]{64}$/.test(r.assetDigest) &&
    ['pending', 'complete'].includes(r.state) &&
    (r.newShapeId === undefined || (hostId(r.newShapeId) && r.newShapeId !== r.oldShapeId)) &&
    (r.state !== 'complete' || r.newShapeId !== undefined) &&
    new TextEncoder().encode(JSON.stringify(r)).byteLength + imageReplacementReservedBytes(r) <=
      16 * 1024
  )
}
export async function imageReplacementKey(
  projectId: string,
  requestId: string,
  pageId: string,
  oldShapeId: string,
): Promise<string> {
  if (!id(projectId, 80) || !id(requestId, 128) || !id(pageId, 80) || !hostId(oldShapeId))
    throw new Error('presentation_image_replacement_state_invalid')
  const bytes = await crypto.subtle.digest(
    'SHA-256',
    new TextEncoder().encode(JSON.stringify([projectId, requestId, pageId, oldShapeId])),
  )
  return Array.from(new Uint8Array(bytes), (byte) => byte.toString(16).padStart(2, '0')).join('')
}
