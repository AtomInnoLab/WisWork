export const presentationAcquisitionErrors = [
  'remote_image_unavailable',
  'remote_webpage_unavailable',
  'quota_exceeded',
  'parse_failed',
  'digest_mismatch',
  'animated_image_unsupported',
  'remote_image_source_conflict',
  'remote_webpage_source_conflict',
  'aborted',
  'invalid_state',
  'acquisition_failed',
] as const
export type PresentationAcquisitionError = (typeof presentationAcquisitionErrors)[number]
export interface PresentationAcquisitionInput {
  kind: 'image' | 'webpage'
  source: string
  sourceUrlHash: string
}
export type PresentationAcquisitionResult =
  | {
      state: 'ready'
      attachmentId: string
      sha256: string
      sizeBytes: number
      assetSha256?: string
    }
  | { state: 'rejected'; error: PresentationAcquisitionError; attachmentId?: string }
export type PresentationAcquisitionRecord = PresentationAcquisitionInput & {
  id: string
  attempt: number
  startedAt: string
} & ({ state: 'fetching' } | (PresentationAcquisitionResult & { finishedAt: string }))
export interface PresentationAcquisitionHistory {
  version: 1
  scope: 'remote_material_acquisition'
  documentId: string
  revision: number
  totalAttempts: number
  records: PresentationAcquisitionRecord[]
}
export function parsePresentationAcquisitionHistory(
  value: unknown,
): PresentationAcquisitionHistory {
  const fail = (): never => {
    throw new Error('invalid_state')
  }
  const object = (v: unknown): v is Record<string, unknown> =>
    !!v && typeof v === 'object' && !Array.isArray(v)
  const exact = (v: object, keys: string[]) =>
    Object.keys(v).sort().join(',') === keys.sort().join(',')
  const integer = (v: unknown) => Number.isSafeInteger(v) && Number(v) >= 0
  const digest = (v: unknown) => typeof v === 'string' && /^[a-f0-9]{64}$/.test(v)
  const id = (v: unknown) => typeof v === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(v)
  const time = (v: unknown) =>
    typeof v === 'string' &&
    /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(v) &&
    Number.isFinite(Date.parse(v)) &&
    new Date(v).toISOString() === v
  if (
    !object(value) ||
    !exact(value, ['version', 'scope', 'documentId', 'revision', 'totalAttempts', 'records'])
  )
    fail()
  const h = value as unknown as PresentationAcquisitionHistory
  if (
    h.version !== 1 ||
    h.scope !== 'remote_material_acquisition' ||
    typeof h.documentId !== 'string' ||
    !h.documentId.trim() ||
    h.documentId.length > 2048 ||
    !integer(h.revision) ||
    !integer(h.totalAttempts) ||
    h.revision < h.totalAttempts ||
    h.revision > h.totalAttempts * 2 ||
    !Array.isArray(h.records) ||
    h.records.length !== Math.min(64, h.totalAttempts)
  )
    fail()
  const ids = new Set<string>()
  let previous = ''
  let completed = 0
  for (const [i, r] of h.records.entries()) {
    if (
      !object(r) ||
      !id(r.id) ||
      ids.has(r.id) ||
      r.attempt !== h.totalAttempts - h.records.length + i + 1 ||
      !['image', 'webpage'].includes(r.kind) ||
      !digest(r.sourceUrlHash) ||
      !time(r.startedAt) ||
      r.startedAt < previous
    )
      fail()
    ids.add(r.id)
    previous = r.startedAt
    try {
      const url = new URL(r.source)
      if (
        !['http:', 'https:'].includes(url.protocol) ||
        url.username ||
        url.password ||
        url.search ||
        url.hash ||
        url.toString() !== r.source ||
        r.source.length > 2048
      )
        fail()
    } catch {
      fail()
    }
    const keys = ['id', 'attempt', 'kind', 'source', 'sourceUrlHash', 'state', 'startedAt']
    if (r.state === 'ready') {
      keys.push('finishedAt', 'attachmentId', 'sha256', 'sizeBytes')
      if (r.kind === 'image') keys.push('assetSha256')
      if (
        !digest(r.attachmentId) ||
        r.sha256 !== r.attachmentId ||
        !integer(r.sizeBytes) ||
        r.sizeBytes > (r.kind === 'image' ? 10 : 5) * 1024 * 1024 ||
        (r.kind === 'image' && !digest(r.assetSha256))
      )
        fail()
    } else if (r.state === 'rejected') {
      keys.push('finishedAt', 'error')
      if (r.attachmentId !== undefined) {
        keys.push('attachmentId')
        if (!digest(r.attachmentId)) fail()
      }
      if (!presentationAcquisitionErrors.includes(r.error)) fail()
    } else if (r.state !== 'fetching') fail()
    if (r.state !== 'fetching') {
      completed++
      if (!time(r.finishedAt) || r.finishedAt < r.startedAt) fail()
    }
    if (!exact(r, keys)) fail()
  }
  if (
    h.revision < h.totalAttempts + completed ||
    h.revision > h.totalAttempts + completed + (h.totalAttempts - h.records.length)
  )
    fail()
  return structuredClone(h)
}
