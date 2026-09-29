export const presentationDeliveryBundleFiles = [
  'presentation.pptx',
  'evidence.json',
  'evidence.md',
  'claims.json',
  'sources.json',
  'quality.json',
  'checkpoints.json',
  'README.md',
] as const
export interface PresentationDeliveryBundleManifest {
  version: 1
  scope: 'current_office_document'
  documentId: string
  projectId: string
  requestId: string
  planRevision: number
  inputDigest: string
  planDigest: string
  createdAt: string
  files: { name: string; sizeBytes: number; sha256: string }[]
  checks: {
    completion: 'not_verified'
    sourceAuthority: 'not_verified'
    timeliness: 'not_verified'
    roundTrip: 'not_run'
    hostQa: 'not_checked' | 'historical_records_only'
    pdf: 'included' | 'not_requested' | 'unavailable'
  }
}
export interface PresentationDeliveryBundleReceipt {
  version: 1
  documentId: string
  projectId: string
  requestId: string
  bundleId: string
  sha256: string
  sizeBytes: number
  receivedBytes: number
  state: 'uploading' | 'ready'
  createdAt: string
  completedAt?: string
  manifest: PresentationDeliveryBundleManifest
}
const fail = (): never => {
  throw new Error('invalid_state')
}
const object = (v: unknown): v is Record<string, unknown> =>
  !!v && typeof v === 'object' && !Array.isArray(v)
const exact = (v: unknown, keys: string[]) =>
  object(v) && Object.keys(v).sort().join(',') === keys.sort().join(',')
const id = (v: unknown) => typeof v === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(v)
const hash = (v: unknown) => typeof v === 'string' && /^[a-f0-9]{64}$/.test(v)
const integer = (v: unknown, min: number, max: number) =>
  typeof v === 'number' && Number.isSafeInteger(v) && v >= min && v <= max
const time = (v: unknown) =>
  typeof v === 'string' &&
  /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(v) &&
  Number.isFinite(Date.parse(v)) &&
  new Date(v).toISOString() === v
const document = (v: unknown) => typeof v === 'string' && !!v.trim() && v.length <= 4096
export function parsePresentationDeliveryBundleManifest(
  value: unknown,
): PresentationDeliveryBundleManifest {
  if (
    !exact(value, [
      'version',
      'scope',
      'documentId',
      'projectId',
      'requestId',
      'planRevision',
      'inputDigest',
      'planDigest',
      'createdAt',
      'files',
      'checks',
    ])
  )
    fail()
  const m = value as PresentationDeliveryBundleManifest
  if (
    m.version !== 1 ||
    m.scope !== 'current_office_document' ||
    !document(m.documentId) ||
    !id(m.projectId) ||
    !id(m.requestId) ||
    !integer(m.planRevision, 1, Number.MAX_SAFE_INTEGER) ||
    !hash(m.inputDigest) ||
    !hash(m.planDigest) ||
    !time(m.createdAt) ||
    !Array.isArray(m.files) ||
    ![8, 9, 10, 11].includes(m.files.length)
  )
    fail()
  const names = new Set<string>()
  let total = 0
  for (const file of m.files) {
    if (
      !exact(file, ['name', 'sizeBytes', 'sha256']) ||
      typeof file.name !== 'string' ||
      ![
        ...presentationDeliveryBundleFiles,
        'presentation.pdf',
        'research.json',
        'research.md',
      ].includes(file.name as (typeof presentationDeliveryBundleFiles)[number]) ||
      names.has(file.name) ||
      !integer(
        file.sizeBytes,
        1,
        file.name === 'presentation.pdf' ? 10 * 1024 * 1024 : 20 * 1024 * 1024,
      ) ||
      !hash(file.sha256)
    )
      fail()
    names.add(file.name)
    total += file.sizeBytes
  }
  if (
    total > 32 * 1024 * 1024 ||
    presentationDeliveryBundleFiles.some((name) => !names.has(name)) ||
    !exact(m.checks, [
      'completion',
      'sourceAuthority',
      'timeliness',
      'roundTrip',
      'hostQa',
      'pdf',
    ]) ||
    m.checks.completion !== 'not_verified' ||
    m.checks.sourceAuthority !== 'not_verified' ||
    m.checks.timeliness !== 'not_verified' ||
    m.checks.roundTrip !== 'not_run' ||
    !['not_checked', 'historical_records_only'].includes(m.checks.hostQa) ||
    !['included', 'not_requested', 'unavailable'].includes(m.checks.pdf) ||
    (m.checks.pdf === 'included') !== names.has('presentation.pdf') ||
    names.has('research.json') !== names.has('research.md')
  )
    fail()
  return structuredClone(m)
}
export function parsePresentationDeliveryBundleReceipt(
  value: unknown,
): PresentationDeliveryBundleReceipt {
  if (
    !object(value) ||
    !exact(value, [
      'version',
      'documentId',
      'projectId',
      'requestId',
      'bundleId',
      'sha256',
      'sizeBytes',
      'receivedBytes',
      'state',
      'createdAt',
      'manifest',
      ...(value.state === 'ready' ? ['completedAt'] : []),
    ])
  )
    fail()
  const r = value as unknown as PresentationDeliveryBundleReceipt
  const m = parsePresentationDeliveryBundleManifest(r.manifest)
  if (
    r.version !== 1 ||
    r.documentId !== m.documentId ||
    r.projectId !== m.projectId ||
    r.requestId !== m.requestId ||
    !hash(r.bundleId) ||
    r.sha256 !== r.bundleId ||
    !integer(r.sizeBytes, 1, 20 * 1024 * 1024) ||
    !integer(r.receivedBytes, 0, r.sizeBytes) ||
    !time(r.createdAt) ||
    !['uploading', 'ready'].includes(r.state) ||
    (r.state === 'ready' &&
      (!time(r.completedAt) || r.completedAt! < r.createdAt || r.receivedBytes !== r.sizeBytes))
  )
    fail()
  return structuredClone(r)
}
