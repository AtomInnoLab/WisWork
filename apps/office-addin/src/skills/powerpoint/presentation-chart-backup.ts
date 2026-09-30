import { MAX_PPTX_PACKAGE_BYTES, presentationPackageDigest } from './powerpoint-package.js'

type Request = (body: unknown, signal?: AbortSignal) => Promise<Response>
type Scope = { request: Request; documentId: string; hostSlideId: string; slideIds: string[] }
export type ChartPackageBackup = { backupId: string; sha256: string; sizeBytes: number }
type Save = Scope & { base64: string; backupId: string }
type Read = Scope & { backup: ChartPackageBackup; expectedPackageDigest?: string }
type Release = Scope & { backup: ChartPackageBackup }
const CHUNK = 128 * 1024
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const encode = (value: Uint8Array) =>
  btoa(Array.from(value, (x) => String.fromCharCode(x)).join(''))
const decode = (value: string) => Uint8Array.from(atob(value), (x) => x.charCodeAt(0))
const sha = async (value: Uint8Array) =>
  Array.from(
    new Uint8Array(await crypto.subtle.digest('SHA-256', Uint8Array.from(value).buffer)),
    (x) => x.toString(16).padStart(2, '0'),
  ).join('')
const fail = (): never => {
  throw new Error('presentation_chart_backup_invalid')
}

export async function describePagePackageBackup(base64: string, signal?: AbortSignal) {
  const bytes = validBase64(base64)
  return {
    packageDigest: await presentationPackageDigest(base64, signal),
    sha256: await sha(bytes),
    sizeBytes: bytes.length,
  }
}

function validScope(value: Scope, backupId: string) {
  if (
    !/^[A-Za-z0-9_-]{1,128}$/.test(backupId) ||
    !value.documentId ||
    !value.hostSlideId ||
    !value.slideIds.length ||
    value.slideIds.some((id) => !id)
  )
    fail()
}
function validBase64(value: string) {
  if (
    !value ||
    value.length > Math.ceil(MAX_PPTX_PACKAGE_BYTES / 3) * 4 ||
    !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(value)
  )
    fail()
  let result: Uint8Array
  try {
    result = decode(value)
  } catch {
    return fail()
  }
  if (!result.length || result.length > MAX_PPTX_PACKAGE_BYTES || encode(result) !== value) fail()
  return result
}
async function call(
  scope: Scope,
  operation: string,
  data: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<Record<string, unknown>> {
  if (signal?.aborted) throw new Error('cancelled')
  const response = await scope.request({ operation, documentId: scope.documentId, ...data }, signal)
  if (signal?.aborted) throw new Error('cancelled')
  let body: unknown
  try {
    body = await response.json()
  } catch {
    fail()
  }
  if (
    operation === 'existing_page_backup_begin' &&
    body &&
    typeof body === 'object' &&
    !Array.isArray(body) &&
    'error' in body &&
    body.error === 'quota_exceeded'
  )
    throw new Error('presentation_existing_backup_capacity')
  if (!response.ok || !body || typeof body !== 'object' || Array.isArray(body) || 'error' in body)
    fail()
  return body as Record<string, unknown>
}
function match(scope: Scope, backup: ChartPackageBackup, value: Record<string, unknown>) {
  return (
    value.backupId === backup.backupId &&
    value.documentId === scope.documentId &&
    value.hostSlideId === scope.hostSlideId &&
    same(value.slideIds, scope.slideIds) &&
    value.sha256 === backup.sha256 &&
    value.sizeBytes === backup.sizeBytes
  )
}

export async function saveChartPackageBackup(
  input: Save,
  signal?: AbortSignal,
): Promise<ChartPackageBackup> {
  validScope(input, input.backupId)
  const bytes = validBase64(input.base64)
  const backup = { backupId: input.backupId, sha256: await sha(bytes), sizeBytes: bytes.length }
  const data = {
    backupId: backup.backupId,
    hostSlideId: input.hostSlideId,
    slideIds: input.slideIds,
    sha256: backup.sha256,
    sizeBytes: backup.sizeBytes,
  }
  let meta = await call(input, 'existing_page_backup_begin', data, signal)
  if (!match(input, backup, meta)) fail()
  while (meta.status !== 'ready' && Number(meta.receivedBytes) < bytes.length) {
    const offset = Number(meta.receivedBytes)
    if (!Number.isSafeInteger(offset) || offset < 0) fail()
    const chunk = bytes.subarray(offset, Math.min(offset + CHUNK, bytes.length))
    meta = await call(
      input,
      'existing_page_backup_chunk',
      { backupId: backup.backupId, offset, base64: encode(chunk) },
      signal,
    )
    if (!match(input, backup, meta) || meta.receivedBytes !== offset + chunk.length) fail()
  }
  if (meta.status !== 'ready')
    meta = await call(input, 'existing_page_backup_finish', { backupId: backup.backupId }, signal)
  if (!match(input, backup, meta) || meta.status !== 'ready' || meta.receivedBytes !== bytes.length)
    fail()
  return backup
}

export async function readChartPackageBackup(input: Read, signal?: AbortSignal): Promise<string> {
  const { backup } = input
  validScope(input, backup.backupId)
  if (
    !/^[a-f0-9]{64}$/.test(backup.sha256) ||
    !Number.isSafeInteger(backup.sizeBytes) ||
    backup.sizeBytes <= 0 ||
    backup.sizeBytes > MAX_PPTX_PACKAGE_BYTES
  )
    fail()
  const meta = await call(
    input,
    'existing_page_backup_status',
    { backupId: backup.backupId },
    signal,
  )
  if (
    !match(input, backup, meta) ||
    meta.status !== 'ready' ||
    meta.receivedBytes !== backup.sizeBytes
  )
    fail()
  const bytes = new Uint8Array(backup.sizeBytes)
  for (let offset = 0; offset < bytes.length; offset += CHUNK) {
    const length = Math.min(CHUNK, bytes.length - offset)
    const part = await call(
      input,
      'existing_page_backup_read',
      { backupId: backup.backupId, offset, length },
      signal,
    )
    if (
      part.backupId !== backup.backupId ||
      part.offset !== offset ||
      part.sizeBytes !== backup.sizeBytes ||
      part.sha256 !== backup.sha256 ||
      typeof part.base64 !== 'string'
    )
      fail()
    const chunk = validBase64(part.base64 as string)
    if (chunk.length !== length) fail()
    bytes.set(chunk, offset)
  }
  if ((await sha(bytes)) !== backup.sha256) fail()
  const base64 = encode(bytes)
  if (
    input.expectedPackageDigest &&
    (await presentationPackageDigest(base64, signal)) !== input.expectedPackageDigest
  )
    fail()
  return base64
}

/** Release only a backup whose durable change intent is known to be absent. */
export async function releaseChartPackageBackup(input: Release): Promise<void> {
  validScope(input, input.backup.backupId)
  const receipt = await call(input, 'existing_page_backup_release', {
    backupId: input.backup.backupId,
    hostSlideId: input.hostSlideId,
    slideIds: input.slideIds,
    sha256: input.backup.sha256,
    sizeBytes: input.backup.sizeBytes,
  })
  if (
    Object.keys(receipt).sort().join(',') !==
      'backupId,documentId,hostSlideId,sha256,sizeBytes,slideIds,status' ||
    receipt.status !== 'released' ||
    !match(input, input.backup, receipt)
  )
    fail()
}
