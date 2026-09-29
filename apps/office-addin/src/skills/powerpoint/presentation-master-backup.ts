type Request = (body: unknown, signal?: AbortSignal) => Promise<Response>
type Scope = {
  request: Request
  documentId: string
  changeId: string
  signal?: AbortSignal
  protocol?: 'master' | 'package'
}
export type MasterBackupRef = { key: string; sha256: string; sizeBytes: number }
type Save = Scope & { key: string; bytes: Uint8Array }
type Read = Scope & { backup: MasterBackupRef }
const MAX_BYTES = 8 * 1024 * 1024
const CHUNK = 128 * 1024
const MAX_JSON = 256 * 1024
function invalid(): never {
  throw new Error('presentation_master_backup_invalid')
}
const cancelled = (signal?: AbortSignal) => {
  if (signal?.aborted) throw new Error('cancelled')
}
const encode = (bytes: Uint8Array) =>
  btoa(Array.from(bytes, (x) => String.fromCharCode(x)).join(''))
const sha = async (bytes: Uint8Array) =>
  Array.from(
    new Uint8Array(await crypto.subtle.digest('SHA-256', Uint8Array.from(bytes).buffer)),
    (x) => x.toString(16).padStart(2, '0'),
  ).join('')
function scopeValid(scope: Scope, key: string) {
  if (
    typeof scope.documentId !== 'string' ||
    !scope.documentId ||
    scope.documentId.length > 2048 ||
    typeof scope.changeId !== 'string' ||
    !/^[A-Za-z0-9_-]{1,128}$/.test(scope.changeId) ||
    typeof key !== 'string' ||
    key.length > 128 ||
    !/^(?:snapshot|(?:page|image|receipt)-[0-9]+)$/.test(key)
  )
    invalid()
}
function refValid(backup: MasterBackupRef) {
  if (
    typeof backup.sha256 !== 'string' ||
    !/^[a-f0-9]{64}$/.test(backup.sha256) ||
    !Number.isSafeInteger(backup.sizeBytes) ||
    backup.sizeBytes < 1 ||
    backup.sizeBytes > MAX_BYTES
  )
    invalid()
}
function exact(body: Record<string, unknown>, required: string[], optional: string[] = []) {
  if (
    required.some((key) => !Object.hasOwn(body, key)) ||
    Object.keys(body).some((key) => !required.includes(key) && !optional.includes(key))
  )
    invalid()
}
async function call(
  scope: Scope,
  key: string,
  operation: string,
  data: Record<string, unknown> = {},
) {
  cancelled(scope.signal)
  if (scope.protocol !== undefined && !['master', 'package'].includes(scope.protocol)) invalid()
  if (scope.protocol === 'package')
    operation = operation.replace('master_backup_', 'package_backup_')
  const body = { operation, documentId: scope.documentId, changeId: scope.changeId, key, ...data }
  if (new TextEncoder().encode(JSON.stringify(body)).length > MAX_JSON) invalid()
  let response: Response
  try {
    response = await scope.request(body, scope.signal)
  } catch (error) {
    cancelled(scope.signal)
    if (error instanceof Error && ['cancelled', 'aborted'].includes(error.message))
      throw new Error('cancelled', { cause: error })
    throw error
  }
  cancelled(scope.signal)
  if (!response.body) invalid()
  const reader = response.body.getReader()
  const chunks: Uint8Array[] = []
  let size = 0
  try {
    for (;;) {
      const part = await reader.read()
      cancelled(scope.signal)
      if (part.done) break
      size += part.value.length
      if (size > MAX_JSON) invalid()
      chunks.push(part.value)
    }
  } catch (error) {
    if (scope.signal?.aborted) throw new Error('cancelled', { cause: error })
    throw error
  } finally {
    await reader.cancel().catch(() => {})
    reader.releaseLock()
  }
  const bytes = new Uint8Array(size)
  let offset = 0
  for (const chunk of chunks) {
    bytes.set(chunk, offset)
    offset += chunk.length
  }
  let result: unknown
  try {
    result = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes))
  } catch {
    invalid()
  }
  if (!result || typeof result !== 'object' || Array.isArray(result)) invalid()
  const value = result as Record<string, unknown>
  if (value.error === 'cancelled' || value.error === 'aborted') throw new Error('cancelled')
  if (
    value.error === 'presentation_master_backup_capacity' ||
    value.error === 'presentation_package_backup_capacity' ||
    value.error === 'quota_exceeded'
  )
    throw new Error('presentation_master_backup_capacity')
  if (!response.ok || Object.hasOwn(value, 'error')) invalid()
  return value
}
const identityKeys = ['documentId', 'changeId', 'key', 'sha256', 'sizeBytes']
function match(scope: Scope, backup: MasterBackupRef, body: Record<string, unknown>) {
  if (
    body.documentId !== scope.documentId ||
    body.changeId !== scope.changeId ||
    body.key !== backup.key ||
    body.sha256 !== backup.sha256 ||
    body.sizeBytes !== backup.sizeBytes
  )
    invalid()
}
function metadata(scope: Scope, backup: MasterBackupRef, body: Record<string, unknown>) {
  exact(body, [...identityKeys, 'status'], ['receivedBytes'])
  match(scope, backup, body)
  if (!['uploading', 'ready'].includes(String(body.status))) invalid()
  if (
    body.receivedBytes !== undefined &&
    (!Number.isSafeInteger(body.receivedBytes) ||
      Number(body.receivedBytes) < 0 ||
      Number(body.receivedBytes) > backup.sizeBytes)
  )
    invalid()
  if (
    body.status === 'ready' &&
    body.receivedBytes !== undefined &&
    body.receivedBytes !== backup.sizeBytes
  )
    invalid()
}
export async function saveMasterBackup(input: Save): Promise<MasterBackupRef> {
  const scope: Scope = {
    request: input.request,
    documentId: input.documentId,
    changeId: input.changeId,
    signal: input.signal,
    protocol: input.protocol,
  }
  const key = input.key
  scopeValid(scope, key)
  cancelled(scope.signal)
  if (
    !(input.bytes instanceof Uint8Array) ||
    input.bytes.length < 1 ||
    input.bytes.length > MAX_BYTES
  )
    invalid()
  const bytes = Uint8Array.from(input.bytes)
  const backup = { key, sha256: await sha(bytes), sizeBytes: bytes.length }
  let meta = await call(scope, key, 'master_backup_begin', {
    sha256: backup.sha256,
    sizeBytes: backup.sizeBytes,
  })
  metadata(scope, backup, meta)
  while (meta.status !== 'ready' && Number(meta.receivedBytes ?? 0) < bytes.length) {
    const offset = Number(meta.receivedBytes ?? 0)
    const chunk = bytes.subarray(offset, Math.min(offset + CHUNK, bytes.length))
    meta = await call(scope, key, 'master_backup_chunk', { offset, base64: encode(chunk) })
    metadata(scope, backup, meta)
    if (meta.receivedBytes !== offset + chunk.length) invalid()
  }
  if (meta.status !== 'ready') meta = await call(scope, key, 'master_backup_finish')
  metadata(scope, backup, meta)
  if (meta.status !== 'ready') invalid()
  await readMasterBackup({ ...scope, backup })
  cancelled(scope.signal)
  return backup
}
export async function readMasterBackup(input: Read): Promise<Uint8Array> {
  const scope: Scope = {
    request: input.request,
    documentId: input.documentId,
    changeId: input.changeId,
    signal: input.signal,
    protocol: input.protocol,
  }
  const backup = { ...input.backup }
  scopeValid(scope, backup.key)
  refValid(backup)
  const meta = await call(scope, backup.key, 'master_backup_status')
  metadata(scope, backup, meta)
  if (meta.status !== 'ready') invalid()
  const bytes = new Uint8Array(backup.sizeBytes)
  for (let offset = 0; offset < bytes.length; offset += CHUNK) {
    const length = Math.min(CHUNK, bytes.length - offset)
    const part = await call(scope, backup.key, 'master_backup_read', { offset, length })
    exact(part, [...identityKeys, 'offset', 'base64'])
    match(scope, backup, part)
    if (
      part.offset !== offset ||
      typeof part.base64 !== 'string' ||
      part.base64.length !== Math.ceil(length / 3) * 4 ||
      !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(part.base64)
    )
      invalid()
    let chunk: Uint8Array
    try {
      chunk = Uint8Array.from(atob(part.base64), (x) => x.charCodeAt(0))
    } catch {
      return invalid()
    }
    if (chunk.length !== length || encode(chunk) !== part.base64) invalid()
    bytes.set(chunk, offset)
  }
  if ((await sha(bytes)) !== backup.sha256) invalid()
  cancelled(scope.signal)
  return bytes
}
