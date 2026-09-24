import type { PresentationGenerationOptions } from './presentation-generation.js'
export interface PresentationImageBackupMetadata {
  attachmentId: string
  sizeBytes: number
  mime: 'image/png' | 'image/jpeg'
}
const LIMIT = 2 * 1024 * 1024
const CHUNK = 128 * 1024
const idValid = (value: unknown): value is string =>
  typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
const fail = (): never => {
  throw new Error('presentation_image_backup_invalid')
}
const integer = (value: unknown, min: number, max: number): value is number =>
  Number.isSafeInteger(value) && Number(value) >= min && Number(value) <= max
function encode(bytes: Uint8Array): string {
  let binary = ''
  for (const byte of bytes) binary += String.fromCharCode(byte)
  return btoa(binary)
}
function decode(value: unknown, limit: number): Uint8Array<ArrayBuffer> {
  if (
    typeof value !== 'string' ||
    value.length > Math.ceil(limit / 3) * 4 ||
    !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(value)
  )
    return fail()
  const bytes = Uint8Array.from(atob(value), (char) => char.charCodeAt(0))
  if (bytes.length > limit || encode(bytes) !== value) return fail()
  return bytes
}
function mime(bytes: Uint8Array): PresentationImageBackupMetadata['mime'] {
  if ([137, 80, 78, 71, 13, 10, 26, 10].every((n, i) => bytes[i] === n)) return 'image/png'
  if (bytes[0] === 255 && bytes[1] === 216 && bytes[2] === 255) return 'image/jpeg'
  return fail()
}
async function digest(bytes: Uint8Array<ArrayBuffer>) {
  return Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), (b) =>
    b.toString(16).padStart(2, '0'),
  ).join('')
}
function record(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return fail()
  return value as Record<string, unknown>
}
function validMetadata(value: PresentationImageBackupMetadata) {
  if (
    !value ||
    Object.keys(value).some((k) => !['attachmentId', 'sizeBytes', 'mime'].includes(k)) ||
    !idValid(value.attachmentId) ||
    !integer(value.sizeBytes, 1, LIMIT) ||
    !['image/png', 'image/jpeg'].includes(value.mime)
  )
    fail()
}
export function createPresentationImageBackup(
  options: Pick<PresentationGenerationOptions, 'available' | 'request'>,
) {
  function scope(documentId: string, signal?: AbortSignal) {
    const check = () => {
      if (signal?.aborted) throw new Error('cancelled')
      if (!options.available()) throw new Error('presentation_image_backup_unavailable')
      if (typeof documentId !== 'string' || !documentId.length || documentId.length > 2048) fail()
    }
    return {
      check,
      async request(body: Record<string, unknown>) {
        check()
        const response = await options.request({ ...body, documentId }, signal)
        check()
        if (!response.ok) throw new Error('presentation_image_backup_unavailable')
        const text = await response.text()
        check()
        if (new TextEncoder().encode(text).length > 256 * 1024) return fail()
        let value: unknown
        try {
          value = JSON.parse(text)
        } catch {
          return fail()
        }
        const result = record(value)
        if ('error' in result) throw new Error('presentation_image_backup_unavailable')
        return result
      },
    }
  }
  async function load(
    documentId: string,
    metadata: PresentationImageBackupMetadata,
    signal?: AbortSignal,
  ): Promise<string> {
    validMetadata(metadata)
    const { check, request } = scope(documentId, signal)
    check()
    const bytes = new Uint8Array(metadata.sizeBytes)
    for (let offset = 0; offset < bytes.length; offset += CHUNK) {
      const length = Math.min(CHUNK, bytes.length - offset)
      const response = await request({
        operation: 'attachment_original',
        attachmentId: metadata.attachmentId,
        offset,
        length,
      })
      if (
        Object.keys(response).some(
          (k) => !['attachmentId', 'offset', 'sizeBytes', 'sha256', 'mime', 'base64'].includes(k),
        ) ||
        response.attachmentId !== metadata.attachmentId ||
        response.sha256 !== metadata.attachmentId ||
        response.offset !== offset ||
        response.sizeBytes !== metadata.sizeBytes ||
        response.mime !== metadata.mime
      )
        return fail()
      const chunk = decode(response.base64, CHUNK)
      if (chunk.length !== length) return fail()
      bytes.set(chunk, offset)
    }
    if ((await digest(bytes)) !== metadata.attachmentId || mime(bytes) !== metadata.mime)
      return fail()
    check()
    return encode(bytes)
  }
  return {
    available: options.available,
    load,
    async save(
      documentId: string,
      base64: string,
      signal?: AbortSignal,
    ): Promise<PresentationImageBackupMetadata> {
      const { check, request } = scope(documentId, signal)
      check()
      const bytes = decode(base64, LIMIT)
      const metadata = {
        attachmentId: await digest(bytes),
        sizeBytes: bytes.length,
        mime: mime(bytes),
      }
      check()
      validMetadata(metadata)
      const list = await request({ operation: 'attachment_list_assets' })
      if (
        Object.keys(list).length !== 1 ||
        !Array.isArray(list.attachments) ||
        list.attachments.length > 32
      )
        return fail()
      const seen = new Set<string>()
      let name = `image-backup-${metadata.attachmentId}.${metadata.mime === 'image/png' ? 'png' : 'jpg'}`
      for (const item of list.attachments) {
        const entry = record(item)
        if (!idValid(entry.attachmentId) || seen.has(entry.attachmentId)) return fail()
        seen.add(entry.attachmentId)
        if (entry.attachmentId !== metadata.attachmentId) continue
        if (
          entry.sha256 !== metadata.attachmentId ||
          entry.sizeBytes !== bytes.length ||
          typeof entry.name !== 'string' ||
          entry.name.length > 180 ||
          !/^[^\\/]+\.(png|jpe?g)$/i.test(entry.name) ||
          Array.from(entry.name).some(
            (char) => char.charCodeAt(0) < 32 || char.charCodeAt(0) === 127,
          )
        )
          return fail()
        name = entry.name
      }
      const validate = (value: Record<string, unknown>) => {
        if (
          value.attachmentId !== metadata.attachmentId ||
          value.sha256 !== metadata.attachmentId ||
          value.name !== name ||
          value.sizeBytes !== bytes.length ||
          !integer(value.receivedBytes, 0, bytes.length) ||
          !['uploading', 'ready', 'failed'].includes(String(value.status)) ||
          (value.status === 'ready' &&
            (value.kind !== 'image' || value.receivedBytes !== bytes.length))
        )
          return fail()
        return value as Record<string, unknown> & { receivedBytes: number }
      }
      let result = validate(
        await request({
          operation: 'attachment_begin',
          attachmentId: metadata.attachmentId,
          sha256: metadata.attachmentId,
          name,
          sizeBytes: bytes.length,
        }),
      )
      while (result.receivedBytes < bytes.length) {
        const offset = result.receivedBytes,
          end = Math.min(offset + CHUNK, bytes.length)
        result = validate(
          await request({
            operation: 'attachment_chunk',
            attachmentId: metadata.attachmentId,
            offset,
            base64: encode(bytes.subarray(offset, end)),
          }),
        )
        if (result.receivedBytes !== end || result.status !== 'uploading') return fail()
      }
      if (result.status !== 'ready')
        result = validate(
          await request({ operation: 'attachment_finish', attachmentId: metadata.attachmentId }),
        )
      if (result.status !== 'ready') return fail()
      if ((await load(documentId, metadata, signal)) !== base64) return fail()
      check()
      return metadata
    },
  }
}
