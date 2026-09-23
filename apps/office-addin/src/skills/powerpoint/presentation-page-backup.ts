import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import type { InMemoryVfs } from '../shared/vfs.js'
import type {
  CompiledPresentationArtifact,
  PresentationImportRecord,
} from './presentation-delivery.js'
import {
  presentationArtifactContent,
  presentationImportKey,
  presentationPageMapping,
  validPresentationImportRecord,
} from './presentation-page-delivery.js'
import { parsePresentationProductionStatus } from './presentation-production.js'

export interface PresentationPageBackupOptions {
  available(): boolean
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  documentId(): Promise<string>
  vfs: InMemoryVfs
  artifact(projectId?: string): CompiledPresentationArtifact | undefined
  readReceipt(key: string): PresentationImportRecord | undefined
  adapter: {
    exportPresentationPagePackage(
      slideId: string,
      signal?: AbortSignal,
    ): Promise<{ slideId: string; slideIds: string[]; base64: string }>
  }
}
const MAX_BYTES = 8 * 1024 * 1024
const CHUNK_BYTES = 128 * 1024
const id = (v: unknown): v is string => typeof v === 'string' && /^[A-Za-z0-9_-]{1,80}$/.test(v)
const requestId = (v: unknown): v is string =>
  typeof v === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(v)
const hostId = (v: unknown): v is string =>
  typeof v === 'string' &&
  v.length > 0 &&
  v.length <= 256 &&
  !Array.from(v).some(
    (c) => c.charCodeAt(0) < 32 || (c.charCodeAt(0) >= 127 && c.charCodeAt(0) <= 159),
  )
const hash = (v: unknown): v is string => typeof v === 'string' && /^[a-f0-9]{64}$/.test(v)
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const object = (v: unknown, keys: string[]): v is Record<string, unknown> =>
  !!v &&
  typeof v === 'object' &&
  !Array.isArray(v) &&
  Object.keys(v).length === keys.length &&
  Object.keys(v).every((k) => keys.includes(k))
const hostIds = (v: unknown): v is string[] =>
  Array.isArray(v) &&
  v.length > 0 &&
  v.length <= 512 &&
  v.every(hostId) &&
  new Set(v).size === v.length
function invalid(): never {
  throw new Error('presentation_response_invalid')
}
function decode(value: unknown, maximum: number): Uint8Array {
  if (typeof value !== 'string' || !value || value.length > Math.ceil(maximum / 3) * 4) invalid()
  const binary = atob(value)
  if (btoa(binary) !== value || binary.length > maximum) invalid()
  return Uint8Array.from(binary, (c) => c.charCodeAt(0))
}
function encode(bytes: Uint8Array): string {
  let binary = ''
  for (const byte of bytes) binary += String.fromCharCode(byte)
  return btoa(binary)
}
async function digest(bytes: Uint8Array): Promise<string> {
  return Array.from(
    new Uint8Array(await crypto.subtle.digest('SHA-256', Uint8Array.from(bytes).buffer)),
    (b) => b.toString(16).padStart(2, '0'),
  ).join('')
}
interface Metadata {
  backupId: string
  projectId: string
  documentId: string
  requestId: string
  parentRequestId: string
  pageId: string
  hostSlideId: string
  slideIds: string[]
  sha256: string
  sizeBytes: number
  parentInputDigest: string
  inputDigest: string
  status: 'uploading' | 'ready'
  receivedBytes: number
}
function metadata(value: unknown): Metadata {
  if (
    !object(value, [
      'backupId',
      'projectId',
      'documentId',
      'requestId',
      'parentRequestId',
      'pageId',
      'hostSlideId',
      'slideIds',
      'sha256',
      'sizeBytes',
      'parentInputDigest',
      'inputDigest',
      'status',
      'receivedBytes',
    ]) ||
    !requestId(value.backupId) ||
    !id(value.projectId) ||
    typeof value.documentId !== 'string' ||
    !value.documentId ||
    value.documentId.length > 4096 ||
    !requestId(value.requestId) ||
    !requestId(value.parentRequestId) ||
    value.requestId === value.parentRequestId ||
    !id(value.pageId) ||
    !hostId(value.hostSlideId) ||
    !hostIds(value.slideIds) ||
    !value.slideIds.includes(value.hostSlideId) ||
    !hash(value.sha256) ||
    !hash(value.parentInputDigest) ||
    !hash(value.inputDigest) ||
    !Number.isSafeInteger(value.sizeBytes) ||
    Number(value.sizeBytes) < 1 ||
    Number(value.sizeBytes) > MAX_BYTES ||
    !Number.isSafeInteger(value.receivedBytes) ||
    Number(value.receivedBytes) < 0 ||
    Number(value.receivedBytes) > Number(value.sizeBytes) ||
    !['uploading', 'ready'].includes(String(value.status)) ||
    (value.status === 'ready' && value.receivedBytes !== value.sizeBytes)
  )
    invalid()
  return structuredClone(value) as unknown as Metadata
}
const idSchema = { type: 'string', pattern: '^[A-Za-z0-9_-]{1,80}$' }
const tools: AgentToolDef[] = [
  'save_presentation_page_backup',
  'read_presentation_page_backup',
].map((name) => ({
  name,
  description: name.startsWith('save')
    ? 'Save an immutable historical single-page PPTX backup of the imported parent page for a compiled page revision. Does not replace or delete any host page; later replacement must recheck its content and order.'
    : 'Download a ready historical page backup into session files. Does not verify current host content or perform restoration.',
  inputSchema: {
    type: 'object',
    properties: {
      project_id: idSchema,
      backup_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' },
      ...(name.startsWith('save')
        ? { page_id: idSchema, request_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,128}$' } }
        : {}),
    },
    required: name.startsWith('save')
      ? ['project_id', 'backup_id', 'page_id', 'request_id']
      : ['project_id', 'backup_id'],
    additionalProperties: false,
  },
}))
export function createPresentationPageBackupSkill(
  options: PresentationPageBackupOptions,
): AgentSkill & { clear(): void } {
  let epoch = 0
  return {
    id: 'office-presentation-page-backup',
    get tools() {
      return options.available() ? tools : []
    },
    systemPrompt:
      'Before future single-page replacement, save_presentation_page_backup persists the current imported parent page as a historical PPTX package bound to the fully compiled child revision. It never replaces, deletes or restores a page. read_presentation_page_backup downloads a historical snapshot, not proof that the current page is unchanged. Future replacement must freshly verify current page content, identity and order.',
    clear() {
      epoch++
    },
    async executeTool(call, signal) {
      const captured = epoch
      const check = () => {
        if (signal?.aborted || epoch !== captured) throw new Error('cancelled')
        if (!options.available()) throw new Error('presentation_unavailable')
      }
      try {
        check()
        if (call.inputError || call.truncated) throw new Error('invalid_tool_input')
        const save = call.name === 'save_presentation_page_backup'
        if (!save && call.name !== 'read_presentation_page_backup')
          throw new Error('invalid_tool_input')
        const input = structuredClone(call.input)
        if (
          !object(
            input,
            save
              ? ['project_id', 'backup_id', 'request_id', 'page_id']
              : ['project_id', 'backup_id'],
          ) ||
          !id(input.project_id) ||
          !requestId(input.backup_id) ||
          (save && (!id(input.page_id) || !requestId(input.request_id)))
        )
          throw new Error('invalid_tool_input')
        const projectId = input.project_id,
          backupId = input.backup_id
        const documentId = await options.documentId()
        check()
        if (typeof documentId !== 'string' || !documentId || documentId.length > 4096)
          throw new Error('presentation_document_changed')
        let bindingCheck = () => {}
        const current = async () => {
          check()
          if ((await options.documentId()) !== documentId)
            throw new Error('presentation_document_changed')
          check()
          bindingCheck()
        }
        const request = async (operation: string, body: Record<string, unknown>) => {
          await current()
          const payload = { operation, documentId, projectId, ...body }
          if (new TextEncoder().encode(JSON.stringify(payload)).length > 256 * 1024)
            throw new Error('presentation_request_too_large')
          const response = await options.request(payload, signal)
          await current()
          const text = await response.text()
          await current()
          if (new TextEncoder().encode(text).length > 256 * 1024) invalid()
          const value: unknown = JSON.parse(text)
          if (value && typeof value === 'object' && 'error' in value) {
            const error = value as { error: unknown; message?: unknown }
            if (
              error.error === 'unsupported' ||
              error.error === 'unsupported_operation' ||
              (error.error === 'invalid_request' &&
                typeof error.message === 'string' &&
                /(?:unsupported|unknown) operation/i.test(error.message))
            )
              throw new Error('presentation_upgrade_required')
            if (typeof error.error === 'string' && /^[a-z_]{1,80}$/.test(error.error))
              throw new Error(`presentation_${error.error}`)
            invalid()
          }
          if (!response.ok) invalid()
          return value
        }
        const parse = (value: unknown) => {
          const result = metadata(value)
          if (
            result.documentId !== documentId ||
            result.projectId !== projectId ||
            result.backupId !== backupId
          )
            invalid()
          return result
        }
        const publicResult = (record: Metadata, path?: string) => ({
          output: JSON.stringify({
            ...record,
            historical: true,
            currentPageVerified: false,
            ...(path ? { path } : {}),
          }),
          mutated: false,
          summary: path ? '已下载原页历史备份，未恢复宿主页' : '已保存原页历史备份，未替换宿主页',
        })
        if (!save) {
          const record = parse(await request('page_backup_status', { backupId }))
          if (record.status !== 'ready') throw new Error('presentation_page_backup_not_ready')
          const bytes = new Uint8Array(record.sizeBytes)
          for (let offset = 0; offset < bytes.length; offset += CHUNK_BYTES) {
            const length = Math.min(CHUNK_BYTES, bytes.length - offset)
            const part = await request('page_backup_read', { backupId, offset, length })
            if (
              !object(part, ['backupId', 'offset', 'sizeBytes', 'sha256', 'base64']) ||
              part.backupId !== backupId ||
              part.offset !== offset ||
              part.sizeBytes !== record.sizeBytes ||
              part.sha256 !== record.sha256
            )
              invalid()
            const chunk = decode(part.base64, length)
            if (chunk.length !== length) invalid()
            bytes.set(chunk, offset)
          }
          const actual = await digest(bytes)
          await current()
          if (actual !== record.sha256) invalid()
          const filename = await digest(
            new TextEncoder().encode(JSON.stringify([documentId, projectId, backupId])),
          )
          await current()
          const path = `/home/user/page-backups/${filename}.pptx`
          if (options.vfs.list('/home/user').includes(path))
            throw new Error('presentation_backup_file_exists')
          options.vfs.writeFile(path, bytes)
          return publicResult(record, path)
        }
        const artifact = options.artifact(projectId)
        if (
          !artifact ||
          artifact.pagePptxBase64 === undefined ||
          artifact.projectId !== projectId ||
          artifact.documentId !== documentId
        )
          throw new Error('presentation_restore_required')
        const content = presentationArtifactContent(artifact),
          snapshot = JSON.stringify(artifact),
          key = presentationImportKey(artifact)
        const receipt = options.readReceipt(key),
          receiptJson = JSON.stringify(receipt)
        if (
          !validPresentationImportRecord(receipt) ||
          receipt.state !== 'complete' ||
          !receipt.checkpoint ||
          receipt.checkpoint.version !== 2 ||
          receipt.documentId !== documentId ||
          !same(
            receipt.checkpoint.sourceSlideIds,
            artifact.pages?.map((p) => p.sourceSlideId),
          )
        )
          throw new Error('presentation_page_binding_invalid')
        const mapping = presentationPageMapping(artifact, receipt, input.page_id as string)
        if (!mapping || !hostId(mapping.slideId)) throw new Error('presentation_page_not_imported')
        bindingCheck = () => {
          if (
            options.artifact(projectId) !== artifact ||
            JSON.stringify(artifact) !== snapshot ||
            JSON.stringify(options.readReceipt(key)) !== receiptJson
          )
            throw new Error('presentation_page_stale')
        }
        const artifactDigest = await digest(new TextEncoder().encode(content))
        await current()
        if (artifactDigest !== receipt.checkpoint.artifactDigest)
          throw new Error('presentation_page_binding_invalid')
        const status = parsePresentationProductionStatus(
          await request('production_status', { requestId: input.request_id }),
        )
        if (
          status.projectId !== projectId ||
          status.requestId !== input.request_id ||
          status.status !== 'compiled' ||
          !status.revision ||
          status.revision.pageId !== input.page_id ||
          status.revision.parentRequestId !== artifact.requestId ||
          status.planRevision !== artifact.planRevision ||
          !same(
            status.pages.map((p) => p.id),
            artifact.pages?.map((p) => p.id),
          )
        )
          throw new Error('presentation_page_binding_invalid')
        const exported = await options.adapter.exportPresentationPagePackage(
          mapping.slideId,
          signal,
        )
        await current()
        if (
          !object(exported, ['slideId', 'slideIds', 'base64']) ||
          exported.slideId !== mapping.slideId ||
          !hostIds(exported.slideIds) ||
          !exported.slideIds.includes(mapping.slideId)
        )
          throw new Error('office_read_failed')
        const slideIds = [...exported.slideIds]
        const bytes = decode(exported.base64, MAX_BYTES)
        const sha256 = await digest(bytes)
        await current()
        const expected = {
          backupId,
          requestId: status.requestId,
          pageId: input.page_id,
          hostSlideId: mapping.slideId,
          slideIds,
          sha256,
          sizeBytes: bytes.length,
        }
        const verify = (value: unknown, previous?: Metadata) => {
          const record = parse(value)
          if (
            Object.entries(expected).some(([k, v]) => !same(record[k as keyof Metadata], v)) ||
            record.parentRequestId !== status.revision!.parentRequestId ||
            record.parentInputDigest !== status.revision!.parentInputDigest ||
            (previous && record.inputDigest !== previous.inputDigest)
          )
            invalid()
          return record
        }
        let record = verify(await request('page_backup_begin', expected))
        while (record.status !== 'ready' && record.receivedBytes < bytes.length) {
          const offset = record.receivedBytes,
            chunk = bytes.subarray(offset, Math.min(offset + CHUNK_BYTES, bytes.length))
          const next = verify(
            await request('page_backup_chunk', { backupId, offset, base64: encode(chunk) }),
            record,
          )
          if (next.receivedBytes !== offset + chunk.length) invalid()
          record = next
        }
        if (record.status !== 'ready')
          record = verify(await request('page_backup_finish', { backupId }), record)
        if (record.status !== 'ready') invalid()
        return publicResult(record)
      } catch (error) {
        const message = error instanceof Error ? error.message : ''
        return {
          output:
            /^(presentation_[a-z_]+|office_[a-z_]+|invalid_tool_input|cancelled|vfs_[a-z_]+)$/.test(
              message,
            )
              ? message
              : 'presentation_page_backup_failed',
          isError: true,
          mutated: false,
          summary: '页面备份未完成，宿主页未修改',
        }
      }
    },
  }
}
