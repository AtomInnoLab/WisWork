import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import type { CompiledPresentationArtifact } from './presentation-delivery.js'
import { presentationArtifactContent } from './presentation-page-delivery.js'
import {
  validatePresentationHistoryEntry,
  type PresentationHistoryEntry,
} from './presentation-change-history.js'
export interface PresentationChangeSetSummary {
  scope: { slideIds: string[]; shapeIds?: string[] }
  intent: string
  operations: { kind: PresentationHistoryEntry['kind']; pageId: string }[]
  preserved: string[]
  validation: string[]
  risk: 'medium' | 'high'
}
export function presentationChangeSetSummary(
  entry: PresentationHistoryEntry,
): PresentationChangeSetSummary {
  if (entry.kind === 'existing')
    return {
      scope: { slideIds: [entry.record.hostSlideId], shapeIds: [entry.record.shapeId] },
      intent: entry.record.kind === 'text' ? '修改现稿文字' : '调整现稿位置与尺寸',
      operations: [{ kind: 'existing', pageId: entry.record.hostSlideId }],
      preserved: ['目标以外对象（需复核）'],
      validation: ['目标对象回读', '受影响页面截图复核'],
      risk: 'medium',
    }
  if (entry.kind === 'existing_batch')
    return {
      scope: entry.record.scope,
      intent: entry.record.intent,
      operations: entry.record.operations.map((op) => ({
        kind: 'existing_batch',
        pageId: op.hostSlideId,
      })),
      preserved: entry.record.preserved,
      validation: entry.record.validation,
      risk: entry.record.risk,
    }
  if (entry.kind === 'existing_image')
    return {
      scope: entry.record.scope,
      intent: '替换现稿图片',
      operations: [{ kind: 'existing_image', pageId: entry.record.hostSlideId }],
      preserved: ['目标以外对象（需复核）'],
      validation: ['原图备份和新图回读', '受影响页面截图复核'],
      risk: 'high',
    }
  if (entry.kind === 'existing_page')
    return {
      scope: entry.record.scope,
      intent: '重做现稿单页',
      operations: [{ kind: 'existing_page', pageId: entry.record.oldSlideId }],
      preserved: ['其它页面（需复核）'],
      validation: ['原页备份与新页回读', '受影响页面截图复核'],
      risk: 'high',
    }
  const { kind, record: r } = entry
  const scope =
    entry.kind === 'page'
      ? {
          slideIds: Array.from(
            new Set(
              [
                entry.record.oldSlideId,
                entry.record.newSlideId,
                entry.record.restoredSlideId,
              ].filter((id): id is string => !!id),
            ),
          ),
        }
      : {
          slideIds: [entry.record.hostSlideId],
          shapeIds:
            entry.kind === 'image'
              ? Array.from(
                  new Set(
                    [
                      entry.record.oldShapeId,
                      entry.record.newShapeId,
                      entry.record.restoredShapeId,
                    ].filter((id): id is string => !!id),
                  ),
                )
              : [entry.record.shapeId],
        }
  return {
    scope,
    intent: { text: '修改文字', geometry: '调整位置与尺寸', image: '替换图片', page: '替换整页' }[
      kind
    ],
    operations: [{ kind, pageId: r.pageId }],
    preserved: [kind === 'page' ? '其它页面（需复核）' : '目标以外对象（需复核）'],
    validation: ['目标对象回读', '受影响页面截图复核'],
    risk: kind === 'text' || kind === 'geometry' ? 'medium' : 'high',
  }
}
export async function selectPresentationHistory(
  entries: PresentationHistoryEntry[],
  artifact: CompiledPresentationArtifact,
  documentId: string,
): Promise<
  Exclude<PresentationHistoryEntry, { kind: 'existing' | 'existing_batch' | 'existing_image' | 'existing_page' }>[]
> {
  if (
    !Array.isArray(entries) ||
    entries.length > 64 ||
    entries.some((e) => !validatePresentationHistoryEntry(e)) ||
    new Set(entries.map((e) => e.id)).size !== entries.length ||
    new Set(entries.map((e) => e.sequence)).size !== entries.length ||
    new TextEncoder().encode(JSON.stringify(entries)).byteLength > 1024 * 1024
  )
    throw new Error('presentation_change_history_invalid')
  if (artifact.documentId !== documentId) throw new Error('presentation_document_changed')
  const digest = Array.from(
    new Uint8Array(
      await crypto.subtle.digest(
        'SHA-256',
        new TextEncoder().encode(presentationArtifactContent(artifact)),
      ),
    ),
    (b) => b.toString(16).padStart(2, '0'),
  ).join('')
  const source = artifact.pagePptxBase64 === undefined ? undefined : 'production'
  return entries
    .filter(
      (
        e,
      ): e is Exclude<
        PresentationHistoryEntry,
        { kind: 'existing' | 'existing_batch' | 'existing_image' | 'existing_page' }
      > =>
        e.kind !== 'existing' &&
        e.kind !== 'existing_batch' &&
        e.kind !== 'existing_image' &&
        e.kind !== 'existing_page',
    )
    .filter((e) => {
      const r = e.record
      if (
        r.documentId !== documentId ||
        r.projectId !== artifact.projectId ||
        !artifact.pages?.some((p) => p.id === r.pageId)
      )
        return false
      if (e.kind === 'page')
        return (
          source === 'production' &&
          [e.record.parentRequestId, e.record.requestId].includes(artifact.requestId)
        )
      if (e.record.requestId !== artifact.requestId || e.record.source !== source) return false
      return e.kind === 'image' || e.record.artifactDigest === digest
    })
    .sort((a, b) => Number(a.legacy) - Number(b.legacy) || b.sequence - a.sequence)
}
interface Options {
  available(): boolean
  artifact(projectId?: string): CompiledPresentationArtifact | undefined
  documentId(): Promise<string>
  listChangeHistory(): PresentationHistoryEntry[]
}
const tool: AgentToolDef = {
  name: 'list_presentation_changes',
  description:
    'List bounded historical single-operation change summaries for the current task. Does not inspect current host content or establish QA acceptance. Use change_id for exact text/geometry read, inspection, undo or recovery; image tools use the recorded original shape_id. Revalidate each operation and confirm every write.',
  inputSchema: {
    type: 'object',
    properties: { project_id: { type: 'string', pattern: '^[A-Za-z0-9_-]{1,80}$' } },
    required: [],
    additionalProperties: false,
  },
}
export function createPresentationHistorySkill(options: Options): AgentSkill & { clear(): void } {
  let epoch = 0
  return {
    id: 'presentation-history',
    get tools() {
      return options.available() ? [tool] : []
    },
    systemPrompt:
      'list_presentation_changes reads historical single-operation ChangeSet summaries for the selected task. They record required validation, not validation success or current host state. Legacy order is unknown. Use the exact change_id for text/geometry history actions; undo one record at a time, respecting current object identity and value checks. Do not infer multi-operation atomicity, redo, or automatic risk approval.',
    clear() {
      epoch++
    },
    async executeTool(call, signal) {
      const captured = epoch
      const check = () => {
        if (signal?.aborted || captured !== epoch) throw new Error('cancelled')
        if (!options.available()) throw new Error('presentation_unavailable')
      }
      try {
        check()
        if (
          call.name !== tool.name ||
          call.inputError ||
          call.truncated ||
          Object.keys(call.input).some((k) => k !== 'project_id') ||
          (call.input.project_id !== undefined &&
            (typeof call.input.project_id !== 'string' ||
              !/^[A-Za-z0-9_-]{1,80}$/.test(call.input.project_id)))
        )
          throw new Error('invalid_tool_input')
        const projectId = call.input.project_id as string | undefined
        const artifact = options.artifact(projectId)
        if (!artifact || (projectId && artifact.projectId !== projectId))
          throw new Error('presentation_restore_required')
        const snapshot = JSON.stringify(artifact),
          documentId = await options.documentId()
        check()
        const entries = structuredClone(options.listChangeHistory()),
          history = JSON.stringify(entries)
        const selected = await selectPresentationHistory(entries, artifact, documentId)
        check()
        if (
          JSON.stringify(options.artifact(projectId)) !== snapshot ||
          (await options.documentId()) !== documentId ||
          JSON.stringify(options.listChangeHistory()) !== history
        )
          throw new Error('presentation_history_changed')
        check()
        const output = JSON.stringify({
          projectId: artifact.projectId,
          requestId: artifact.requestId,
          currentHostVerified: false,
          historyLimit: 64,
          changes: selected.map((e) => ({
            kind: e.kind,
            state: e.record.state,
            sequence: e.legacy ? null : e.sequence,
            legacyOrderUnknown: e.legacy,
            ...('changeId' in e.record
              ? { change_id: e.record.changeId }
              : { shape_id: e.record.oldShapeId }),
            page_id: e.record.pageId,
            changeSet: presentationChangeSetSummary(e),
          })),
        })
        if (new TextEncoder().encode(output).byteLength > 256 * 1024)
          throw new Error('presentation_history_output_too_large')
        return { output, mutated: false, summary: '当前任务变更历史；尚未核验宿主现状或QA' }
      } catch (error) {
        const message = error instanceof Error ? error.message : ''
        return {
          output: /^(presentation_[a-z_]+|invalid_tool_input|cancelled)$/.test(message)
            ? message
            : 'presentation_history_read_failed',
          isError: true,
          mutated: false,
          summary: '变更历史读取未完成',
        }
      }
    },
  }
}
