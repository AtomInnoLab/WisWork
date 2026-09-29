import type { AgentSkill } from '@wiswork/agent-core'
import {
  parsePresentationResearchDeleteReceipt,
  type PresentationResearchDeleteReceipt,
  parsePresentationResearchRecord,
  parsePresentationResearchSummary,
  type PresentationResearchRecord,
  type PresentationResearchSummary,
} from '@wiswork/project-store/presentation-research'

import { readPresentationResearchCapabilities } from '../skills/powerpoint/presentation-research-capabilities.js'
export interface PresentationResearchDeleteAttempt {
  documentId: string
  projectId: string
  ledgerId: string
  sequence: number
  draftDigest: string
  deleteId: string
  expectedRevision: number
}
export interface PresentationResearchSnapshot {
  phase: 'idle' | 'loading' | 'reading' | 'exporting' | 'deleting' | 'checkingDelete'
  available?: boolean
  cleanupAvailable?: boolean
  deleteAttempt?: PresentationResearchDeleteAttempt
  deleteReceipt?: PresentationResearchDeleteReceipt
  projectId?: string
  summary?: PresentationResearchSummary
  record?: PresentationResearchRecord
  notice?: string
  error?: string
}
export interface PresentationResearchController {
  snapshot(): PresentationResearchSnapshot
  subscribe(listener: () => void): () => void
  refresh(): Promise<void>
  selectProject(projectId: string): Promise<void>
  read(ledgerId: string): Promise<void>
  export(ledgerId: string): Promise<void>
  deleteRecord(ledgerId: string, draftDigest: string): Promise<void>
  retryDelete(): Promise<void>
  checkDeleteStatus(): Promise<void>
  cancel(): void
  clear(): void
}
const idValid = (value: unknown): value is string =>
  typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
const message = (error: unknown) => {
  const code = error instanceof Error ? error.message : ''
  if (code === 'presentation_cleanup_recovery_unavailable')
    return '无法保存清理恢复身份，未执行删除；请恢复本机界面存储后重试。'
  if (code === 'presentation_upgrade_required')
    return '当前 PC 版本不支持研究归档清理；已有记录仍保留。'
  if (code === 'presentation_record_running') return '记录缺少结束回执，不能清理。'
  if (code === 'presentation_record_protected')
    return '该研究记录仍被当前或历史计划、冻结任务引用，不能清理。'
  if (code === 'presentation_cleanup_quota_exceeded')
    return '清理回执已达容量上限；未删除该研究记录，已有资料保留。'
  if (code === 'presentation_not_found')
    return '尚未找到删除回执；请刷新状态，或明确重试同一次删除。'
  if (code === 'presentation_request_conflict')
    return '删除身份发生冲突；请读取回执核对，勿提交新的删除尝试。'
  if (code === 'presentation_quota_exceeded')
    return '研究账本已达本机容量上限；已有记录仍可读取，新记录暂无法保存。'
  if (code === 'presentation_revision_conflict')
    return '研究记录已有更新；请刷新核对，当前删除尝试不能更换身份或重算版本。'
  if (code === 'presentation_document_changed') return '文档已改变，请在目标文档刷新研究记录。'
  if (code === 'cancelled') return '已停止等待；本机可能已保存研究记录，请刷新读取，勿自动重放。'
  return '研究记录暂时无法读取或导出；已有资料保留，请刷新本机记录后重试。'
}
export function createPresentationResearchController(options: {
  available(): boolean
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  documentId(): Promise<string>
  lastProject?(): string | undefined
  readDeleteAttempt?(documentId: string): unknown
  writeDeleteAttempt?(
    documentId: string,
    value: PresentationResearchDeleteAttempt | undefined,
  ): void
  executeTool: AgentSkill['executeTool']
}): PresentationResearchController {
  let state: PresentationResearchSnapshot = { phase: 'idle' }
  let selected: string | undefined
  let epoch = 0
  let active: AbortController | undefined
  const listeners = new Set<() => void>()
  const publish = (next: PresentationResearchSnapshot) => {
    state = next
    for (const listener of listeners) listener()
  }
  const clear = () => {
    epoch++
    active?.abort()
    active = undefined
    selected = undefined
    publish({ phase: 'idle', available: state.available, cleanupAvailable: state.cleanupAvailable })
  }
  const run = async (kind: 'list' | 'read' | 'export', ledgerId?: string) => {
    if (active) return
    if (ledgerId !== undefined && !idValid(ledgerId)) return
    const projectId = selected ?? options.lastProject?.()
    if (projectId !== undefined && !idValid(projectId)) {
      publish({ phase: 'idle', error: message(undefined) })
      return
    }
    const controller = new AbortController()
    const captured = ++epoch
    active = controller
    const previous = state.projectId === projectId ? state : { phase: 'idle' as const }
    publish({
      ...previous,
      projectId,
      phase: kind === 'list' ? 'loading' : kind === 'read' ? 'reading' : 'exporting',
      error: undefined,
      notice: undefined,
    })
    const check = () => {
      if (captured !== epoch || controller.signal.aborted) throw new Error('cancelled')
      if (!options.available()) throw new Error('presentation_unavailable')
    }
    let documentId: string | undefined
    let supported = previous.available
    try {
      check()
      documentId = await options.documentId()
      const current = async () => {
        check()
        if ((await options.documentId()) !== documentId)
          throw new Error('presentation_document_changed')
        check()
      }
      await current()
      const capability = await readPresentationResearchCapabilities(async (operation, body) => {
        await current()
        const response = await options.request(
          { operation, documentId, ...body },
          controller.signal,
        )
        await current()
        if (!response.ok) throw new Error('presentation_service_unavailable')
        const text = await response.text()
        await current()
        if (new TextEncoder().encode(text).byteLength > 1024)
          throw new Error('presentation_response_invalid')
        const value = JSON.parse(text)
        if (['invalid_request', 'upgrade_required'].includes(value?.error))
          throw new Error('presentation_upgrade_required')
        return value
      })
      supported = capability.available
      if (!capability.available) {
        publish({ phase: 'idle', available: false, cleanupAvailable: false })
        return
      }
      if (!projectId) {
        publish({ phase: 'idle', available: true, cleanupAvailable: capability.cleanupAvailable })
        return
      }
      const execute = async (name: string, input: Record<string, unknown>, max: number) => {
        await current()
        const result = await options.executeTool(
          { id: `research-control-${captured}`, name, input },
          controller.signal,
        )
        await current()
        if (result.isError) throw new Error(result.output)
        if (new TextEncoder().encode(result.output).byteLength > max)
          throw new Error('presentation_response_invalid')
        return JSON.parse(result.output) as unknown
      }
      const summary = parsePresentationResearchSummary(
        await execute('list_research_ledgers', { project_id: projectId }, 64 * 1024),
      )
      if (summary.documentId !== documentId || summary.projectId !== projectId)
        throw new Error('presentation_response_invalid')
      let record = previous.record
      if (
        record &&
        !summary.records.some(
          (item) =>
            item.id === record!.id &&
            item.draftDigest === record!.draftDigest &&
            item.state === record!.state,
        )
      )
        record = undefined
      let notice: string | undefined
      if (kind === 'read') {
        record = parsePresentationResearchRecord(
          await execute(
            'read_research_ledger',
            { project_id: projectId, ledger_id: ledgerId },
            512 * 1024,
          ),
        )
        if (
          record.documentId !== documentId ||
          record.projectId !== projectId ||
          record.id !== ledgerId
        )
          throw new Error('presentation_response_invalid')
      } else if (kind === 'export') {
        await execute(
          'export_research_ledger',
          { project_id: projectId, ledger_id: ledgerId },
          64 * 1024,
        )
        notice = '研究 JSON 与 Markdown 已放回会话附件；结论、来源权威性与时效仍待核验。'
      }
      let deleteAttempt = previous.deleteAttempt
      if (!deleteAttempt && capability.cleanupAvailable && options.readDeleteAttempt) {
        const saved = options.readDeleteAttempt(documentId)
        if (saved && typeof saved === 'object' && !Array.isArray(saved)) {
          const value = saved as PresentationResearchDeleteAttempt
          if (
            Object.keys(value).sort().join(',') ===
              'deleteId,documentId,draftDigest,expectedRevision,ledgerId,projectId,sequence' &&
            value.documentId === documentId &&
            value.projectId === projectId &&
            [value.ledgerId, value.projectId, value.deleteId].every(idValid) &&
            /^[a-f0-9]{64}$/.test(value.draftDigest) &&
            Number.isSafeInteger(value.sequence) &&
            value.sequence > 0 &&
            Number.isSafeInteger(value.expectedRevision) &&
            value.expectedRevision > 0
          )
            deleteAttempt = structuredClone(value)
        }
      }
      await current()
      publish({
        phase: 'idle',
        available: true,
        cleanupAvailable: capability.cleanupAvailable,
        ...(deleteAttempt ? { deleteAttempt } : {}),
        ...(previous.deleteReceipt ? { deleteReceipt: previous.deleteReceipt } : {}),
        projectId,
        summary,
        ...(record ? { record } : {}),
        ...(notice ? { notice } : {}),
      })
    } catch (error) {
      if (captured !== epoch) return
      const changed =
        documentId && (await options.documentId().catch(() => undefined)) !== documentId
      if (captured !== epoch) return
      publish({
        ...(changed ? {} : { ...previous, available: supported }),
        ...(error instanceof Error && error.message === 'presentation_upgrade_required'
          ? { available: false, cleanupAvailable: false }
          : {}),
        phase: 'idle',
        ...(changed ? {} : { projectId }),
        error: message(changed ? new Error('presentation_document_changed') : error),
      })
    } finally {
      if (captured === epoch) active = undefined
    }
  }
  const runDelete = async (
    kind: 'delete' | 'retry' | 'status',
    ledgerId?: string,
    draftDigest?: string,
  ) => {
    if (active || !state.cleanupAvailable || !state.projectId) return
    const controller = new AbortController()
    const captured = ++epoch
    const previous = state
    active = controller
    let attempt = previous.deleteAttempt
    let documentId: string | undefined
    const check = async () => {
      if (captured !== epoch || controller.signal.aborted) throw new Error('cancelled')
      if (!options.available()) throw new Error('presentation_unavailable')
      if ((await options.documentId()) !== documentId)
        throw new Error('presentation_document_changed')
      if (captured !== epoch || controller.signal.aborted) throw new Error('cancelled')
    }
    const request = async (operation: string, body: Record<string, unknown>) => {
      await check()
      const response = await options.request(
        { operation, documentId, projectId: attempt!.projectId, ...body },
        controller.signal,
      )
      await check()
      if (!response.ok) throw new Error('presentation_service_unavailable')
      const text = await response.text()
      await check()
      if (new TextEncoder().encode(text).byteLength > 4096)
        throw new Error('presentation_response_invalid')
      const value = JSON.parse(text)
      if (value?.error) throw new Error(`presentation_${value.error}`)
      const receipt = parsePresentationResearchDeleteReceipt(value)
      if (
        receipt.documentId !== documentId ||
        receipt.projectId !== attempt!.projectId ||
        receipt.ledgerId !== attempt!.ledgerId ||
        receipt.sequence !== attempt!.sequence ||
        receipt.draftDigest !== attempt!.draftDigest ||
        receipt.deleteId !== attempt!.deleteId ||
        receipt.revision !== attempt!.expectedRevision + 1
      )
        throw new Error('presentation_response_invalid')
      return receipt
    }
    try {
      documentId = await options.documentId()
      await check()
      if (kind === 'delete') {
        if (previous.summary?.documentId !== documentId)
          throw new Error('presentation_document_changed')
        if (attempt) throw new Error('presentation_request_conflict')
        const record = previous.summary?.records.find((item) => item.id === ledgerId)
        if (!record || record.draftDigest !== draftDigest)
          throw new Error('presentation_revision_conflict')
        if (record.state === 'running') throw new Error('presentation_record_running')
        attempt = {
          documentId,
          projectId: previous.projectId!,
          ledgerId: record.id,
          sequence: record.sequence,
          draftDigest: record.draftDigest,
          deleteId: crypto.randomUUID(),
          expectedRevision: previous.summary!.revision,
        }
      }
      if (!attempt) return
      if (attempt.documentId !== documentId || attempt.projectId !== previous.projectId)
        throw new Error('presentation_document_changed')
      if (kind !== 'status') {
        try {
          options.writeDeleteAttempt?.(documentId, attempt)
        } catch {
          throw new Error('presentation_cleanup_recovery_unavailable')
        }
      }
      publish({
        ...previous,
        deleteAttempt: attempt,
        phase: kind === 'status' ? 'checkingDelete' : 'deleting',
        error: undefined,
        notice: undefined,
      })
      let receipt: PresentationResearchDeleteReceipt
      if (kind === 'status')
        receipt = await request('research_delete_status', { deleteId: attempt.deleteId })
      else {
        try {
          receipt = await request('research_delete', {
            ledgerId: attempt.ledgerId,
            deleteId: attempt.deleteId,
            expectedDraftDigest: attempt.draftDigest,
            expectedRevision: attempt.expectedRevision,
          })
        } catch (error) {
          await check()
          if (
            error instanceof Error &&
            /^presentation_(record_protected|record_running|cleanup_quota_exceeded|revision_conflict|request_conflict|record_deleted|aborted|upgrade_required|invalid_state)$/.test(
              error.message,
            )
          )
            throw error
          receipt = await request('research_delete_status', { deleteId: attempt.deleteId })
        }
      }
      await check()
      options.writeDeleteAttempt?.(documentId, undefined)
      publish({
        ...state,
        phase: 'checkingDelete',
        deleteAttempt: undefined,
        deleteReceipt: receipt,
        record: previous.record?.id === attempt.ledgerId ? undefined : previous.record,
        notice: '本机研究归档已清理；原附件、PowerPoint 文稿、交付包与导出副本仍保留。',
        error: undefined,
      })
      const result = await options.executeTool(
        {
          id: `research-delete-list-${captured}`,
          name: 'list_research_ledgers',
          input: { project_id: attempt.projectId },
        },
        controller.signal,
      )
      await check()
      if (result.isError || new TextEncoder().encode(result.output).byteLength > 64 * 1024)
        throw new Error('presentation_response_invalid')
      const summary = parsePresentationResearchSummary(JSON.parse(result.output))
      if (summary.documentId !== documentId || summary.projectId !== attempt.projectId)
        throw new Error('presentation_response_invalid')
      publish({ ...state, summary, phase: 'idle' })
    } catch (error) {
      if (captured !== epoch) return
      const changed =
        (error instanceof Error && error.message === 'presentation_document_changed') ||
        (documentId && (await options.documentId().catch(() => undefined)) !== documentId)
      if (captured !== epoch) return
      let cleanupError = error
      if (
        !changed &&
        documentId &&
        error instanceof Error &&
        /^presentation_(record_protected|record_running|cleanup_quota_exceeded|revision_conflict)$/.test(
          error.message,
        )
      ) {
        try {
          options.writeDeleteAttempt?.(documentId, undefined)
          attempt = undefined
          state = { ...state, deleteAttempt: undefined }
        } catch {
          cleanupError = new Error('presentation_cleanup_recovery_unavailable')
        }
      }
      publish({
        ...(changed ? {} : state),
        phase: 'idle',
        ...(changed || state.deleteReceipt ? {} : attempt ? { deleteAttempt: attempt } : {}),
        error: message(changed ? new Error('presentation_document_changed') : cleanupError),
      })
    } finally {
      if (captured === epoch) {
        active = undefined
        if (state.phase !== 'idle') publish({ ...state, phase: 'idle' })
      }
    }
  }
  return {
    snapshot: () => state,
    subscribe: (listener) => {
      listeners.add(listener)
      return () => {
        listeners.delete(listener)
      }
    },
    refresh: async () => {
      await run('list')
      if (state.deleteAttempt && !active) await runDelete('status')
    },
    selectProject: async (projectId) => {
      if (!idValid(projectId)) return
      clear()
      selected = projectId
      await run('list')
    },
    read: (ledgerId) => run('read', ledgerId),
    export: (ledgerId) => run('export', ledgerId),
    deleteRecord: (ledgerId, draftDigest) => runDelete('delete', ledgerId, draftDigest),
    retryDelete: () => runDelete('retry'),
    checkDeleteStatus: () => runDelete('status'),
    cancel: () => {
      epoch++
      active?.abort()
      active = undefined
      publish({ ...state, phase: 'idle', notice: message(new Error('cancelled')) })
    },
    clear,
  }
}
