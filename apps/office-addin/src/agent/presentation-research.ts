import type { AgentSkill } from '@wiswork/agent-core'
import {
  parsePresentationResearchRecord,
  parsePresentationResearchSummary,
  type PresentationResearchRecord,
  type PresentationResearchSummary,
} from '@wiswork/project-store/presentation-research'

export interface PresentationResearchSnapshot {
  phase: 'idle' | 'loading' | 'reading' | 'exporting'
  available?: boolean
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
  cancel(): void
  clear(): void
}
const idValid = (value: unknown): value is string =>
  typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
const message = (error: unknown) => {
  const code = error instanceof Error ? error.message : ''
  if (code === 'presentation_quota_exceeded')
    return '研究账本已达本机容量上限；已有记录仍可读取，新记录暂无法保存。'
  if (code === 'presentation_revision_conflict')
    return '研究记录已有更新，请刷新后明确提交新的整理尝试。'
  if (code === 'presentation_document_changed') return '文档已改变，请在目标文档刷新研究记录。'
  if (code === 'cancelled') return '已停止等待；本机可能已保存研究记录，请刷新读取，勿自动重放。'
  return '研究记录暂时无法读取或导出；已有资料保留，请刷新本机记录后重试。'
}
export function createPresentationResearchController(options: {
  available(): boolean
  request(body: unknown, signal?: AbortSignal): Promise<Response>
  documentId(): Promise<string>
  lastProject?(): string | undefined
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
    publish({ phase: 'idle', available: state.available })
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
      const response = await options.request(
        { operation: 'research_capabilities', documentId },
        controller.signal,
      )
      await current()
      if (!response.ok) throw new Error('presentation_service_unavailable')
      const text = await response.text()
      await current()
      if (new TextEncoder().encode(text).byteLength > 1024)
        throw new Error('presentation_response_invalid')
      const capability = JSON.parse(text)
      if (['invalid_request', 'upgrade_required'].includes(capability?.error)) {
        publish({ phase: 'idle', available: false })
        return
      }
      if (
        !capability ||
        Object.keys(capability).sort().join(',') !== 'available,version' ||
        capability.version !== 1 ||
        capability.available !== true
      )
        throw new Error('presentation_response_invalid')
      supported = true
      if (!projectId) {
        publish({ phase: 'idle', available: true })
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
      await current()
      publish({
        phase: 'idle',
        available: true,
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
        phase: 'idle',
        ...(changed ? {} : { projectId }),
        error: message(changed ? new Error('presentation_document_changed') : error),
      })
    } finally {
      if (captured === epoch) active = undefined
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
    refresh: () => run('list'),
    selectProject: async (projectId) => {
      if (!idValid(projectId)) return
      clear()
      selected = projectId
      await run('list')
    },
    read: (ledgerId) => run('read', ledgerId),
    export: (ledgerId) => run('export', ledgerId),
    cancel: () => {
      epoch++
      active?.abort()
      active = undefined
      publish({ ...state, phase: 'idle', notice: message(new Error('cancelled')) })
    },
    clear,
  }
}
