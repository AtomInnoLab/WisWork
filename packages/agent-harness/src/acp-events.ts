import type { SessionNotification, SessionUpdate, ToolKind } from '@agentclientprotocol/sdk'
import type { AgentToolCall, ToolExecution } from '@wiswork/agent-core'

const READ_TOOLS = /^(get|list|read|inspect|screenshot|verify|review)_/
const SEARCH_TOOLS = /^(search|find)_/
const FETCH_TOOLS = /^(fetch|download|insert_web_image)/
const EDIT_TOOLS = /^(add|apply|build|create|duplicate|edit|insert|replace|set|update|write)_/
const DELETE_TOOLS = /^(delete|remove)_/

export function acpToolKind(name: string): ToolKind {
  if (READ_TOOLS.test(name)) return 'read'
  if (SEARCH_TOOLS.test(name)) return 'search'
  if (FETCH_TOOLS.test(name)) return 'fetch'
  if (EDIT_TOOLS.test(name)) return 'edit'
  if (DELETE_TOOLS.test(name)) return 'delete'
  return 'other'
}

export function acpToolTitle(name: string): string {
  const words = name.replace(/[_-]+/g, ' ').trim()
  return words ? words[0]!.toUpperCase() + words.slice(1) : 'Agent tool'
}

const PRESENTATION_STAGES: ReadonlyArray<readonly [RegExp, string]> = [
  [/^(?:restore|resume)_presentation_project$/, '项目恢复'],
  [/^(?:save|read)_presentation_plan$/, '演示文稿计划'],
  [
    /^(?:start|pause|resume|cancel)_presentation_production_job$|^run_presentation_production$|^compile_deck_with_pptxgenjs$/,
    '逐页制作',
  ],
  [
    /^prepare_presentation_production_import$|^import_generated_presentation$|^read_presentation_import_status$/,
    '页面导入',
  ],
  [
    /^(?:capture_presentation_page_qa|record_presentation_page_review|read_presentation_qa|verify_slides|screenshot_slide)$/,
    '页面审查',
  ],
  [/^(?:read|record|export)_presentation_(?:delivery_report|issue_action)$/, '内容证据'],
  [
    /^(?:list_presentation_attachments|read_presentation_attachment|audit_presentation_sources)$/,
    '参考资料',
  ],
  [/^(?:edit|replace|restore|resume)_.*presentation_.*$|^list_presentation_changes$/, '页面修改'],
]

export type PresentationStage =
  | 'project_recovery'
  | 'planning'
  | 'production'
  | 'import'
  | 'review'
  | 'evidence'
  | 'sources'
  | 'editing'

const PRESENTATION_STAGE_KEYS: readonly PresentationStage[] = [
  'project_recovery',
  'planning',
  'production',
  'import',
  'review',
  'evidence',
  'sources',
  'editing',
]

export function acpPresentationStage(name: string): PresentationStage | undefined {
  const index = PRESENTATION_STAGES.findIndex(([pattern]) => pattern.test(name))
  return index < 0 ? undefined : PRESENTATION_STAGE_KEYS[index]
}

/** User-facing activity only. Project completion still comes from persisted receipts. */
export function acpToolActivity(name: string, state: 'running' | 'complete' | 'error'): string {
  const stage = PRESENTATION_STAGES.find(([pattern]) => pattern.test(name))?.[1]
  if (stage)
    return state === 'running'
      ? `正在处理${stage}…`
      : state === 'error'
        ? `${stage}处理未完成`
        : `${stage}操作已结束`
  const attachment = name === 'read' || name === 'bash'
  const read = /^(?:get_|read_|list_|search_|screenshot_|verify_)/.test(name)
  const action = attachment ? '处理附件' : read ? '读取内容' : '准备修改'
  return state === 'running'
    ? `正在${action}…`
    : state === 'error'
      ? `${action}未完成`
      : `已${action}`
}

export function acpNotification(sessionId: string, update: SessionUpdate): SessionNotification {
  return { sessionId, update }
}

export function acpToolStarted(call: AgentToolCall): SessionUpdate {
  return {
    sessionUpdate: 'tool_call',
    toolCallId: call.id,
    name: call.name,
    title: acpToolTitle(call.name),
    kind: acpToolKind(call.name),
    status: 'in_progress',
  }
}

export function acpToolFinished(call: AgentToolCall, execution: ToolExecution): SessionUpdate {
  return {
    sessionUpdate: 'tool_call_update',
    toolCallId: call.id,
    status: execution.isError ? 'failed' : 'completed',
    title: execution.summary,
    content: [
      {
        type: 'content',
        content: { type: 'text', text: execution.summary },
      },
    ],
    _meta: {
      'com.wiswork/mutated': execution.mutated === true,
    },
  }
}
