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

export type PresentationStage =
  | 'project_recovery'
  | 'planning'
  | 'production'
  | 'import'
  | 'review'
  | 'evidence'
  | 'sources'
  | 'editing'
  | 'research'
  | 'style'
  | 'assets'
  | 'baseline'
  | 'delivery'
  | 'checkpoint'

const STAGE_LABELS: Readonly<Record<PresentationStage, string>> = {
  project_recovery: '项目恢复',
  planning: '演示文稿计划',
  production: '逐页制作',
  import: '页面导入',
  review: '页面审查',
  evidence: '内容证据',
  sources: '参考资料',
  editing: '页面修改',
  research: '资料研究',
  style: '视觉方向',
  assets: '素材准备',
  baseline: '文档基线',
  delivery: '文件交付',
  checkpoint: '保存点',
}

// Each pattern enumerates known actions and targets; unknown writes remain generic.
const PRESENTATION_STAGES: ReadonlyArray<readonly [PresentationStage, RegExp]> = [
  ['project_recovery', /^(?:restore|resume)_presentation_project$/],
  [
    'planning',
    /^(?:(?:save|read)_presentation_plan|read_presentation_domain_skill|read_presentation_preference_candidates|save_presentation_preference|list_presentation_preferences|delete_presentation_preference|import_presentation_preference)$/,
  ],
  [
    'style',
    /^(?:(?:save|read)_presentation_brand_kit|list_presentation_brand_kits|inspect_slide_masters|edit_slide_master|edit_slide_master_xml)$/,
  ],
  [
    'research',
    /^(?:web_search|web_fetch|(?:build|read|export)_research_ledger|list_research_ledgers)$/,
  ],
  ['assets', /^(?:image_search|insert-image)$/],
  [
    'sources',
    /^(?:list_presentation_attachments|read_presentation_attachment|audit_presentation_sources)$/,
  ],
  [
    'baseline',
    /^(?:read_presentation_baseline(?:_page|_complex_page|_chart_source|_notes|_source_links|_rich_text)?|check_presentation_baseline(?:_windows)?|list_slide_shapes|read_slide_text|read_presentation_page|read_presentation_page_geometry)$/,
  ],
  ['checkpoint', /^(?:save|read)_presentation_page_backup$/],
  [
    'delivery',
    /^(?:export_presentation_pdf|export_current_presentation_bundle|restore_presentation_delivery_bundle)$/,
  ],
  [
    'import',
    /^(?:prepare_presentation_production_import|import_generated_presentation|import_presentation_production|read_presentation_import_status|read_presentation_production_import_status|reconcile_presentation_production_import)$/,
  ],
  [
    'production',
    /^(?:(?:start|read|pause|resume|cancel)_presentation_production_job|(?:start|run|read)_presentation_production|compile_deck_with_pptxgenjs|rebuild_presentation_page|read_presentation_page_artifact|add_slide_ir_objects)$/,
  ],
  [
    'review',
    /^(?:capture_presentation_page_qa|record_presentation_page_review|read_presentation_qa|read_presentation_page_reviews|compare_presentation_page_structure|verify_slides|screenshot_slide|capture_existing_presentation_change|record_existing_presentation_change_review|capture_existing_presentation_batch_page|record_existing_presentation_batch_page_review|(?:capture|record)_existing_presentation_image_review|(?:capture|record)_existing_presentation_page_change|(?:list_presentation_review_comments|add_presentation_review_comment|resolve_presentation_review_comment))$/,
  ],
  [
    'evidence',
    /^(?:(?:read|export)_presentation_delivery_report|record_presentation_issue_action|(?:read|record)_presentation_claim_review|read_presentation_claim_evidence|check_presentation_page_content)$/,
  ],
  [
    'editing',
    /^(?:replace_presentation_page|list_presentation_changes|list_existing_presentation_changes|edit_existing_presentation_(?:text|text_range|geometry|table_cell|batch|table_batch)|(?:inspect|undo|resume|reapply|release)_existing_presentation_change|replace_existing_presentation_image|(?:inspect|resume|undo|reapply)_existing_presentation_image_change|(?:stage|reapply|inspect|reconcile|resume|commit|discard|undo|release)_existing_presentation_page_change|reconcile_pending_existing_presentation_page_change|prepare_existing_presentation_(?:composite_revision|text_revision|image_revision|original_page_restore)|(?:inspect|resume|undo|reapply|release)_existing_presentation_batch|(?:stage|inspect|reconcile|resume|discard|commit|undo)_presentation_page_replacement|(?:read|undo|inspect|resume)_presentation_(?:geometry|text)_change|(?:read|inspect|resume|undo)_presentation_image_replacement|edit_presentation_page_(?:text|geometry)|replace_presentation_page_image|edit_slide_(?:text|xml|chart)|update_slide_chart_values|(?:inspect|resume|undo|release|reapply)_slide_chart_values_change|duplicate_slide|execute_office_js)$/,
  ],
]

export function acpPresentationStageLabel(stage: PresentationStage): string {
  return STAGE_LABELS[stage]
}

export function acpPresentationStage(name: string): PresentationStage | undefined {
  return PRESENTATION_STAGES.find(([, pattern]) => pattern.test(name))?.[0]
}

/** User-facing activity only. Project completion still comes from persisted receipts. */
export function acpToolActivity(
  name: string,
  state: 'running' | 'complete' | 'error',
  presentationContext = true,
): string {
  const key = presentationContext ? acpPresentationStage(name) : undefined
  const stage = key ? acpPresentationStageLabel(key) : undefined
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
