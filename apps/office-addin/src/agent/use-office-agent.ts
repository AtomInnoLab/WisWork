import {
  suspendToolExecution,
  parseAgentResumeMessages,
  encodeOfficeScreenshotResult,
  type AgentMessage,
  type AgentSkill,
  type AgentToolCall,
  type ToolExecution,
  type ToolExecutionOutcome,
} from '@wiswork/agent-core'
import { acpToolActivity, createAgentHarness } from '@wiswork/agent-harness'
import { presentationRecoveryReceiptFeedback } from './presentation-recovery-feedback.js'
import type { OfficePowerPointVisualReviewer } from '../skills/powerpoint/powerpoint-verification.js'
import { withPrefetchedPowerPointImage } from '../skills/powerpoint/powerpoint-import-media.js'
import { useSyncExternalStore } from 'react'
import type { OfficeHost } from '../office-document.js'
import type {
  OfficeProposal,
  ProposalController,
  StructuredProposal,
  StructuredProposalController,
} from './proposal-controller.js'
import {
  appendPresentationEvent,
  boundedText,
  emptyPresentationTimeline,
  replacePresentationEvent,
  type OfficeClarificationQuestion,
  type OfficePresentationTimeline,
  type ProposalPresentationEvent,
  type ToolPresentationEvent,
} from './presentation-state.js'
import {
  isDiagnosticToolError,
  type OfficeDiagnostics,
  type PresentationDiagnosticContext,
} from '../diagnostics/office-diagnostics.js'
import type { PresentationVerificationStringKey } from '@wiswork/i18n'
import { MAX_OBSERVED_TOOL_CALLS, type OfficeAgentTransport } from './transport.js'

export type AgentSessionStatus = 'idle' | 'working' | 'done' | 'cancelled' | 'error'

export interface OfficeAgentSnapshot {
  assistantText: string
  activity: string
  busy: boolean
  applying: boolean
  status: AgentSessionStatus
  error?: string
  errorMessage?: string
  retryable: boolean
  recoveryAvailable?: boolean
  proposal?: OfficeProposal | StructuredProposal
  questionnaire?: readonly OfficeClarificationQuestion[]
  timeline: OfficePresentationTimeline
}

export interface OfficeAgentSession {
  snapshot(): OfficeAgentSnapshot
  subscribe(listener: () => void): () => void
  send(instruction: string): void
  reviseDesignContract?(designMd: string): void
  stop(): void
  confirm(id: string): Promise<void>
  reject(): void
  newTask(): void
  retry(): void
  resumeInterrupted?(): Promise<void>
  answerQuestionnaire?(answers: string): void
  skipQuestionnaire?(): void
  logout(): void
  authenticationLost(): void
  dispose(): void
}

interface SafeSessionError {
  code: string
  message: string
  retryable: boolean
}

const confirmationErrors: Readonly<Record<string, SafeSessionError>> = Object.freeze({
  proposal_missing: {
    code: 'proposal_missing',
    message: 'This proposed change is no longer available.',
    retryable: false,
  },
  proposal_stale: {
    code: 'proposal_stale',
    message: '文档内容已发生变化，刚才的修改未应用。',
    retryable: true,
  },
  presentation_lock_review_pending: {
    code: 'presentation_lock_review_pending',
    message: '正在核对锁页状态，请等待核对完成。',
    retryable: false,
  },
  presentation_lock_review_unavailable: {
    code: 'presentation_lock_review_unavailable',
    message: '无法可靠核对锁页状态，本次修改未应用。请恢复连接或页面身份后重新生成提案。',
    retryable: false,
  },
  presentation_lock_review_stale: {
    code: 'presentation_lock_review_stale',
    message: '锁页、计划或页面身份已变化，本次修改未应用。请重新读取并生成提案。',
    retryable: false,
  },
  presentation_existing_backup_capacity: {
    code: 'presentation_existing_backup_capacity',
    message:
      '保存点备份容量已满。请在修改差异与撤销中释放已结束记录的备份，再重新发起修改。释放后该记录无法重新应用；不会自动重试。',
    retryable: false,
  },
  office_write_failed: {
    code: 'office_write_failed',
    message: 'The approved change could not be applied.',
    retryable: false,
  },
  office_overwrite_required: {
    code: 'office_overwrite_required',
    message: 'The target cells contain data. Choose an empty range or explicitly allow overwrite.',
    retryable: false,
  },
  office_verify_failed: {
    code: 'office_verify_failed',
    message: 'The approved change could not be verified.',
    retryable: false,
  },
  office_recovery_failed: {
    code: 'office_recovery_failed',
    message: 'The document could not be restored after the failed change.',
    retryable: false,
  },
  office_concurrent_change: {
    code: 'office_concurrent_change',
    message: 'The document changed during the operation. Inspect it before trying again.',
    retryable: false,
  },
  office_state_uncertain: {
    code: 'office_state_uncertain',
    message:
      'The change may be partially applied. Wait for reconciliation; if editing stays blocked, reload the document before trying again.',
    retryable: false,
  },
  office_applied_unverified: {
    code: 'office_applied_unverified',
    message:
      'The change may have been applied, but verification was unavailable. Inspect the document before continuing.',
    retryable: false,
  },
})

const runErrors: Readonly<Record<string, SafeSessionError>> = Object.freeze({
  session_expired: {
    code: 'session_expired',
    message:
      'The connection authorization expired. Reconnect to WisWork PC, then continue the unfinished work; existing changes are preserved.',
    retryable: false,
  },
  auth_required: {
    code: 'auth_required',
    message: 'Sign in to WisWork PC, reconnect, and try again.',
    retryable: false,
  },
  network_error: {
    code: 'network_error',
    message: 'The connection was interrupted. Check WisWork PC and try again.',
    retryable: true,
  },
  provider_unavailable: {
    code: 'provider_unavailable',
    message: 'The Agent service is temporarily unavailable. Your progress is saved.',
    retryable: true,
  },
  request_timeout: {
    code: 'request_timeout',
    message: 'The Agent took too long to respond. Try again.',
    retryable: true,
  },
  presentation_run_checkpoint_unavailable: {
    code: 'presentation_run_checkpoint_unavailable',
    message:
      'The tool may have changed this presentation, but its completion record could not be saved. Inspect the document before starting another task.',
    retryable: false,
  },
  transport_stream_budget_exceeded: {
    code: 'transport_stream_budget_exceeded',
    message:
      'The task reached its response limit. Existing changes are preserved; continue the unfinished pages in a new request.',
    retryable: false,
  },
})

const safeConfirmationError = (error: unknown): SafeSessionError => {
  const code = error instanceof Error ? error.message : ''
  if (/^office_recovery_failed:word_[a-z_]+$/.test(code))
    return {
      code,
      message: `The document could not be restored after the failed change (${code.split(':')[1]}).`,
      retryable: false,
    }
  return (
    confirmationErrors[code] ?? {
      code: 'office_write_failed',
      message: 'The approved change could not be applied.',
      retryable: false,
    }
  )
}

const safeRunError = (error: string): SafeSessionError => {
  const aliases: Readonly<Record<string, string>> = {
    transport_timeout: 'request_timeout',
    transport_auth: 'auth_required',
    transport_http_401: 'auth_required',
    transport_http_403: 'auth_required',
    transport_http_408: 'request_timeout',
    transport_http_429: 'provider_unavailable',
    transport_http_500: 'provider_unavailable',
    transport_http_502: 'provider_unavailable',
    transport_http_503: 'provider_unavailable',
    transport_http_504: 'request_timeout',
    transport_network: 'network_error',
  }
  const code = Object.hasOwn(aliases, error) ? aliases[error]! : error
  return (
    (Object.hasOwn(runErrors, code) ? runErrors[code] : undefined) ?? {
      code: 'agent_run_failed',
      message: 'The Agent could not complete this request. Try again.',
      retryable: true,
    }
  )
}

function toolActivity(name: string, state: 'running' | 'complete' | 'error'): string {
  const labels: Readonly<Record<string, string>> = {
    web_search: '网页搜索',
    web_fetch: '读取网页',
    image_search: '图片搜索',
    insert_web_image: '插入网络图片',
    'insert-image': '插入图片',
    plan_deck: '规划演示文稿',
    screenshot_slide: '检查幻灯片',
    verify_slides: '验证演示文稿',
    inspect_slide_masters: '检查母版',
    list_slide_shapes: '读取页面元素',
    read_slide_text: '读取幻灯片内容',
    edit_slide_text: '编辑幻灯片文字',
    edit_slide_xml: '编辑幻灯片版式',
    edit_slide_chart: '编辑图表',
    duplicate_slide: '复制幻灯片',
    execute_office_js: '制作幻灯片',
  }
  const label = labels[name]
  if (label)
    return state === 'running' ? label : state === 'error' ? label + '未完成' : label + '完成'
  const attachment = name === 'read' || name === 'bash'
  const read = /^(?:get_|read_|list_|search_|screenshot_|verify_)/.test(name)
  const action = attachment ? '处理附件' : read ? '读取内容' : '准备修改'
  return state === 'running'
    ? '正在' + action + '…'
    : state === 'error'
      ? action + '未完成'
      : '已' + action
}

const DIAGNOSTIC_TOOL_ERRORS = new Set([
  'cancelled',
  'design_contract_review_required',
  'design_contract_prototype_required',
  'design_contract_production_incomplete',
  'design_contract_verification_failed',
  'design_contract_visual_review_failed',
  'design_contract_invalid_status',
  'design_contract_review_not_pending',
  'design_contract_acceptance_mismatch',
  'design_contract_screenshot_required',
  'image_fetch_unavailable',
  'image_limit',
  'image_mime_unsupported',
  'invalid_image',
  'invalid_tool_input',
  'office_api_unsupported',
  'office_read_failed',
  'office_screenshot_unavailable',
  'office_overwrite_required',
  'office_recovery_failed',
  'office_concurrent_change',
  'office_state_uncertain',
  'office_verify_failed',
  'office_write_failed',
  'proposal_missing',
  'proposal_stale',
])

const IMAGE_FAILURE_MESSAGES: Readonly<Record<string, string>> = {
  image_fetch_unavailable: '图片暂时无法获取',
  image_limit: '图片超过大小限制',
  image_mime_unsupported: '图片格式不受支持',
  invalid_image: '图片数据无效',
  invalid_tool_input: '图片插入参数无效',
  cancelled: '图片操作已取消',
}

const AUTOMATIC_POWERPOINT_MUTATION_TOOLS = new Set([
  'set_slide_background',
  'edit_slide_text',
  'execute_office_js',
  'edit_slide_xml',
  'edit_slide_chart',
  'edit_slide_master',
  'edit_slide_master_xml',
  'duplicate_slide',
  'insert-image',
  'insert_web_image',
])

const AUTOMATIC_RECOVERY_DELAYS_MS = [2_000, 8_000] as const
const AUTOMATIC_RECOVERY_ERRORS = new Set([
  'network_error',
  'provider_unavailable',
  'request_timeout',
])
const RECOVERY_INSTRUCTION =
  'Resume the interrupted task from the current Office document state. Inspect the document and active design contract first, preserve completed work, and continue only unfinished or failed steps. Do not repeat a successful write. Verify the final result before completing.'

function diagnosticToolError(output: string): string {
  if (output.startsWith('design_contract_visual_review_failed:'))
    return 'design_contract_visual_review_failed'
  if (output === 'raw_office_program_invalid') return 'invalid_tool_input'
  const safe = (value: string) =>
    isDiagnosticToolError(value) ||
    DIAGNOSTIC_TOOL_ERRORS.has(value) ||
    /^office_recovery_failed:word_[a-z_]+$/.test(value)
  if (safe(output)) return output
  try {
    const parsed = JSON.parse(output) as { error?: unknown; reason?: unknown }
    if (parsed.error === 'office_read_failed' && parsed.reason === 'office_screenshot_unavailable')
      return 'office_screenshot_unavailable'
    return typeof parsed.error === 'string' && safe(parsed.error)
      ? parsed.error
      : 'agent_run_failed'
  } catch {
    return 'agent_run_failed'
  }
}

export function presentationClarificationText(
  question: string,
  translate?: (key: PresentationVerificationStringKey) => string,
): string {
  if (!question || question === 'presentation_scope_required')
    return translate?.('clarify') ?? 'More information needed'
  return boundedText(question)
}

function screenshotWaitingDiagnostic(
  call: AgentToolCall,
  execution: ToolExecution,
): string | undefined {
  if (
    execution.isError ||
    execution.mutated !== false ||
    ![
      'capture_presentation_page_qa',
      'record_presentation_page_review',
      'compare_presentation_page_structure',
    ].includes(call.name) ||
    new TextEncoder().encode(execution.output).byteLength > 4096
  )
    return undefined
  try {
    const value = JSON.parse(execution.output) as Record<string, unknown>
    const pageId = call.input.page_id
    if (
      !value ||
      typeof value !== 'object' ||
      Array.isArray(value) ||
      Object.keys(value).length !== 4 ||
      value.status !== 'waiting_screenshot' ||
      value.retryable !== true ||
      typeof pageId !== 'string' ||
      !/^[a-zA-Z0-9_-]{1,80}$/.test(pageId) ||
      value.pageId !== pageId ||
      typeof value.hostSlideId !== 'string' ||
      !value.hostSlideId.trim() ||
      value.hostSlideId.length > 256
    )
      return undefined
    return 'presentation_screenshot_waiting'
  } catch {
    return undefined
  }
}

function presentationDiagnosticContext(
  call: AgentToolCall,
): PresentationDiagnosticContext | undefined {
  const string = (value: unknown) => (typeof value === 'string' ? value : undefined)
  const projectId = string(call.input.project_id)
  const requestId = string(call.input.request_id)
  const pageId = string(call.input.page_id ?? call.input.host_slide_id ?? call.input.slide_id)
  if (!projectId && !requestId && !pageId) return undefined
  return {
    tool_call_id: call.id,
    ...(projectId ? { project_id: projectId } : {}),
    ...(requestId ? { request_id: requestId } : {}),
    ...(pageId ? { page_id: pageId } : {}),
  }
}

export function createOfficeAgentSession(dependencies: {
  host?: OfficeHost
  transport: OfficeAgentTransport
  skill: AgentSkill
  proposals: ProposalController | StructuredProposalController
  diagnostics?: Pick<OfficeDiagnostics, 'startTrace' | 'setTool' | 'record' | 'clear'>
  runCheckpoint?: {
    interrupted: boolean
    scrubFailed?: boolean
    recovery?: {
      runId?: string
      instruction: string
      phase: 'running' | 'tool_pending' | 'tool_completed'
      toolName?: string
      toolCallId?: string
      restartSafe?: boolean
      messages?: AgentMessage[]
      importReceipt?: {
        state: 'complete' | 'partial' | 'uncertain'
        completed: number
        total?: number
      }
      changeReceipt?: { total: number; unresolved: number }
    }
    readRecovery?(): NonNullable<
      Parameters<typeof createOfficeAgentSession>[0]['runCheckpoint']
    >['recovery']
    validateDocument?(): Promise<boolean>
    begin(runId: string, instruction: string): Promise<void>
    tool?(
      runId: string,
      phase: 'tool_pending' | 'tool_completed',
      toolName: string,
      mutated?: boolean,
      toolCallId?: string,
    ): Promise<void>
    finish(runId: string): Promise<void>
    conversation?(runId: string, messages: readonly AgentMessage[]): Promise<void>
    adopt?(runId: string, messages: readonly AgentMessage[]): Promise<void>
  }
  presentationText?: (key: PresentationVerificationStringKey) => string
  /**
   * PC-managed PowerPoint autonomy. Ordinary bounded proposals still use the Office
   * validate/write/verify transaction, but do not interrupt the run with UI confirmation.
   * Elevated raw Office proposals are never eligible.
   */
  automaticPowerPointMutations?: boolean
  remoteTools?: {
    setToolHandler?(
      handler:
        | ((call: {
            turnId: string
            callId: string
            generation: number
            toolName: string
            input: Record<string, unknown>
            signal: AbortSignal
          }) => Promise<{ output: string; isError?: boolean }>)
        | undefined,
    ): void
  }
}): OfficeAgentSession {
  const { proposals } = dependencies
  const isAutomaticProposal = (
    proposal: OfficeProposal | StructuredProposal | undefined,
  ): proposal is StructuredProposal =>
    dependencies.automaticPowerPointMutations === true &&
    !dependencies.runCheckpoint &&
    (!proposal || !('lockReview' in proposal) || proposal.lockReview === undefined) &&
    proposal !== undefined &&
    'impact' in proposal &&
    proposal.impact.host.toLowerCase() === 'powerpoint' &&
    typeof proposal.toolName === 'string' &&
    AUTOMATIC_POWERPOINT_MUTATION_TOOLS.has(proposal.toolName)
  const visibleProposal = () => {
    const proposal = proposals.pending()
    return isAutomaticProposal(proposal) ? undefined : proposal
  }
  const presentation = dependencies.skill.presentation as
    | (NonNullable<AgentSkill['presentation']> & {
        setReviewer?: (reviewer: OfficePowerPointVisualReviewer) => void
      })
    | undefined
  presentation?.setReviewer?.({
    review: (request) =>
      new Promise((resolve) => {
        let output = ''
        let settled = false
        const reviewHandle: { current?: { cancel(): void } } = {}
        const unavailable = {
          status: 'cannot_verify' as const,
          failedCheckIds: [],
          observations: [{ code: 'review_unavailable' as const, severity: 'warning' as const }],
          fixIntents: [],
        }
        const finish = (value: Parameters<typeof resolve>[0], cancel = false) => {
          if (settled) return
          settled = true
          clearTimeout(timer)
          request.signal?.removeEventListener('abort', abort)
          if (cancel) reviewHandle.current?.cancel()
          resolve(value)
        }
        const abort = () => finish(unavailable, true)
        const timer = setTimeout(() => finish(unavailable, true), 15_000)
        request.signal?.addEventListener('abort', abort, { once: true })
        if (request.signal?.aborted) abort()
        reviewHandle.current = dependencies.transport.stream(
          {
            system:
              'Review only the supplied PowerPoint screenshots against the bounded check IDs. Return strict JSON matching VisualReviewResult. Do not request tools, infer hidden text, or add targets.',
            messages: [
              {
                role: 'user',
                text: JSON.stringify({ facts: request.facts }),
                images: request.images.map(({ base64, mime }) => ({ base64, mime })),
              },
            ],
            tools: [],
          },
          {
            onDelta: (text) => {
              if (output.length < 64 * 1024) output += text
            },
            onToolCall: () => finish(unavailable, true),
            onDone: () => {
              try {
                finish(JSON.parse(output))
              } catch {
                finish(unavailable)
              }
            },
            onError: () => finish(unavailable),
          },
        )
        if (settled) reviewHandle.current.cancel()
      }),
  })
  const diagnose = (
    action: (diagnostics: NonNullable<typeof dependencies.diagnostics>) => void,
  ) => {
    if (!dependencies.diagnostics) return
    try {
      action(dependencies.diagnostics)
    } catch {
      /* diagnostics never changes an Agent run */
    }
  }
  const listeners = new Set<() => void>()
  let disposed = false
  const pendingAuditedRead = (
    recovery: NonNullable<typeof dependencies.runCheckpoint>['recovery'],
  ) =>
    recovery?.phase === 'tool_pending' &&
    recovery.restartSafe === true &&
    !recovery.messages &&
    Boolean(recovery.toolName) &&
    Boolean(recovery.toolCallId)
  const receiptFeedback = presentationRecoveryReceiptFeedback(dependencies.runCheckpoint?.recovery)
  let state: Omit<OfficeAgentSnapshot, 'proposal'> = {
    assistantText: '',
    activity: '',
    busy: false,
    applying: false,
    status: 'idle',
    retryable: false,
    recoveryAvailable:
      !dependencies.runCheckpoint?.scrubFailed &&
      dependencies.runCheckpoint?.recovery?.restartSafe !== false &&
      (!dependencies.runCheckpoint?.recovery?.messages ||
        Boolean(
          dependencies.runCheckpoint.recovery.runId &&
          dependencies.runCheckpoint.adopt &&
          dependencies.runCheckpoint.conversation &&
          parseAgentResumeMessages(dependencies.runCheckpoint.recovery.messages),
        )) &&
      (dependencies.runCheckpoint?.recovery?.phase === 'running' ||
        pendingAuditedRead(dependencies.runCheckpoint?.recovery) ||
        (dependencies.runCheckpoint?.recovery?.phase === 'tool_completed' &&
          dependencies.runCheckpoint?.recovery?.restartSafe === true)) &&
      Boolean(dependencies.runCheckpoint?.recovery?.instruction),
    timeline: dependencies.runCheckpoint?.interrupted
      ? appendPresentationEvent(emptyPresentationTimeline(), {
          id: 'event-1',
          kind: 'system',
          text: dependencies.runCheckpoint.scrubFailed
            ? '上次运行已中断。旧版检查点中的请求原文仍保留在本 PPTX：清理保存失败。请先保存可写副本并重新打开，期间不能继续该运行。'
            : `上次前台 Agent 运行在面板关闭时中断。${receiptFeedback}${dependencies.runCheckpoint.recovery?.messages ? '已保留完整只读结果，可在核对文档后继续；旧结果代表历史读取，当前状态仍需重新核对。' : dependencies.runCheckpoint.recovery?.phase === 'running' ? '尚未调用工具，可在核对文档后主动重新运行原请求。' : pendingAuditedRead(dependencies.runCheckpoint.recovery) ? '已审计的只读调用尚未完成；核对文档后可主动重新读取，旧结果不会复用。' : dependencies.runCheckpoint.recovery?.phase === 'tool_completed' && dependencies.runCheckpoint.recovery.restartSafe ? '此前仅运行了可重读工具，可在核对文档后主动重新运行原请求。' : '请先核对项目、页面和写入记录；未自动重放写入。'}运行阶段保存在演示文稿设置中，请求与可恢复的读取结果仅保存在本机浏览器。`,
        })
      : emptyPresentationTimeline(),
  }
  let cached: OfficeAgentSnapshot = { ...state, proposal: visibleProposal() }

  const publish = (next: Partial<typeof state> = {}) => {
    if (disposed) return
    state = { ...state, ...next }
    cached = { ...state, proposal: visibleProposal() }
    listeners.forEach((listener) => listener())
  }

  let nextEventId = dependencies.runCheckpoint?.interrupted ? 1 : 0
  let sessionEpoch = 0
  let pendingStart = false
  let activeRunId: string | undefined
  let toolsStarted = false
  let recoveryPending = false
  let checkpointBeginFailed = false
  const readRecovery = () => {
    try {
      return (
        dependencies.runCheckpoint?.readRecovery?.() ??
        (dependencies.runCheckpoint?.readRecovery
          ? undefined
          : dependencies.runCheckpoint?.recovery)
      )
    } catch {
      return undefined
    }
  }
  const safeRecovery = () => {
    const checkpoint = dependencies.runCheckpoint
    if (
      !checkpoint ||
      checkpoint.scrubFailed ||
      (activeRunId && unsettledToolRuns.has(activeRunId))
    )
      return undefined
    if (toolsStarted && !checkpoint.readRecovery) return undefined
    const record = readRecovery()
    if (checkpoint.readRecovery && !record) return undefined
    if (
      !record &&
      !toolsStarted &&
      new TextEncoder().encode(lastInstruction).byteLength <= 8 * 1024
    )
      return { instruction: lastInstruction, phase: 'running' as const }
    if (record?.restartSafe === false) return undefined
    if (
      record?.messages &&
      (!record.runId ||
        !checkpoint.adopt ||
        !checkpoint.conversation ||
        !parseAgentResumeMessages(record.messages))
    )
      return undefined
    if (
      !record?.instruction ||
      new TextEncoder().encode(record.instruction).byteLength > 8 * 1024 ||
      record.importReceipt?.state === 'uncertain' ||
      record.changeReceipt?.total
    )
      return undefined
    return record.phase === 'running' ||
      pendingAuditedRead(record) ||
      (record.phase === 'tool_completed' && record.restartSafe === true)
      ? record
      : undefined
  }
  const unsettledToolRuns = new Set<string>()
  const finishCheckpoint = () => {
    const id = activeRunId
    activeRunId = undefined
    if (id && !unsettledToolRuns.has(id))
      void dependencies.runCheckpoint?.finish(id).catch(() => undefined)
  }
  let activeAssistantId: string | undefined
  let lastInstruction = ''
  let clarificationResolve: ((value: ToolExecution) => void) | undefined
  let questionnaireAnsweredPendingPlan = false
  let runStartedAt = 0
  let automaticRecoveryAttempt = 0
  let recoveryTimer: ReturnType<typeof setTimeout> | undefined
  const staleTools = new Set<string>()
  const toolStartedAt = new Map<string, number>()
  // Canonical observation is authoritative even when a relay execution receipt arrives first.
  const observedTools = new Map<string, 'running' | 'settled'>()
  const recordedToolFailures = new Set<string>()
  const recordToolFailure = (
    callId: string,
    toolName: string,
    errorCode: string,
    durationMs: number,
    error?: unknown,
  ) => {
    if (recordedToolFailures.has(callId) || recordedToolFailures.size >= MAX_OBSERVED_TOOL_CALLS)
      return
    recordedToolFailures.add(callId)
    diagnose((diagnostics) => {
      diagnostics.setTool(toolName)
      diagnostics.record({
        phase: 'tool',
        errorCode,
        durationMs: Math.max(0, durationMs),
        ...(error === undefined ? {} : { error }),
      })
    })
  }
  const eventId = () => `event-${++nextEventId}`
  const append = (event: Parameters<typeof appendPresentationEvent>[1]) => {
    state = { ...state, timeline: appendPresentationEvent(state.timeline, event) }
  }
  const replace = (id: string, update: Parameters<typeof replacePresentationEvent>[2]) => {
    state = { ...state, timeline: replacePresentationEvent(state.timeline, id, update) }
  }
  const closeAssistantSegment = () => {
    if (!activeAssistantId) return
    replace(activeAssistantId, (event) => ({ ...event, streaming: false }))
    activeAssistantId = undefined
  }
  const pendingProposalEvent = () =>
    [...state.timeline]
      .reverse()
      .find(
        (event): event is ProposalPresentationEvent =>
          event.kind === 'proposal' && event.state === 'pending',
      )
  const appendPendingProposal = () => {
    const proposal = visibleProposal()
    const existing =
      proposal &&
      state.timeline.find((event) => event.kind === 'proposal' && event.proposal.id === proposal.id)
    if (existing?.kind === 'proposal' && existing.state === 'pending') {
      replace(existing.id, (event) =>
        event.kind === 'proposal' ? { ...event, proposal: proposal! } : event,
      )
      return
    }
    if (
      proposal &&
      !state.timeline.some(
        (event) => event.kind === 'proposal' && event.proposal.id === proposal.id,
      )
    ) {
      append({ id: eventId(), kind: 'proposal', proposal, state: 'pending' })
    }
  }

  const finalProposalExecution = async (
    proposalId: string,
    initial: ToolExecution,
    toolName: string,
  ): Promise<ToolExecution> => {
    const decision = await proposals.waitForDecision(proposalId)
    if (decision.status === 'confirmed') {
      const postWrite = decision.postWrite
      if (postWrite) {
        const pages = postWrite.status === 'captured' ? postWrite.pages : []
        return {
          output: JSON.stringify({
            proposalId,
            status: 'applied',
            qaPassed: false,
            visualReview: 'pending',
            instruction:
              'Capture and inspect each affected page with a one-page review tool before recording a visual judgment.',
            postWrite:
              postWrite.status === 'captured'
                ? {
                    status: 'captured',
                    pages: pages.map(({ slideId, digest }) => ({ slideId, digest })),
                  }
                : postWrite,
          }),
          mutated: true,
          summary: 'Applied approved change; visual review pending',
          display: pages.length
            ? {
                kind: 'images',
                items: pages.map((page) => ({
                  url: `data:image/png;base64,${page.pngBase64}`,
                  title: page.slideId,
                })),
              }
            : undefined,
        }
      }
      return {
        output: JSON.stringify({ proposalId, status: 'applied' }),
        mutated: true,
        summary: 'Applied approved change',
      }
    }
    if (decision.status === 'applied_unverified') {
      const writePending = decision.safeCode === 'office_write_pending'
      return {
        output: JSON.stringify({
          proposalId,
          status: writePending ? 'write_pending' : 'applied_unverified',
          ...(writePending
            ? {
                safeCode: 'office_write_pending',
                instruction:
                  'The write outcome is unresolved. Do not claim success and do not attempt another edit.',
              }
            : {}),
        }),
        mutated: true,
        summary: writePending
          ? (dependencies.presentationText?.('write_pending_quarantined') ??
            'Write may still be running; further edits are frozen pending reconciliation or reload.')
          : 'Applied; verification unavailable',
        stopToolBatch: true,
      }
    }
    if (decision.status === 'failed') {
      const stale =
        decision.error === 'proposal_stale' ||
        ['presentation_lock_review_stale', 'presentation_lock_review_unavailable'].includes(
          decision.error,
        )
      if (stale) staleTools.add(toolName)
      return {
        output: JSON.stringify({
          proposalId,
          status: 'failed',
          error: decision.error,
          ...(decision.errorLocation ? { errorLocation: decision.errorLocation } : {}),
          instruction:
            decision.error === 'proposal_stale'
              ? 'Do not retry this write in the current turn.'
              : decision.error === 'office_verify_failed'
                ? 'Some operations may already have been applied. Read the current slide and shapes before making a small corrective edit; do not repeat the whole batch. If fontFamily failed, Office did not confirm that font: preserve the current family and apply size/color separately, then screenshot and review the result.'
                : undefined,
        }),
        isError: true,
        // Verification is after execute: keep the repair/final-review loop active
        // even when this was the only write attempted in the turn.
        mutated: decision.error === 'office_verify_failed',
        summary: 'Approved change failed',
        stopToolBatch: stale,
      }
    }
    return {
      output: JSON.stringify({
        proposalId,
        status: decision.status === 'rejected' ? 'user_rejected_change' : 'cancelled',
        instruction: 'Do not retry this write in the current turn.',
      }),
      isError: decision.status === 'cancelled',
      mutated: false,
      summary: decision.status === 'rejected' ? 'Change rejected' : initial.summary,
      stopToolBatch: decision.status === 'rejected',
    }
  }

  const sessionSkill: AgentSkill = {
    ...dependencies.skill,
    reviewFinalResponse(context) {
      if (questionnaireAnsweredPendingPlan)
        return '[System correction] The questionnaire is complete. Continue the WisWork Slides workflow with plan_deck, research, implementation, screenshot inspection, repair, and verify_slides in this same run.'
      return dependencies.skill.reviewFinalResponse?.(context)
    },
    async executeTool(call, signal): Promise<ToolExecutionOutcome> {
      toolsStarted = true
      if (call.name === 'ask_clarification') {
        if (clarificationResolve)
          return {
            output: 'questionnaire_in_progress',
            isError: true,
            mutated: false,
            summary: 'Questionnaire already active',
          }
        const raw = Array.isArray(call.input.questions) ? call.input.questions : []
        const questions: OfficeClarificationQuestion[] = raw
          .slice(0, 1)
          .map((value, index) => {
            const item =
              value && typeof value === 'object' ? (value as Record<string, unknown>) : {}
            return {
              id: typeof item.id === 'string' ? item.id.slice(0, 40) : `q${index + 1}`,
              label: typeof item.label === 'string' ? boundedText(item.label).slice(0, 300) : '',
              ...(typeof item.description === 'string'
                ? { description: boundedText(item.description).slice(0, 300) }
                : {}),
              options: Array.isArray(item.options)
                ? item.options
                    .flatMap((option) => {
                      if (typeof option === 'string') return [boundedText(option).slice(0, 120)]
                      if (!option || typeof option !== 'object' || Array.isArray(option)) return []
                      const label = (option as Record<string, unknown>).label
                      return typeof label === 'string' ? [boundedText(label).slice(0, 120)] : []
                    })
                    .filter(Boolean)
                    .slice(0, 5)
                : [],
            }
          })
          .filter((question) => question.label && question.options.length >= 2)
        if (questions.length < 1)
          return {
            output: 'invalid_tool_input',
            isError: true,
            mutated: false,
            summary: 'Questionnaire input invalid',
          }
        return suspendToolExecution(
          new Promise<ToolExecution>((resolve) => {
            clarificationResolve = resolve
            publish({ questionnaire: questions, activity: '等待你完成问卷' })
            signal?.addEventListener(
              'abort',
              () => {
                if (clarificationResolve === resolve) {
                  clarificationResolve = undefined
                  publish({ questionnaire: undefined })
                }
                resolve({
                  output: 'cancelled',
                  isError: true,
                  mutated: false,
                  summary: 'Questionnaire cancelled',
                })
              },
              { once: true },
            )
          }),
        )
      }
      if (staleTools.has(call.name)) {
        return {
          output: JSON.stringify({
            status: 'failed',
            error: 'proposal_stale',
            instruction: 'Do not retry this write in the current turn.',
          }),
          isError: true,
          mutated: false,
          summary: 'Write blocked after stale validation',
          stopToolBatch: true,
        }
      }
      const runId = activeRunId
      const epoch = sessionEpoch
      const currentRun = () => epoch === sessionEpoch && runId === activeRunId && !disposed
      const checkpointCompleted = async (mutated: boolean): Promise<boolean> => {
        if (runId && !currentRun()) unsettledToolRuns.delete(runId)
        if (runId && currentRun() && dependencies.runCheckpoint?.tool)
          try {
            await dependencies.runCheckpoint.tool(
              runId,
              'tool_completed',
              call.name,
              mutated,
              call.id,
            )
            unsettledToolRuns.delete(runId)
          } catch {
            return false
          }
        return true
      }
      const checkpointed = async (result: ToolExecution): Promise<ToolExecution> =>
        (await checkpointCompleted(result.mutated === true))
          ? result
          : {
              output: JSON.stringify({ error: 'presentation_run_checkpoint_unavailable' }),
              isError: true,
              mutated: result.mutated,
              summary: 'Run checkpoint unavailable',
              fatalError: 'presentation_run_checkpoint_unavailable',
            }
      if (runId && dependencies.runCheckpoint?.tool) {
        try {
          await dependencies.runCheckpoint.tool(runId, 'tool_pending', call.name, false, call.id)
          unsettledToolRuns.add(runId)
        } catch {
          return {
            output: JSON.stringify({ error: 'presentation_run_checkpoint_unavailable' }),
            isError: true,
            mutated: false,
            summary: 'Run checkpoint unavailable',
            stopToolBatch: true,
            fatalError: 'presentation_run_checkpoint_unavailable',
          }
        }
      }
      if (!currentRun())
        return {
          output: JSON.stringify({ error: 'presentation_run_cancelled' }),
          isError: true,
          mutated: false,
          summary: 'Run cancelled',
          stopToolBatch: true,
        }
      const outcome = await dependencies.skill.executeTool(call, signal)
      if ('kind' in outcome && outcome.kind === 'tool-execution-suspension')
        return suspendToolExecution(outcome.result.then(checkpointed))
      if (call.name === 'plan_deck' && !outcome.isError) questionnaireAnsweredPendingPlan = false
      const proposal = proposals.pending()
      if (!proposal) return checkpointed(outcome)
      const final = finalProposalExecution(proposal.id, outcome, call.name).then(checkpointed)
      if (isAutomaticProposal(proposal)) {
        void proposals.confirm(proposal.id).catch(() => undefined)
      }
      return suspendToolExecution(final)
    },
  }
  dependencies.transport.setToolActivityHandler?.((activity) => {
    if (disposed || !state.busy) return
    const existing = state.timeline.find(
      (event) => event.kind === 'tool' && event.callId === activity.callId,
    )
    const summary = toolActivity(activity.toolName, activity.state)
    if (activity.state === 'running') {
      if (observedTools.has(activity.callId) || observedTools.size >= MAX_OBSERVED_TOOL_CALLS)
        return
      observedTools.set(activity.callId, 'running')
      if (existing) return
      closeAssistantSegment()
      diagnose((diagnostics) => diagnostics.setTool(activity.toolName))
      append({
        id: eventId(),
        kind: 'tool',
        callId: activity.callId,
        name: activity.toolName,
        summary,
        state: 'running',
        ...(activity.query ? { output: activity.query } : {}),
      })
    } else {
      if (!existing || existing.kind !== 'tool' || observedTools.get(activity.callId) !== 'running')
        return
      observedTools.set(activity.callId, 'settled')
      const imageError =
        activity.state === 'error' &&
        (activity.toolName === 'insert_web_image' || activity.toolName === 'insert-image') &&
        typeof activity.summary === 'string' &&
        Object.hasOwn(IMAGE_FAILURE_MESSAGES, activity.summary)
          ? activity.summary
          : undefined
      if (imageError)
        recordToolFailure(
          activity.callId,
          activity.toolName,
          imageError,
          Date.now() - activity.startedAt,
        )
      replace(existing.id, (event) =>
        event.kind !== 'tool'
          ? event
          : {
              ...event,
              state: activity.state,
              summary:
                summary +
                (activity.resultCount === undefined ? '' : ` · ${activity.resultCount} 条结果`),
              durationMs: Math.max(0, Date.now() - activity.startedAt),
              output: imageError
                ? `${IMAGE_FAILURE_MESSAGES[imageError]}（${imageError}）`
                : activity.query
                  ? [activity.query, activity.summary].filter(Boolean).join('\n')
                  : (event.output ?? activity.summary ?? ''),
              ...(activity.display ? { display: activity.display } : {}),
            },
      )
    }
    publish({ activity: summary })
  })
  dependencies.remoteTools?.setToolHandler?.(async (call) => {
    if (disposed || call.signal.aborted) return { output: 'tool_cancelled', isError: true }
    const definition = sessionSkill.tools.find((tool) => tool.name === call.toolName)
    if (!definition) return { output: 'unknown_tool', isError: true }
    const epoch = sessionEpoch
    const current = () => !disposed && epoch === sessionEpoch
    closeAssistantSegment()
    diagnose((diagnostics) => diagnostics.setTool(call.toolName))
    const existing = state.timeline.find(
      (event) => event.kind === 'tool' && event.callId === call.callId,
    )
    const presentationId = existing?.id ?? eventId()
    const startedAt = Date.now()
    const runningSummary = toolActivity(call.toolName, 'running')
    if (!existing)
      append({
        id: presentationId,
        kind: 'tool',
        callId: call.callId,
        name: boundedText(call.toolName),
        summary: runningSummary,
        state: 'running',
      })
    publish({ activity: runningSummary })
    const invalidateRemoteProposal = () => {
      if (!current()) return
      proposals.newTurn()
      replace(presentationId, (event) =>
        event.kind === 'tool' && event.state === 'running'
          ? { ...event, summary: toolActivity(call.toolName, 'error'), state: 'error' }
          : event,
      )
      publish()
    }
    call.signal.addEventListener('abort', invalidateRemoteProposal, { once: true })
    try {
      const outcome = await sessionSkill.executeTool(
        withPrefetchedPowerPointImage({
          id: call.callId,
          invocationId: `${call.turnId}:${call.callId}`,
          name: call.toolName,
          input: call.input,
        }),
        call.signal,
      )
      let settled =
        'kind' in outcome && outcome.kind === 'tool-execution-suspension'
          ? await Promise.race([
              outcome.result,
              new Promise<never>((_, reject) => {
                if (call.signal.aborted) reject(new DOMException('Aborted', 'AbortError'))
                else
                  call.signal.addEventListener(
                    'abort',
                    () => reject(new DOMException('Aborted', 'AbortError')),
                    { once: true },
                  )
              }),
            ])
          : outcome
      let screenshotOutput: string | undefined
      if (call.toolName === 'screenshot_slide' && !settled.isError) {
        try {
          screenshotOutput = encodeOfficeScreenshotResult(settled.output, settled.modelContent)
        } catch {
          settled = { ...settled, output: 'office_screenshot_unavailable', isError: true }
        }
      }
      const finishedSummary = toolActivity(call.toolName, settled.isError ? 'error' : 'complete')
      if (!current() || call.signal.aborted) return { output: 'tool_cancelled', isError: true }
      const observed = observedTools.has(call.callId)
      replace(presentationId, (event) =>
        event.kind === 'tool' && event.state === 'running'
          ? {
              ...event,
              summary: observed ? event.summary : finishedSummary,
              state: observed ? event.state : settled.isError ? 'error' : 'complete',
              durationMs: Date.now() - startedAt,
              output: boundedText(settled.output),
              ...(settled.display ? { display: settled.display } : {}),
            }
          : event,
      )
      publish(observed ? {} : { activity: finishedSummary })
      if (settled.isError) {
        const errorCode = diagnosticToolError(settled.output)
        recordToolFailure(
          call.callId,
          call.toolName,
          errorCode,
          Date.now() - startedAt,
          settled.diagnosticError,
        )
      }
      return {
        output: screenshotOutput ?? settled.output,
        ...(settled.isError ? { isError: true } : {}),
      }
    } catch {
      if (!current()) return { output: 'tool_cancelled', isError: true }
      const failedSummary = toolActivity(call.toolName, 'error')
      const observed = observedTools.has(call.callId)
      replace(presentationId, (event) =>
        event.kind === 'tool' && event.state === 'running'
          ? {
              ...event,
              summary: observed ? event.summary : failedSummary,
              state: observed ? event.state : 'error',
              durationMs: Date.now() - startedAt,
            }
          : event,
      )
      publish(observed ? {} : { activity: failedSummary })
      recordToolFailure(
        call.callId,
        call.toolName,
        call.signal.aborted ? 'cancelled' : 'tool_execution_failed',
        Date.now() - startedAt,
      )
      if (call.signal.aborted) proposals.newTurn()
      return { output: 'tool_execution_failed', isError: true }
    } finally {
      call.signal.removeEventListener('abort', invalidateRemoteProposal)
    }
  })
  const clearConversation = () => {
    checkpointBeginFailed = false
    activeAssistantId = undefined
    questionnaireAnsweredPendingPlan = false
    state = {
      ...state,
      assistantText: '',
      activity: '',
      busy: false,
      applying: false,
      status: 'idle',
      error: undefined,
      errorMessage: undefined,
      retryable: false,
      recoveryAvailable: false,
      timeline: emptyPresentationTimeline(),
      questionnaire: undefined,
    }
  }

  const harness = createAgentHarness({
    transport: dependencies.transport,
    skill: sessionSkill,
    systemSuffix: () =>
      resumingReadConversation
        ? '\nThis run resumes a saved read-only conversation. Restored tool results are historical observations. Revalidate relevant live document/project state before making changes or claiming its current state; do not treat cached results as proof that the document is unchanged.'
        : resumingPendingRead
          ? '\nAn audited read-only tool was interrupted before its result was saved. No result from that call was restored. Re-read the live document/project state and reestablish any session-only baseline before making changes or claiming a result.'
          : '',
    events: {
      onPresentationClarify: ({ question }) =>
        append({
          id: eventId(),
          kind: 'system',
          text: presentationClarificationText(question, dependencies.presentationText),
        }),
      onPresentationPlan: ({ steps, requiresConfirmation }) =>
        append({
          id: eventId(),
          kind: 'system',
          text: [
            dependencies.presentationText?.('plan') ?? 'plan',
            ...steps.map(
              (step) =>
                `• ${
                  dependencies.presentationText?.(
                    step === 'presentation_verify_postconditions'
                      ? 'verify_postconditions'
                      : 'apply_bounded_edits',
                  ) ?? step
                }`,
            ),
            ...(requiresConfirmation
              ? [dependencies.presentationText?.('needs_user') ?? 'needs_user']
              : []),
          ].join('\n'),
        }),
      onPresentationCorrection: () =>
        append({
          id: eventId(),
          kind: 'system',
          text: dependencies.presentationText?.('correction') ?? 'correction',
        }),
      onPresentationReceipt: ({ facts }) => dependencies.presentationText?.(facts.status),
      onText: (assistantText) => {
        // Presentation is driven by ACP agent_message_chunk updates below.
        publish({ assistantText: boundedText(assistantText) })
      },
      onToolStart: (call) => {
        toolStartedAt.set(call.id, Date.now())
        diagnose((diagnostics) => {
          const context = presentationDiagnosticContext(call)
          if (context) diagnostics.setTool(call.name, context)
          else diagnostics.setTool(call.name)
        })
      },
      onToolExecuted: (event) => {
        const errorCode = event.execution.isError
          ? diagnosticToolError(event.execution.output)
          : screenshotWaitingDiagnostic(event.call, event.execution)
        if (errorCode) {
          diagnose((diagnostics) => {
            const context = presentationDiagnosticContext(event.call)
            if (context) diagnostics.setTool(event.call.name, context)
            else diagnostics.setTool(event.call.name)
            diagnostics.record({
              phase: 'tool',
              errorCode,
              ...(event.execution.diagnosticError === undefined
                ? {}
                : { error: event.execution.diagnosticError }),
              durationMs: Math.max(
                0,
                Date.now() - (toolStartedAt.get(event.call.id) ?? Date.now()),
              ),
            })
          })
        }
        toolStartedAt.delete(event.call.id)
        appendPendingProposal()
      },
      onTurnEnd: async () => {
        const runId = activeRunId
        const epoch = sessionEpoch
        if (runId && dependencies.runCheckpoint?.conversation) {
          await dependencies.runCheckpoint.conversation(runId, harness.messages)
          if (disposed || epoch !== sessionEpoch || runId !== activeRunId) return
        }
        activeAssistantId = undefined
        publish({ activity: 'Thinking…' })
      },
      onDone: (result) => {
        diagnose((diagnostics) => {
          diagnostics.setTool('agent_run')
          diagnostics.record({
            phase: 'run',
            errorCode: result.cancelled ? 'cancelled' : 'agent_run_completed',
            durationMs: Math.max(0, Date.now() - runStartedAt),
          })
        })
        finishCheckpoint()
        automaticRecoveryAttempt = 0
        if (recoveryTimer) clearTimeout(recoveryTimer)
        recoveryTimer = undefined
        toolStartedAt.clear()
        closeAssistantSegment()
        publish({
          busy: false,
          activity: '',
          status: result.cancelled ? 'cancelled' : 'done',
        })
        if (!result.presentation && result.cancelled)
          append({
            id: eventId(),
            kind: 'system',
            text: dependencies.presentationText?.('cancelled') ?? 'cancelled',
          })
      },
      onError: (error) => {
        const transient = [
          'network_error',
          'provider_unavailable',
          'request_timeout',
          'transport_timeout',
        ].includes(error)
        if (error === 'presentation_run_checkpoint_unavailable') activeRunId = undefined
        else if (!(transient && dependencies.runCheckpoint)) finishCheckpoint()
        const safeError = safeRunError(error)
        const retryable =
          safeError.retryable &&
          (!dependencies.runCheckpoint || (transient && Boolean(safeRecovery())))
        const latestRecovery = readRecovery()
        const baseMessage =
          error === 'presentation_run_checkpoint_unavailable' &&
          latestRecovery?.restartSafe === true
            ? '读取结果未能保存，后续模型请求已停止，运行检查点已保留。请检查本机浏览器存储，重新打开后核对文档并恢复任务。'
            : dependencies.runCheckpoint && transient && !retryable
              ? '运行已中断，检查点已保留。请核对项目、页面和写入记录；不能重跑可能已写入的原请求。'
              : dependencies.runCheckpoint && transient
                ? latestRecovery?.messages
                  ? '服务暂时中断，读取结果已保留。可在核对当前文档后主动继续，已完成读取不自动重放。'
                  : '服务暂时中断，运行阶段已保留。可在核对当前文档后主动重新运行安全请求；未自动重放。'
                : safeError.message
        const message =
          dependencies.runCheckpoint &&
          (transient || error === 'presentation_run_checkpoint_unavailable')
            ? `${baseMessage}影响范围：本次前台运行；后台进度以项目工作台为准。${presentationRecoveryReceiptFeedback(latestRecovery)}${latestRecovery?.restartSafe === true ? '' : '未自动重放写入，请按工作台记录核对后继续。'}`
            : baseMessage
        diagnose((diagnostics) => {
          diagnostics.setTool('agent_run')
          diagnostics.record({
            phase: 'transport',
            errorCode: safeError.code,
            durationMs: Math.max(0, Date.now() - runStartedAt),
          })
        })
        activeAssistantId = undefined
        const delay = AUTOMATIC_RECOVERY_DELAYS_MS[automaticRecoveryAttempt]
        if (
          !dependencies.runCheckpoint &&
          AUTOMATIC_RECOVERY_ERRORS.has(safeError.code) &&
          delay !== undefined &&
          !disposed
        ) {
          automaticRecoveryAttempt += 1
          recoveryTimer = setTimeout(() => {
            recoveryTimer = undefined
            startRun(RECOVERY_INSTRUCTION, undefined, undefined, false, '', true)
          }, delay)
          publish({
            busy: true,
            activity: 'Connection interrupted. Progress saved; recovering…',
            status: 'working',
            error: undefined,
            errorMessage: undefined,
            retryable: false,
          })
          return
        }
        append({
          id: eventId(),
          kind: 'error',
          text: message,
          code: safeError.code,
        })
        publish({
          busy: false,
          activity: '',
          status: 'error',
          error: safeError.code,
          errorMessage: message,
          retryable,
        })
      },
    },
  })

  const unsubscribeAcp = harness.subscribeAcp(({ update }) => {
    if (update.sessionUpdate === 'agent_message_chunk' && update.content.type === 'text') {
      const chunkText = update.content.text
      const messageId = update.messageId ?? activeAssistantId ?? eventId()
      if (activeAssistantId !== messageId) {
        if (activeAssistantId) {
          replace(activeAssistantId, (event) => ({ ...event, streaming: false }))
        }
        activeAssistantId = messageId
        append({
          id: messageId,
          kind: 'assistant',
          text: boundedText(chunkText),
          streaming: true,
        })
      } else {
        replace(messageId, (event) => ({
          ...event,
          text: boundedText(`${event.kind === 'assistant' ? event.text : ''}${chunkText}`),
          streaming: true,
        }))
      }
      return
    }
    if (update.sessionUpdate === 'tool_call') {
      if (activeAssistantId) {
        replace(activeAssistantId, (event) => ({ ...event, streaming: false }))
        activeAssistantId = undefined
      }
      const name = update.name ?? 'tool'
      const summary = acpToolActivity(
        name,
        'running',
        dependencies.host === undefined || dependencies.host === 'powerpoint',
      )
      append({
        id: eventId(),
        kind: 'tool',
        callId: update.toolCallId,
        name: boundedText(name),
        summary,
        state: 'running',
      })
      publish({ activity: summary })
      return
    }
    if (update.sessionUpdate === 'tool_call_update') {
      const tool = [...state.timeline]
        .reverse()
        .find(
          (item): item is ToolPresentationEvent =>
            item.kind === 'tool' && item.callId === update.toolCallId,
        )
      if (!tool) return
      const terminal = update.status === 'failed' ? 'error' : 'complete'
      const summary = acpToolActivity(
        tool.name,
        terminal,
        dependencies.host === undefined || dependencies.host === 'powerpoint',
      )
      replace(tool.id, (item) =>
        item.kind === 'tool'
          ? {
              ...item,
              summary,
              state: terminal,
            }
          : item,
      )
      publish({ activity: summary })
    }
  })

  const unsubscribeProposals = proposals.subscribe(() => {
    const pending = visibleProposal()
    if (pending) {
      appendPendingProposal()
      publish({ activity: 'Waiting for your approval' })
    } else {
      publish()
    }
  })

  let resumingReadConversation = false
  let resumingPendingRead = false
  const startRun = (
    instruction: string,
    messages?: readonly AgentMessage[],
    resumedRunId?: string,
    pendingRead = false,
    displayText?: string,
    recovering = false,
  ) => {
    const value = instruction.trim()
    if (!value || harness.snapshot.busy || pendingStart || state.applying || disposed) return
    if (!recovering) {
      if (recoveryTimer) clearTimeout(recoveryTimer)
      recoveryTimer = undefined
      automaticRecoveryAttempt = 0
      diagnose((diagnostics) => diagnostics.startTrace())
    }
    sessionEpoch += 1
    observedTools.clear()
    recordedToolFailures.clear()
    staleTools.clear()
    checkpointBeginFailed = false
    toolsStarted = false
    runStartedAt = Date.now()
    resumingReadConversation = Boolean(messages)
    resumingPendingRead = pendingRead
    proposals.newTurn()
    if (!recovering) lastInstruction = value
    activeAssistantId = undefined
    if (displayText !== '')
      append({
        id: eventId(),
        kind: messages ? 'assistant' : 'user',
        text: messages ? '从已保存的读取结果继续。' : boundedText(displayText ?? value),
      })
    publish({
      assistantText: '',
      activity: 'Thinking…',
      busy: true,
      status: 'working',
      error: undefined,
      errorMessage: undefined,
      retryable: false,
      recoveryAvailable: false,
    })
    if (!dependencies.runCheckpoint) {
      if (recovering) harness.resume(value)
      else harness.run(value)
      return
    }
    pendingStart = true
    const epoch = sessionEpoch
    const runId = messages && resumedRunId ? resumedRunId : crypto.randomUUID()
    const checkpointStart =
      messages && resumedRunId && dependencies.runCheckpoint.adopt
        ? dependencies.runCheckpoint.adopt(runId, messages)
        : dependencies.runCheckpoint.begin(runId, value)
    void checkpointStart
      .then(async () => {
        if (disposed || epoch !== sessionEpoch) {
          pendingStart = false
          void dependencies.runCheckpoint?.finish(runId).catch(() => undefined)
          return
        }
        activeRunId = runId
        if (messages) {
          pendingStart = false
          if (!harness.resume(messages)) throw new Error('presentation_run_checkpoint_unavailable')
        } else {
          pendingStart = false
          harness.run(value)
        }
      })
      .catch(() => {
        pendingStart = false
        if (disposed || epoch !== sessionEpoch) return
        if (activeRunId && !unsettledToolRuns.has(activeRunId)) activeRunId = undefined
        checkpointBeginFailed = !messages
        const checkpointError = messages
          ? '无法恢复运行上下文，检查点已保留。请核对当前文档和本地存储后重试。'
          : '无法保存运行检查点，请确认文档可保存后重试。'
        append({
          id: eventId(),
          kind: 'error',
          text: checkpointError,
          code: 'presentation_run_checkpoint_unavailable',
        })
        publish({
          busy: false,
          activity: '',
          status: 'error',
          error: 'presentation_run_checkpoint_unavailable',
          errorMessage: checkpointError,
          retryable: messages ? Boolean(safeRecovery()) : true,
        })
      })
  }

  const resumeRecovery = async (interrupted: boolean) => {
    const checkpoint = dependencies.runCheckpoint
    const epoch = sessionEpoch
    let checkpointSnapshot: string | undefined
    try {
      const latest = checkpoint?.readRecovery?.()
      if (checkpointBeginFailed && latest) return
      checkpointSnapshot = JSON.stringify(latest)
    } catch {
      return
    }
    const candidate = () =>
      !interrupted && checkpointBeginFailed && !toolsStarted && !activeRunId && lastInstruction
        ? { instruction: lastInstruction, phase: 'running' as const }
        : safeRecovery()
    const original = candidate()
    const originalSnapshot = original ? JSON.stringify(original) : undefined
    if (
      !checkpoint?.validateDocument ||
      recoveryPending ||
      state.busy ||
      state.applying ||
      disposed ||
      (interrupted ? !state.recoveryAvailable : !state.retryable) ||
      !originalSnapshot
    )
      return
    recoveryPending = true
    try {
      if (!(await checkpoint.validateDocument())) return
      if (JSON.stringify(checkpoint.readRecovery?.()) !== checkpointSnapshot) return
      if (
        epoch !== sessionEpoch ||
        state.busy ||
        state.applying ||
        disposed ||
        (interrupted ? !state.recoveryAvailable : !state.retryable)
      )
        return
      const record = candidate()
      if (record && JSON.stringify(record) === originalSnapshot)
        startRun(
          record.instruction,
          'messages' in record ? record.messages : undefined,
          'runId' in record ? record.runId : undefined,
          pendingAuditedRead(record),
        )
    } catch {
      /* identity or recovery unavailable: do not replay */
    } finally {
      recoveryPending = false
    }
  }

  return {
    snapshot: () => cached,
    subscribe(listener) {
      if (disposed) return () => undefined
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    send(instruction) {
      startRun(instruction)
    },
    reviseDesignContract(designMd) {
      startRun(
        `Revise the active presentation DESIGN.md to exactly the contract below. Keep the existing page plan unless consistency requires a change. Call plan_deck with the revised style and do not edit slides in this turn.\n\n${designMd}`,
        undefined,
        undefined,
        false,
        '更新 DESIGN.md',
      )
    },
    stop() {
      if (disposed) return
      if (recoveryPending) {
        sessionEpoch += 1
        publish({ retryable: false, recoveryAvailable: false, status: 'cancelled' })
        return
      }
      if (pendingStart) {
        sessionEpoch += 1
        publish({ busy: false, activity: '', status: 'cancelled' })
        return
      }
      const wasRecovering = recoveryTimer !== undefined
      if (recoveryTimer) clearTimeout(recoveryTimer)
      recoveryTimer = undefined
      automaticRecoveryAttempt = 0
      const event = pendingProposalEvent()
      if (state.applying) {
        sessionEpoch += 1
        proposals.newTurn()
        harness.stop()
        if (event) {
          replace(event.id, (item) =>
            item.kind === 'proposal' ? { ...item, state: 'rejected' } : item,
          )
        }
        publish({ applying: false, activity: '', status: 'cancelled' })
        return
      }
      if (event) {
        proposals.newTurn()
        replace(event.id, (item) =>
          item.kind === 'proposal' ? { ...item, state: 'rejected' } : item,
        )
      }
      harness.stop()
      if (wasRecovering)
        publish({ busy: false, activity: '', status: 'cancelled', retryable: false })
    },
    async confirm(id) {
      if (disposed || state.applying || (harness.snapshot.busy && proposals.pending()?.id !== id))
        return
      const epoch = sessionEpoch
      const event = pendingProposalEvent()
      if (event?.proposal.id === id) {
        replace(event.id, (item) =>
          item.kind === 'proposal' ? { ...item, state: 'applying', error: undefined } : item,
        )
      }
      publish({
        applying: true,
        error: undefined,
        errorMessage: undefined,
        retryable: false,
      })
      try {
        const decision = proposals.waitForDecision(id)
        await proposals.confirm(id)
        if (epoch !== sessionEpoch) return
        const outcome = await decision
        const writePending =
          outcome.status === 'applied_unverified' && outcome.safeCode === 'office_write_pending'
        if (event && writePending)
          replace(event.id, (item) =>
            item.kind === 'proposal'
              ? {
                  ...item,
                  state: 'uncertain',
                  error:
                    dependencies.presentationText?.('write_pending_quarantined') ??
                    'Write may still be running; further edits are frozen pending reconciliation or reload.',
                }
              : item,
          )
        else if (event)
          replace(event.id, (item) =>
            item.kind === 'proposal' ? { ...item, state: 'applied' } : item,
          )
        publish({
          activity: writePending
            ? (dependencies.presentationText?.('write_pending_quarantined') ??
              'Write may still be running; further edits are frozen pending reconciliation or reload.')
            : event?.proposal &&
                'impact' in event.proposal &&
                event.proposal.impact.host === 'local_team'
              ? '团队记录已保存'
              : 'Document updated',
          error: undefined,
          errorMessage: undefined,
          retryable: false,
        })
      } catch (error) {
        if (epoch !== sessionEpoch) return
        const safeError = safeConfirmationError(error)
        if (event)
          replace(event.id, (item) =>
            item.kind === 'proposal' ? { ...item, state: 'error', error: safeError.message } : item,
          )
        publish({
          error: safeError.code,
          errorMessage: safeError.message,
          retryable: safeError.retryable,
          status: 'error',
        })
      } finally {
        if (epoch === sessionEpoch) publish({ applying: false })
      }
    },
    reject() {
      if (disposed || state.applying) return
      const event = pendingProposalEvent()
      proposals.reject()
      if (event)
        replace(event.id, (item) =>
          item.kind === 'proposal' ? { ...item, state: 'rejected' } : item,
        )
      publish({
        activity: 'Proposal rejected',
        error: undefined,
        errorMessage: undefined,
        retryable: false,
      })
    },
    newTask() {
      if (disposed) return
      if (recoveryTimer) clearTimeout(recoveryTimer)
      recoveryTimer = undefined
      automaticRecoveryAttempt = 0
      sessionEpoch += 1
      finishCheckpoint()
      harness.reset()
      proposals.logout()
      lastInstruction = ''
      clearConversation()
      publish()
    },
    retry() {
      if (
        !lastInstruction ||
        !state.retryable ||
        harness.snapshot.busy ||
        state.applying ||
        disposed
      )
        return
      if (dependencies.runCheckpoint) void resumeRecovery(false)
      else {
        automaticRecoveryAttempt = 0
        startRun(RECOVERY_INSTRUCTION, undefined, undefined, false, '', true)
      }
    },
    async resumeInterrupted() {
      await resumeRecovery(true)
    },
    answerQuestionnaire(answers) {
      const resolve = clarificationResolve
      if (!resolve) return
      clarificationResolve = undefined
      questionnaireAnsweredPendingPlan = true
      publish({ questionnaire: undefined, activity: '继续规划演示文稿…' })
      resolve({
        output: `User questionnaire answers:\n${boundedText(answers)}\nContinue with plan_deck, research, slide creation, screenshots, and verify_slides now.`,
        mutated: false,
        summary: 'Collected questionnaire answers',
      })
    },
    skipQuestionnaire() {
      const resolve = clarificationResolve
      if (!resolve) return
      clarificationResolve = undefined
      questionnaireAnsweredPendingPlan = true
      publish({ questionnaire: undefined, activity: '继续规划演示文稿…' })
      resolve({
        output:
          'The user delegated these choices. Decide professionally and continue with plan_deck now.',
        mutated: false,
        summary: 'Questionnaire choices delegated',
      })
    },
    logout() {
      if (disposed) return
      sessionEpoch += 1
      finishCheckpoint()
      diagnose((diagnostics) => diagnostics.clear())
      harness.reset()
      proposals.logout()
      lastInstruction = ''
      clearConversation()
      publish()
    },
    authenticationLost() {
      if (disposed) return
      sessionEpoch += 1
      harness.reset()
      proposals.logout()
      lastInstruction = ''
      clearConversation()
      publish()
    },
    dispose() {
      if (disposed) return
      disposed = true
      if (recoveryTimer) clearTimeout(recoveryTimer)
      recoveryTimer = undefined
      dependencies.transport.setToolActivityHandler?.(undefined)
      dependencies.remoteTools?.setToolHandler?.(undefined)
      sessionEpoch += 1
      unsubscribeAcp()
      unsubscribeProposals()
      harness.dispose()
      proposals.logout()
      lastInstruction = ''
      clearConversation()
      cached = { ...state, proposal: undefined }
      listeners.clear()
    },
  }
}

export function bindAuthLoss(
  auth: { subscribeAuthLoss(listener: () => void): () => void },
  session: OfficeAgentSession,
  signedOut: () => void,
): () => void {
  return auth.subscribeAuthLoss(() => {
    session.authenticationLost()
    signedOut()
  })
}

export function useOfficeAgent(session: OfficeAgentSession): OfficeAgentSnapshot {
  return useSyncExternalStore(session.subscribe, session.snapshot, session.snapshot)
}
