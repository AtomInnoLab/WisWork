import {
  suspendToolExecution,
  parseAgentResumeMessages,
  type AgentMessage,
  type AgentSkill,
  type AgentToolCall,
  type AgentTransport,
  type ToolExecution,
  type ToolExecutionOutcome,
} from '@wiswork/agent-core'
import { acpToolActivity, createAgentHarness } from '@wiswork/agent-harness'
import { presentationRecoveryReceiptFeedback } from './presentation-recovery-feedback.js'
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
  type OfficePresentationTimeline,
  type ProposalPresentationEvent,
  type ToolPresentationEvent,
} from './presentation-state.js'
import {
  isDiagnosticToolError,
  type OfficeDiagnostics,
  type PresentationDiagnosticContext,
} from '../diagnostics/office-diagnostics.js'

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
  timeline: OfficePresentationTimeline
}

export interface OfficeAgentSession {
  snapshot(): OfficeAgentSnapshot
  subscribe(listener: () => void): () => void
  send(instruction: string): void
  stop(): void
  confirm(id: string): Promise<void>
  reject(): void
  newTask(): void
  retry(): void
  resumeInterrupted?(): Promise<void>
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
    message: 'The change may be partially applied. Inspect the document before trying again.',
    retryable: false,
  },
})

const runErrors: Readonly<Record<string, SafeSessionError>> = Object.freeze({
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
    message: 'The Agent service is temporarily unavailable. Try again.',
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

const safeRunError = (error: string): SafeSessionError =>
  runErrors[error === 'transport_timeout' ? 'request_timeout' : error] ?? {
    code: 'agent_run_failed',
    message: 'The Agent could not complete this request. Try again.',
    retryable: true,
  }

function diagnosticToolError(output: string): string {
  if (isDiagnosticToolError(output)) return output
  try {
    const parsed = JSON.parse(output) as { error?: unknown }
    return typeof parsed.error === 'string' && isDiagnosticToolError(parsed.error)
      ? parsed.error
      : 'agent_run_failed'
  } catch {
    return 'agent_run_failed'
  }
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
  transport: AgentTransport
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
}): OfficeAgentSession {
  const { proposals } = dependencies
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
        (dependencies.runCheckpoint?.recovery?.phase === 'tool_completed' &&
          dependencies.runCheckpoint?.recovery?.restartSafe === true)) &&
      Boolean(dependencies.runCheckpoint.recovery.instruction),
    timeline: dependencies.runCheckpoint?.interrupted
      ? appendPresentationEvent(emptyPresentationTimeline(), {
          id: 'event-1',
          kind: 'system',
          text: dependencies.runCheckpoint.scrubFailed
            ? '上次运行已中断。旧版检查点中的请求原文仍保留在本 PPTX：清理保存失败。请先保存可写副本并重新打开，期间不能继续该运行。'
            : `上次前台 Agent 运行在面板关闭时中断。${receiptFeedback}${dependencies.runCheckpoint.recovery?.messages ? '已保留完整只读结果，可在核对文档后继续；旧结果代表历史读取，当前状态仍需重新核对。' : dependencies.runCheckpoint.recovery?.phase === 'running' ? '尚未调用工具，可在核对文档后主动重新运行原请求。' : dependencies.runCheckpoint.recovery?.restartSafe ? '此前仅运行了可重读工具，可在核对文档后主动重新运行原请求。' : '请先核对项目、页面和写入记录；未自动重放写入。'}运行阶段保存在演示文稿设置中，请求与可恢复的读取结果仅保存在本机浏览器。`,
        })
      : emptyPresentationTimeline(),
  }
  let cached: OfficeAgentSnapshot = { ...state, proposal: proposals.pending() }

  const publish = (next: Partial<typeof state> = {}) => {
    if (disposed) return
    state = { ...state, ...next }
    cached = { ...state, proposal: proposals.pending() }
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
  let runStartedAt = 0
  const staleTools = new Set<string>()
  const toolStartedAt = new Map<string, number>()
  const eventId = () => `event-${++nextEventId}`
  const append = (event: Parameters<typeof appendPresentationEvent>[1]) => {
    state = { ...state, timeline: appendPresentationEvent(state.timeline, event) }
  }
  const replace = (id: string, update: Parameters<typeof replacePresentationEvent>[2]) => {
    state = { ...state, timeline: replacePresentationEvent(state.timeline, id, update) }
  }
  const pendingProposalEvent = () =>
    [...state.timeline]
      .reverse()
      .find(
        (event): event is ProposalPresentationEvent =>
          event.kind === 'proposal' && event.state === 'pending',
      )
  const appendPendingProposal = () => {
    const proposal = proposals.pending()
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
          instruction: stale ? 'Do not retry this write in the current turn.' : undefined,
        }),
        isError: true,
        mutated: false,
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
    async executeTool(call, signal): Promise<ToolExecutionOutcome> {
      toolsStarted = true
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
      const proposal = proposals.pending()
      if (proposal)
        return suspendToolExecution(
          finalProposalExecution(proposal.id, outcome, call.name).then(checkpointed),
        )
      return checkpointed(outcome)
    },
  }
  const clearConversation = () => {
    checkpointBeginFailed = false
    activeAssistantId = undefined
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
    }
  }

  const harness = createAgentHarness({
    transport: dependencies.transport,
    skill: sessionSkill,
    systemSuffix: () =>
      resumingReadConversation
        ? '\nThis run resumes a saved read-only conversation. Restored tool results are historical observations. Revalidate relevant live document/project state before making changes or claiming its current state; do not treat cached results as proof that the document is unchanged.'
        : '',
    events: {
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
        toolStartedAt.clear()
        if (activeAssistantId) {
          replace(activeAssistantId, (event) => ({ ...event, streaming: false }))
          activeAssistantId = undefined
        }
        publish({
          busy: false,
          activity: '',
          status: result.cancelled ? 'cancelled' : 'done',
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
    const pending = proposals.pending()
    if (pending) {
      appendPendingProposal()
      publish({ activity: 'Waiting for your approval' })
    } else {
      publish()
    }
  })

  let resumingReadConversation = false
  const startRun = (
    instruction: string,
    messages?: readonly AgentMessage[],
    resumedRunId?: string,
  ) => {
    const value = instruction.trim()
    if (!value || harness.snapshot.busy || pendingStart || state.applying || disposed) return
    diagnose((diagnostics) => diagnostics.startTrace())
    staleTools.clear()
    checkpointBeginFailed = false
    toolsStarted = false
    runStartedAt = Date.now()
    resumingReadConversation = Boolean(messages)
    proposals.newTurn()
    lastInstruction = value
    activeAssistantId = undefined
    append({
      id: eventId(),
      kind: messages ? 'assistant' : 'user',
      text: messages ? '从已保存的读取结果继续。' : boundedText(value),
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
      harness.run(value)
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
        await proposals.confirm(id)
        if (epoch !== sessionEpoch) return
        if (event)
          replace(event.id, (item) =>
            item.kind === 'proposal' ? { ...item, state: 'applied' } : item,
          )
        publish({
          activity: 'Document updated',
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
      const instruction = lastInstruction
      if (dependencies.runCheckpoint) void resumeRecovery(false)
      else startRun(instruction)
    },
    async resumeInterrupted() {
      await resumeRecovery(true)
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
