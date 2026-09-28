import {
  suspendToolExecution,
  type AgentSkill,
  type AgentTransport,
  type ToolExecution,
  type ToolExecutionOutcome,
} from '@wiswork/agent-core'
import { acpToolActivity, createAgentHarness } from '@wiswork/agent-harness'
import { useSyncExternalStore } from 'react'
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
import type { OfficeDiagnostics } from '../diagnostics/office-diagnostics.js'

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

const DIAGNOSTIC_TOOL_ERRORS = new Set([
  'cancelled',
  'invalid_tool_input',
  'office_api_unsupported',
  'office_read_failed',
  'office_overwrite_required',
  'office_recovery_failed',
  'office_concurrent_change',
  'office_state_uncertain',
  'office_verify_failed',
  'office_write_failed',
  'proposal_missing',
  'proposal_stale',
])

function diagnosticToolError(output: string): string {
  const safe = (value: string) =>
    DIAGNOSTIC_TOOL_ERRORS.has(value) || /^office_recovery_failed:word_[a-z_]+$/.test(value)
  if (safe(output)) return output
  try {
    const parsed = JSON.parse(output) as { error?: unknown }
    return typeof parsed.error === 'string' && safe(parsed.error)
      ? parsed.error
      : 'agent_run_failed'
  } catch {
    return 'agent_run_failed'
  }
}

export function createOfficeAgentSession(dependencies: {
  transport: AgentTransport
  skill: AgentSkill
  proposals: ProposalController | StructuredProposalController
  diagnostics?: Pick<OfficeDiagnostics, 'startTrace' | 'setTool' | 'record' | 'clear'>
  runCheckpoint?: {
    interrupted: boolean
    scrubFailed?: boolean
    recovery?: {
      instruction: string
      phase: 'running' | 'tool_pending' | 'tool_completed'
      toolName?: string
    }
    validateDocument?(): Promise<boolean>
    begin(runId: string, instruction: string): Promise<void>
    tool?(runId: string, phase: 'tool_pending' | 'tool_completed', toolName: string): Promise<void>
    finish(runId: string): Promise<void>
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
  let state: Omit<OfficeAgentSnapshot, 'proposal'> = {
    assistantText: '',
    activity: '',
    busy: false,
    applying: false,
    status: 'idle',
    retryable: false,
    recoveryAvailable:
      dependencies.runCheckpoint?.recovery?.phase === 'running' &&
      Boolean(dependencies.runCheckpoint.recovery.instruction),
    timeline: dependencies.runCheckpoint?.interrupted
      ? appendPresentationEvent(emptyPresentationTimeline(), {
          id: 'event-1',
          kind: 'system',
          text: dependencies.runCheckpoint.scrubFailed
            ? '上次运行已中断。旧版检查点中的请求原文仍保留在本 PPTX：清理保存失败。请先保存可写副本并重新打开，期间不能继续该运行。'
            : `上次前台 Agent 运行在面板关闭时中断。${dependencies.runCheckpoint.recovery?.toolName ? `最近工具：${dependencies.runCheckpoint.recovery.toolName}（${dependencies.runCheckpoint.recovery.phase}）。` : ''}请先核对项目、页面和写入记录；未自动重放写入。运行阶段保存在演示文稿设置中，请求仅保存在本机浏览器。`,
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
  const finishCheckpoint = () => {
    const id = activeRunId
    activeRunId = undefined
    if (id) void dependencies.runCheckpoint?.finish(id).catch(() => undefined)
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
      if (decision.error === 'proposal_stale') staleTools.add(toolName)
      return {
        output: JSON.stringify({
          proposalId,
          status: 'failed',
          error: decision.error,
          instruction:
            decision.error === 'proposal_stale'
              ? 'Do not retry this write in the current turn.'
              : undefined,
        }),
        isError: true,
        mutated: false,
        summary: 'Approved change failed',
        stopToolBatch: decision.error === 'proposal_stale',
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
      const checkpointCompleted = async (): Promise<boolean> => {
        if (runId && currentRun() && dependencies.runCheckpoint?.tool)
          try {
            await dependencies.runCheckpoint.tool(runId, 'tool_completed', call.name)
          } catch {
            return false
          }
        return true
      }
      const checkpointed = async (result: ToolExecution): Promise<ToolExecution> =>
        (await checkpointCompleted())
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
          await dependencies.runCheckpoint.tool(runId, 'tool_pending', call.name)
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
    events: {
      onText: (assistantText) => {
        // Presentation is driven by ACP agent_message_chunk updates below.
        publish({ assistantText: boundedText(assistantText) })
      },
      onToolStart: (call) => {
        toolStartedAt.set(call.id, Date.now())
        diagnose((diagnostics) => diagnostics.setTool(call.name))
      },
      onToolExecuted: (event) => {
        if (event.execution.isError) {
          const errorCode = diagnosticToolError(event.execution.output)
          diagnose((diagnostics) =>
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
            }),
          )
        }
        toolStartedAt.delete(event.call.id)
        appendPendingProposal()
      },
      onTurnEnd: () => {
        activeAssistantId = undefined
        publish({ activity: 'Thinking…' })
      },
      onDone: (result) => {
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
        if (error === 'presentation_run_checkpoint_unavailable') activeRunId = undefined
        else finishCheckpoint()
        const safeError = safeRunError(error)
        diagnose((diagnostics) =>
          diagnostics.record({
            phase: 'transport',
            errorCode: safeError.code,
            durationMs: Math.max(0, Date.now() - runStartedAt),
          }),
        )
        activeAssistantId = undefined
        append({
          id: eventId(),
          kind: 'error',
          text: safeError.message,
          code: safeError.code,
        })
        publish({
          busy: false,
          activity: '',
          status: 'error',
          error: safeError.code,
          errorMessage: safeError.message,
          retryable: safeError.retryable,
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
      const summary = acpToolActivity(name, 'running')
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
      const summary = acpToolActivity(tool.name, terminal)
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

  const startRun = (instruction: string) => {
    const value = instruction.trim()
    if (!value || harness.snapshot.busy || pendingStart || state.applying || disposed) return
    diagnose((diagnostics) => diagnostics.startTrace())
    staleTools.clear()
    runStartedAt = Date.now()
    proposals.newTurn()
    lastInstruction = value
    activeAssistantId = undefined
    append({ id: eventId(), kind: 'user', text: boundedText(value) })
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
    const runId = crypto.randomUUID()
    void dependencies.runCheckpoint
      .begin(runId, value.length <= 1000 ? value : '')
      .then(() => {
        pendingStart = false
        if (disposed || epoch !== sessionEpoch) {
          void dependencies.runCheckpoint?.finish(runId).catch(() => undefined)
          return
        }
        activeRunId = runId
        harness.run(value)
      })
      .catch(() => {
        pendingStart = false
        if (disposed || epoch !== sessionEpoch) return
        append({
          id: eventId(),
          kind: 'error',
          text: '无法保存运行检查点，请确认文档可保存后重试。',
          code: 'presentation_run_checkpoint_unavailable',
        })
        publish({
          busy: false,
          activity: '',
          status: 'error',
          error: 'presentation_run_checkpoint_unavailable',
          errorMessage: '无法保存运行检查点，请确认文档可保存后重试。',
          retryable: true,
        })
      })
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
      startRun(instruction)
    },
    async resumeInterrupted() {
      const checkpoint = dependencies.runCheckpoint
      if (
        !checkpoint?.recovery?.instruction ||
        checkpoint.recovery.phase !== 'running' ||
        state.busy ||
        state.applying ||
        disposed ||
        !checkpoint.validateDocument
      )
        return
      try {
        if (!(await checkpoint.validateDocument())) return
      } catch {
        return
      }
      startRun(checkpoint.recovery.instruction)
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
