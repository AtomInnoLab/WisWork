import {
  suspendToolExecution,
  type AgentSkill,
  type AgentTransport,
  type ToolExecution,
  type ToolExecutionOutcome,
} from '@wiswork/agent-core'
import { createAgentHarness } from '@wiswork/agent-harness'
import type { OfficePowerPointVisualReviewer } from '../skills/powerpoint/powerpoint-verification.js'
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
  type OfficeClarificationQuestion,
  type OfficePresentationTimeline,
  type ProposalPresentationEvent,
} from './presentation-state.js'
import type { PresentationVerificationStringKey } from '@wiswork/i18n'
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
  proposal?: OfficeProposal | StructuredProposal
  questionnaire?: readonly OfficeClarificationQuestion[]
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

function toolActivity(name: string, state: 'running' | 'complete' | 'error'): string {
  const labels: Readonly<Record<string, string>> = {
    web_search: '网页搜索',
    web_fetch: '读取网页',
    image_search: '图片搜索',
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
  'image_fetch_unavailable',
  'image_limit',
  'image_mime_unsupported',
  'invalid_image',
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

export function presentationClarificationText(
  question: string,
  translate?: (key: PresentationVerificationStringKey) => string,
): string {
  if (!question || question === 'presentation_scope_required')
    return translate?.('clarify') ?? 'More information needed'
  return boundedText(question)
}

export function createOfficeAgentSession(dependencies: {
  transport: AgentTransport
  skill: AgentSkill
  proposals: ProposalController | StructuredProposalController
  diagnostics?: Pick<OfficeDiagnostics, 'startTrace' | 'setTool' | 'record' | 'clear'>
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
  let state: Omit<OfficeAgentSnapshot, 'proposal'> = {
    assistantText: '',
    activity: '',
    busy: false,
    applying: false,
    status: 'idle',
    retryable: false,
    timeline: emptyPresentationTimeline(),
  }
  let cached: OfficeAgentSnapshot = { ...state, proposal: visibleProposal() }

  const publish = (next: Partial<typeof state> = {}) => {
    if (disposed) return
    state = { ...state, ...next }
    cached = { ...state, proposal: visibleProposal() }
    listeners.forEach((listener) => listener())
  }

  let nextEventId = 0
  let sessionEpoch = 0
  let activeAssistantId: string | undefined
  let cumulativeAssistantText = ''
  let assistantSegmentPrefix = ''
  let lastInstruction = ''
  let clarificationResolve: ((value: ToolExecution) => void) | undefined
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
      const outcome = await dependencies.skill.executeTool(call, signal)
      if ('kind' in outcome && outcome.kind === 'tool-execution-suspension') return outcome
      const proposal = proposals.pending()
      if (!proposal) return outcome
      const final = finalProposalExecution(proposal.id, outcome, call.name)
      if (isAutomaticProposal(proposal)) {
        void proposals.confirm(proposal.id).catch(() => {
          // The proposal controller records a stable failed decision. finalProposalExecution
          // resumes the same tool call with that safe result instead of creating an unhandled task.
        })
      }
      return suspendToolExecution(final)
    },
  }
  dependencies.remoteTools?.setToolHandler?.(async (call) => {
    const definition = sessionSkill.tools.find((tool) => tool.name === call.toolName)
    if (!definition) return { output: 'unknown_tool', isError: true }
    assistantSegmentPrefix = cumulativeAssistantText
    closeAssistantSegment()
    diagnose((diagnostics) => diagnostics.setTool(call.toolName))
    const presentationId = eventId()
    const startedAt = Date.now()
    const runningSummary = toolActivity(call.toolName, 'running')
    append({
      id: presentationId,
      kind: 'tool',
      callId: call.callId,
      name: boundedText(call.toolName),
      summary: runningSummary,
      state: 'running',
    })
    publish({ activity: runningSummary })
    const invalidateRemoteProposal = () => proposals.newTurn()
    call.signal.addEventListener('abort', invalidateRemoteProposal, { once: true })
    try {
      const outcome = await sessionSkill.executeTool(
        {
          id: call.callId,
          invocationId: `${call.turnId}:${call.callId}`,
          name: call.toolName,
          input: call.input,
        },
        call.signal,
      )
      const settled =
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
      const finishedSummary = toolActivity(call.toolName, settled.isError ? 'error' : 'complete')
      replace(presentationId, (event) =>
        event.kind === 'tool'
          ? {
              ...event,
              summary: finishedSummary,
              state: settled.isError ? 'error' : 'complete',
              durationMs: Date.now() - startedAt,
              output: boundedText(settled.output),
              ...(settled.display ? { display: settled.display } : {}),
            }
          : event,
      )
      publish({ activity: finishedSummary })
      if (settled.isError) {
        const errorCode = diagnosticToolError(settled.output)
        diagnose((diagnostics) =>
          diagnostics.record({
            phase: 'tool',
            errorCode,
            durationMs: Math.max(0, Date.now() - startedAt),
          }),
        )
      }
      call.signal.removeEventListener('abort', invalidateRemoteProposal)
      return { output: settled.output, ...(settled.isError ? { isError: true } : {}) }
    } catch {
      const failedSummary = toolActivity(call.toolName, 'error')
      replace(presentationId, (event) =>
        event.kind === 'tool'
          ? { ...event, summary: failedSummary, state: 'error', durationMs: Date.now() - startedAt }
          : event,
      )
      publish({ activity: failedSummary })
      diagnose((diagnostics) =>
        diagnostics.record({
          phase: 'tool',
          errorCode: call.signal.aborted ? 'cancelled' : 'tool_execution_failed',
          durationMs: Math.max(0, Date.now() - startedAt),
        }),
      )
      if (call.signal.aborted) proposals.newTurn()
      return { output: 'tool_execution_failed', isError: true }
    } finally {
      call.signal.removeEventListener('abort', invalidateRemoteProposal)
    }
  })
  const clearConversation = () => {
    activeAssistantId = undefined
    cumulativeAssistantText = ''
    assistantSegmentPrefix = ''
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
      timeline: emptyPresentationTimeline(),
      questionnaire: undefined,
    }
  }

  const harness = createAgentHarness({
    transport: dependencies.transport,
    skill: sessionSkill,
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
        if (!assistantText.startsWith(assistantSegmentPrefix)) assistantSegmentPrefix = ''
        cumulativeAssistantText = assistantText
        const segment = boundedText(assistantText.slice(assistantSegmentPrefix.length))
        if (!activeAssistantId) {
          activeAssistantId = eventId()
          append({
            id: activeAssistantId,
            kind: 'assistant',
            text: segment,
            streaming: true,
          })
        } else {
          replace(activeAssistantId, (event) => ({
            ...event,
            text: segment,
            streaming: true,
          }))
        }
        publish({ assistantText: segment })
      },
      onToolStart: (call) => {
        toolStartedAt.set(call.id, Date.now())
        diagnose((diagnostics) => diagnostics.setTool(call.name))
        cumulativeAssistantText = ''
        assistantSegmentPrefix = ''
        closeAssistantSegment()
        const summary = toolActivity(call.name, 'running')
        append({
          id: eventId(),
          kind: 'tool',
          callId: call.id,
          name: boundedText(call.name),
          summary,
          state: 'running',
        })
        publish({ activity: summary })
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
        const tool = [...state.timeline]
          .reverse()
          .find((item) => item.kind === 'tool' && item.callId === event.call.id)
        if (tool) {
          const summary = toolActivity(
            event.call.name,
            event.execution.isError ? 'error' : 'complete',
          )
          replace(tool.id, (item) => {
            if (item.kind !== 'tool') return item
            return {
              ...item,
              summary,
              state: event.execution.isError ? 'error' : 'complete',
              output: boundedText(event.execution.output),
              ...(event.execution.display ? { display: event.execution.display } : {}),
            }
          })
        }
        appendPendingProposal()
        publish({
          activity: toolActivity(event.call.name, event.execution.isError ? 'error' : 'complete'),
        })
      },
      onTurnEnd: () => {
        activeAssistantId = undefined
        publish({ activity: 'Thinking…' })
      },
      onDone: (result) => {
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

  const unsubscribeProposals = proposals.subscribe(() => {
    const pending = visibleProposal()
    if (pending) {
      appendPendingProposal()
      publish({ activity: 'Waiting for your approval' })
    } else {
      publish()
    }
  })

  const startRun = (instruction: string) => {
    const value = instruction.trim()
    if (!value || harness.snapshot.busy || state.applying || disposed) return
    diagnose((diagnostics) => diagnostics.startTrace())
    staleTools.clear()
    runStartedAt = Date.now()
    proposals.newTurn()
    lastInstruction = value
    activeAssistantId = undefined
    cumulativeAssistantText = ''
    assistantSegmentPrefix = ''
    append({ id: eventId(), kind: 'user', text: boundedText(value) })
    publish({
      assistantText: '',
      activity: 'Thinking…',
      busy: true,
      status: 'working',
      error: undefined,
      errorMessage: undefined,
      retryable: false,
    })
    harness.run(value)
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
      sessionEpoch += 1
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
    answerQuestionnaire(answers) {
      const resolve = clarificationResolve
      if (!resolve) return
      clarificationResolve = undefined
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
      dependencies.remoteTools?.setToolHandler?.(undefined)
      sessionEpoch += 1
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
