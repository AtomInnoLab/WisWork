import { createAgentHarness, type AgentHarness } from '@wiswork/agent-harness'
import type {
  AgentLoopOptions,
  AgentStreamCallbacks,
  AgentStreamRequest,
  AgentTransport,
  ToolExecution,
  ToolExecutionOutcome,
  ToolExecutionSuspension,
  AgentToolCall,
  ToolExecutedEvent,
} from '@wiswork/agent-core'
import {
  createPcHostRegistration,
  isPcHostCodexUnavailable,
  type PcEnhancedHost,
  type PcHostCodexApi,
} from '@wiswork/agent-runtime'
import { useEffect, useRef } from 'react'

interface LifecycleAgentController<TSnapshot> extends AgentHarness<TSnapshot> {
  activate(): void
  deactivate(): void
}

export interface AgentControllerRef<TSnapshot> {
  current: AgentHarness<TSnapshot> | null
}

const hostToolExecution = (execution: ToolExecution): ToolExecution => ({
  output: execution.output,
  summary: execution.summary,
  ...(execution.isError === undefined ? {} : { isError: execution.isError }),
  ...(execution.mutated === undefined ? {} : { mutated: execution.mutated }),
  ...(execution.stopToolBatch === undefined ? {} : { stopToolBatch: execution.stopToolBatch }),
  ...(execution.modelContent === undefined ? {} : { modelContent: execution.modelContent }),
})

/** Preserve only enumerated public codes, never raw IPC messages or document data. */
export function safeEnhancedError(error: unknown): string {
  const message = typeof error === 'string' ? error : error instanceof Error ? error.message : ''
  return (
    message.match(
      /\benhanced_(?:questionnaire_incomplete|turn_in_progress|turn_timeout|proposal_expired|auth_required|usage_limit|context_limit|request_rejected|service_unavailable|connection_failed|response_incompatible|document_unavailable|runtime_unavailable)\b/,
    )?.[0] ?? 'enhanced_turn_failed'
  )
}

function createSlidesEnhancedHarness<TSnapshot>(
  options: AgentLoopOptions<TSnapshot>,
  api: PcHostCodexApi,
  documentId: string,
  generation: number,
): { harness: AgentHarness<TSnapshot>; close(): Promise<void> } {
  const toolBatchQuietWindowMs = 20
  let callbacks: AgentStreamCallbacks | null = null
  let closed = false
  let turnSettled = true
  let turnEpoch = 0
  let terminalError: string | undefined
  let pendingUiTurnEnd = false
  let toolBatchTimer: ReturnType<typeof setTimeout> | null = null
  const executions = new Map<string, ToolExecution>()
  const toolControllers = new Map<string, AbortController>()
  const remoteCalls = new Map<string, AgentToolCall>()
  type Activity = {
    epoch: number
    remote: boolean
    done: boolean
    local?: ToolExecutedEvent<TSnapshot>
  }
  const activities = new Map<string, Activity>()
  const activityKey = (call: AgentToolCall) => call.invocationId ?? `${turnEpoch}:${call.id}`
  const startActivity = (call: AgentToolCall, remote: boolean) => {
    const key = activityKey(call)
    const existing = activities.get(key)
    if (existing) {
      if (remote) existing.remote = true
      return existing
    }
    // Match the document session's bounded tool-call lifetime.
    if (activities.size >= 1_024) return undefined
    const activity: Activity = { epoch: turnEpoch, remote, done: false }
    activities.set(key, activity)
    options.events?.onToolStart?.(call)
    return activity
  }
  const completeActivity = (event: ToolExecutedEvent<TSnapshot>, remote: boolean) => {
    if (!remote) event = { ...event, call: remoteCalls.get(event.call.id) ?? event.call }
    const activity = startActivity(event.call, remote)
    if (!activity || activity.done || activity.epoch !== turnEpoch) return
    if (!remote && activity.remote) {
      activity.local = event
      return
    }
    activity.done = true
    const local = activity.local
    activity.local = undefined
    options.events?.onToolExecuted?.(
      local && !!local.execution.isError === !!event.execution.isError ? local : event,
    )
    const call = remoteCalls.get(event.call.id)
    if (call) remoteCalls.set(event.call.id, { ...call, input: {} })
    if (
      pendingUiTurnEnd &&
      ![...activities.values()].some(
        (candidate) => candidate.epoch === turnEpoch && !candidate.done && candidate.local,
      )
    ) {
      pendingUiTurnEnd = false
      options.events?.onTurnEnd?.()
    }
  }
  const registration = api.register(
    createPcHostRegistration({ host: 'slides', documentId, generation, skill: options.skill }),
  )
  const clearToolBatchTimer = () => {
    if (toolBatchTimer === null) return
    clearTimeout(toolBatchTimer)
    toolBatchTimer = null
  }
  const settleTurn = (error?: string) => {
    if (turnSettled) return
    clearToolBatchTimer()
    turnSettled = true
    if (error) callbacks?.onError(error)
    else callbacks?.onDone()
  }
  const unsubscribeEvents = api.onEvent((event) => {
    if (closed || !callbacks) return
    if (event.type === 'text') callbacks.onDelta(event.text)
    else if (event.type === 'tool-start') startActivity(event.call, true)
    else if (event.type === 'tool-executed')
      completeActivity(event.event as ToolExecutedEvent<TSnapshot>, true)
    // startTurn resolves only after Shell releases document.busy. Runtime
    // terminal events arrive earlier and must not unlock the composer yet.
    else if (event.type === 'error') terminalError = safeEnhancedError(event.code)
  })
  const unsubscribeTools = api.onToolCall((request) => {
    if (
      closed ||
      !callbacks ||
      request.documentId !== documentId ||
      request.generation !== generation
    )
      return
    remoteCalls.set(request.call.id, request.call)
    callbacks.onToolCall(request.call)
    clearToolBatchTimer()
    const batchCallbacks = callbacks
    toolBatchTimer = setTimeout(() => {
      toolBatchTimer = null
      if (!closed && callbacks === batchCallbacks) batchCallbacks.onDone()
    }, toolBatchQuietWindowMs)
  })
  const unsubscribeToolCancels =
    api.onToolCancel?.((request) => {
      if (request.documentId !== documentId || request.generation !== generation) return
      toolControllers.get(request.callId)?.abort()
      toolControllers.delete(request.callId)
      executions.delete(request.callId)
    }) ?? (() => undefined)
  const transport: AgentTransport = {
    stream(request: AgentStreamRequest, next: AgentStreamCallbacks) {
      callbacks = next
      const toolMessage = request.messages.at(-1)
      if (toolMessage?.role === 'tool') {
        const submissions: Promise<void>[] = []
        for (const result of toolMessage.results) {
          const execution = executions.get(result.id) ?? {
            output: result.output,
            summary: result.name,
            ...(result.isError ? { isError: true } : {}),
          }
          executions.delete(result.id)
          submissions.push(api.toolResult({ documentId, generation, callId: result.id, execution }))
        }
        void Promise.all(submissions)
          .then(() => {
            // The runtime can finish while the renderer is still executing the
            // preceding tool batch. Replay that terminal edge to the new
            // tool-result stream so the local AgentLoop cannot stay busy.
            if (!closed && callbacks === next && turnSettled) next.onDone()
          })
          .catch((error) => next.onError(safeEnhancedError(error)))
      } else {
        const user = [...request.messages].reverse().find((message) => message.role === 'user')
        const epoch = ++turnEpoch
        pendingUiTurnEnd = false
        turnSettled = false
        terminalError = undefined
        void registration
          .then(() => api.status())
          .then((status) => {
            if (status.activeAgentRuntime !== 'enhanced' || status.documentId !== documentId)
              throw new Error('enhanced_document_unavailable')
            return api.startTurn({ documentId, text: user?.role === 'user' ? user.text : '' })
          })
          .then(() => {
            if (!closed && epoch === turnEpoch) settleTurn(terminalError)
          })
          .catch((error) => {
            if (!closed && epoch === turnEpoch)
              settleTurn(terminalError ?? safeEnhancedError(error))
          })
      }
      return { cancel: () => void api.cancelTurn(documentId).catch(() => undefined) }
    },
  }
  const skill = {
    ...options.skill,
    // This harness is a tool executor for the remote model, not a second
    // autonomous presentation run. Keep local completion policies so a remote
    // model cannot terminate after a successful questionnaire or before the
    // contract-bound production/verification loop is complete.
    presentation: undefined,
    async executeTool(call: Parameters<typeof options.skill.executeTool>[0], signal?: AbortSignal) {
      const controller = new AbortController()
      toolControllers.set(call.id, controller)
      const abort = () => controller.abort(signal?.reason)
      signal?.addEventListener('abort', abort, { once: true })
      if (signal?.aborted) abort()
      let outcome: Awaited<ReturnType<typeof options.skill.executeTool>>
      try {
        outcome = await options.skill.executeTool(call, controller.signal)
      } finally {
        signal?.removeEventListener('abort', abort)
        toolControllers.delete(call.id)
      }
      controller.signal.throwIfAborted()
      if ('kind' in outcome && outcome.kind === 'tool-execution-suspension') {
        void outcome.result.then((execution) =>
          executions.set(call.id, hostToolExecution(execution)),
        )
      } else executions.set(call.id, hostToolExecution(outcome))
      return outcome
    },
  }
  const events: AgentLoopOptions<TSnapshot>['events'] = {
    ...options.events,
    onToolStart(call) {
      startActivity(remoteCalls.get(call.id) ?? call, false)
    },
    onToolExecuted(event) {
      completeActivity(event, false)
    },
    onTurnEnd() {
      // The local executor advances before IPC results return. Keep its current
      // chat bubble open until every executed tool receives its authoritative receipt.
      if (
        [...activities.values()].some(
          (activity) => activity.epoch === turnEpoch && !activity.done && activity.local,
        )
      )
        pendingUiTurnEnd = true
      else options.events?.onTurnEnd?.()
    },
    onError(error) {
      if (!closed) void api.cancelTurn(documentId).catch(() => undefined)
      options.events?.onError?.(error)
    },
  }
  const harness = createAgentHarness({ ...options, transport, skill, events })
  let closePromise: Promise<void> | null = null
  return {
    harness,
    close() {
      if (closePromise) return closePromise
      closed = true
      clearToolBatchTimer()
      harness.dispose()
      unsubscribeEvents()
      unsubscribeTools()
      unsubscribeToolCancels()
      for (const controller of toolControllers.values()) controller.abort()
      toolControllers.clear()
      activities.clear()
      remoteCalls.clear()
      closePromise = registration
        .catch(() => undefined)
        .then(() => api.cancelTurn(documentId).catch(() => undefined))
        .then(() => api.unregister(documentId, generation).catch(() => undefined))
      return closePromise
    },
  }
}

export function classifySlidesQcFailure(
  error: unknown,
  signal?: AbortSignal,
): 'cancelled' | 'failed' {
  return signal?.aborted || (error instanceof Error && error.name === 'AbortError')
    ? 'cancelled'
    : 'failed'
}

export const createAgentController = <TSnapshot>(
  options: AgentLoopOptions<TSnapshot>,
  runtime?: { readonly host: PcEnhancedHost; readonly api: PcHostCodexApi },
): LifecycleAgentController<TSnapshot> => {
  let inner: AgentHarness<TSnapshot> | null = runtime ? null : createAgentHarness(options)
  let terminal = false
  let activation = 0
  let generation = 0
  let closeEnhanced: (() => Promise<void>) | null = null
  let lifecycle = Promise.resolve()
  let enhancedActive = false
  const documentId = `${runtime?.host ?? 'standard'}:${crypto.randomUUID()}`
  const createSelected = async (token: number) => {
    if (!runtime) return
    const status = await runtime.api.status()
    if (terminal || token !== activation) return
    if (status.activeAgentRuntime === 'standard') {
      inner = createAgentHarness(options)
      return
    }
    const selected = createSlidesEnhancedHarness(options, runtime.api, documentId, generation)
    inner = selected.harness
    closeEnhanced = selected.close
    enhancedActive = true
  }
  const controller: LifecycleAgentController<TSnapshot> = {
    get snapshot() {
      return inner?.snapshot ?? { status: 'idle', busy: false, generation: 0 }
    },
    get messages() {
      return inner?.messages ?? []
    },
    subscribe(listener) {
      return inner?.subscribe(listener) ?? (() => undefined)
    },
    run(instruction, images) {
      return inner?.run(instruction, images) ?? false
    },
    stop() {
      inner?.stop()
    },
    reset() {
      if (runtime && enhancedActive) {
        const close = closeEnhanced
        closeEnhanced = null
        enhancedActive = false
        inner = null
        generation += 1
        const token = ++activation
        lifecycle = lifecycle.then(() => close?.()).then(() => undefined)
        void lifecycle
          .then(() => createSelected(token))
          .catch(() => options.events?.onError?.('enhanced_document_unavailable'))
      } else inner?.reset()
    },
    restore(messages) {
      inner?.restore(messages)
    },
    appendAssistantContext(text) {
      return inner?.appendAssistantContext(text) ?? false
    },
    suspendToolExecution(result) {
      if (!inner?.suspendToolExecution) throw new Error('enhanced_suspension_owned_by_shell')
      return inner.suspendToolExecution(result)
    },
    ownsToolExecutionSuspension(value: ToolExecutionOutcome): value is ToolExecutionSuspension {
      return inner?.ownsToolExecutionSuspension?.(value) === true
    },
    activate() {
      if (!terminal && !inner) {
        if (!runtime) inner = createAgentHarness(options)
        else {
          const token = ++activation
          void lifecycle
            .then(() => createSelected(token))
            .catch((error) => {
              if (isPcHostCodexUnavailable(error) && !terminal) inner = createAgentHarness(options)
              else options.events?.onError?.('enhanced_document_unavailable')
            })
        }
      }
    },
    deactivate() {
      const close = closeEnhanced
      if (close) {
        lifecycle = lifecycle.then(() => close()).then(() => undefined)
      } else inner?.dispose()
      closeEnhanced = null
      enhancedActive = false
      inner = null
      activation++
    },
    dispose() {
      terminal = true
      controller.deactivate()
    },
  }
  return controller
}

export function disposeAgentController<TSnapshot>(ref: AgentControllerRef<TSnapshot>): void {
  ref.current?.dispose()
  ref.current = null
}

export function useAgentControllerCleanup<TSnapshot>(ref: AgentControllerRef<TSnapshot>): void {
  const ownerRef = useRef<LifecycleAgentController<TSnapshot> | null>(null)
  useEffect(() => {
    const owner = ownerRef.current ?? (ref.current as LifecycleAgentController<TSnapshot> | null)
    ownerRef.current = owner
    owner?.activate()
    ref.current = owner
    return () => {
      owner?.deactivate()
      if (ref.current === owner) ref.current = null
    }
  }, [ref])
}

export function recordSlidesRunAttachments<TAttachment>(
  attachments: readonly TAttachment[],
  persistUser: (attachments: readonly TAttachment[]) => void,
): void {
  persistUser(attachments)
}

export async function beginSlidesHostRun({
  beginHistoryBatch,
  isCurrent,
  markHistoryActive,
  finishHistoryBatch,
  run,
}: {
  beginHistoryBatch: () => Promise<boolean>
  isCurrent: () => boolean
  markHistoryActive: () => void
  finishHistoryBatch: () => Promise<unknown>
  run: () => boolean
}): Promise<boolean> {
  const opened = await beginHistoryBatch()
  // Once the host has opened a batch it must be tracked before any stale/cancel
  // branch tries to close it; otherwise finishHistoryBatch sees "inactive" and
  // leaks the host session in a batched state.
  if (opened) markHistoryActive()
  if (!isCurrent()) {
    if (opened) await finishHistoryBatch()
    return false
  }
  // AgentHarness may synchronously project a launch failure through onError
  // before run() returns. Mark the opened batch first so that callback can
  // close it; a rejected launch closes it through the same host wrapper.
  const started = run()
  if (!started) {
    if (opened) await finishHistoryBatch()
    return false
  }
  return true
}

export async function completeSlidesHostRun({
  cancelled,
  qualityReviewOwner = 'host',
  finishHistoryBatch,
  isCurrent,
  hasQcPages,
  clearQcPages,
  runQc,
  setBusy,
  publishHistorySnapshot,
}: {
  cancelled: boolean
  qualityReviewOwner?: 'host' | 'agent'
  finishHistoryBatch: () => Promise<unknown>
  isCurrent?: () => boolean
  hasQcPages: () => boolean
  clearQcPages: () => void
  runQc: () => void
  setBusy: (busy: boolean) => void
  publishHistorySnapshot?: (snapshot: unknown) => void
}): Promise<void> {
  try {
    const snapshot = await finishHistoryBatch()
    if (isCurrent && !isCurrent()) return
    publishHistorySnapshot?.(snapshot)
  } finally {
    if (!isCurrent || isCurrent()) {
      setBusy(false)
      if (cancelled || qualityReviewOwner === 'agent') clearQcPages()
      else if (hasQcPages()) runQc()
    }
  }
}

export function stopSlidesHostRun({
  dismissClarification,
  abortQc,
  stop,
}: {
  dismissClarification: () => void
  abortQc: () => void
  stop: () => void
}): void {
  dismissClarification()
  abortQc()
  stop()
}
