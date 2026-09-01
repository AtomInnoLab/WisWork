import { createAgentHarness, type AgentHarness } from '@wiswork/agent-harness'
import type {
  AgentLoopOptions,
  AgentStreamCallbacks,
  AgentStreamRequest,
  AgentTransport,
  ToolExecution,
  ToolExecutionOutcome,
  ToolExecutionSuspension,
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

function createSlidesEnhancedHarness<TSnapshot>(
  options: AgentLoopOptions<TSnapshot>,
  api: PcHostCodexApi,
  documentId: string,
  generation: number,
): { harness: AgentHarness<TSnapshot>; close(): void } {
  let callbacks: AgentStreamCallbacks | null = null
  let closed = false
  const executions = new Map<string, ToolExecution>()
  const registration = api.register(
    createPcHostRegistration({ host: 'slides', documentId, generation, skill: options.skill }),
  )
  const unsubscribeEvents = api.onEvent((event) => {
    if (closed || !callbacks) return
    if (event.type === 'text') callbacks.onDelta(event.text)
    else if (event.type === 'done') callbacks.onDone()
    else if (event.type === 'error') callbacks.onError(event.code)
  })
  const unsubscribeTools = api.onToolCall((request) => {
    if (
      closed ||
      !callbacks ||
      request.documentId !== documentId ||
      request.generation !== generation
    )
      return
    callbacks.onToolCall(request.call)
    callbacks.onDone()
  })
  const transport: AgentTransport = {
    stream(request: AgentStreamRequest, next: AgentStreamCallbacks) {
      callbacks = next
      const toolMessage = request.messages.at(-1)
      if (toolMessage?.role === 'tool') {
        for (const result of toolMessage.results) {
          const execution = executions.get(result.id) ?? {
            output: result.output,
            summary: result.name,
            ...(result.isError ? { isError: true } : {}),
          }
          executions.delete(result.id)
          void api
            .toolResult({ documentId, generation, callId: result.id, execution })
            .catch(() => next.onError('enhanced_turn_failed'))
        }
      } else {
        const user = [...request.messages].reverse().find((message) => message.role === 'user')
        void registration
          .then(() => api.status())
          .then((status) => {
            if (status.activeAgentRuntime !== 'enhanced' || status.documentId !== documentId)
              throw new Error('enhanced_document_unavailable')
            return api.startTurn({ documentId, text: user?.role === 'user' ? user.text : '' })
          })
          .catch(() => next.onError('enhanced_turn_failed'))
      }
      return { cancel: () => void api.cancelTurn(documentId).catch(() => undefined) }
    },
  }
  const skill = {
    ...options.skill,
    async executeTool(call: Parameters<typeof options.skill.executeTool>[0], signal?: AbortSignal) {
      const outcome = await options.skill.executeTool(call, signal)
      if ('kind' in outcome && outcome.kind === 'tool-execution-suspension') {
        void outcome.result.then((execution) => executions.set(call.id, execution))
      } else executions.set(call.id, outcome)
      return outcome
    },
  }
  const harness = createAgentHarness({ ...options, transport, skill })
  return {
    harness,
    close() {
      if (closed) return
      closed = true
      harness.dispose()
      unsubscribeEvents()
      unsubscribeTools()
      void registration
        .catch(() => undefined)
        .then(() => api.cancelTurn(documentId).catch(() => undefined))
        .then(() => api.unregister(documentId, generation).catch(() => undefined))
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
  let closeEnhanced: (() => void) | null = null
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
      if (runtime && (inner as unknown as { mode?: string } | null)?.mode === 'enhanced') {
        const previous = generation
        closeEnhanced?.()
        closeEnhanced = null
        inner = null
        generation += 1
        const token = ++activation
        void runtime.api
          .unregister(documentId, previous)
          .catch(() => undefined)
          .then(() => createSelected(token))
          .catch(() => options.events?.onError?.('enhanced_document_unavailable'))
      } else inner?.reset()
    },
    restore(messages) {
      inner?.restore(messages)
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
        else
          void createSelected(++activation).catch((error) => {
            if (isPcHostCodexUnavailable(error) && !terminal) inner = createAgentHarness(options)
            else options.events?.onError?.('enhanced_document_unavailable')
          })
      }
    },
    deactivate() {
      if (closeEnhanced) closeEnhanced()
      else inner?.dispose()
      closeEnhanced = null
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
  finishHistoryBatch,
  isCurrent,
  hasQcPages,
  clearQcPages,
  runQc,
  setBusy,
  publishHistorySnapshot,
}: {
  cancelled: boolean
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
      if (cancelled) clearQcPages()
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
