import { createAgentHarness, type AgentHarness } from '@wiswork/agent-harness'
import type {
  AgentLoopOptions,
  ToolExecutionOutcome,
  ToolExecutionSuspension,
} from '@wiswork/agent-core'
import {
  createEnhancedRendererClient,
  EnhancedAgentRuntime,
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
  const documentId = `${runtime?.host ?? 'standard'}:${crypto.randomUUID()}`
  const createSelected = async (token: number) => {
    if (!runtime) return
    const status = await runtime.api.status()
    if (terminal || token !== activation) return
    if (status.activeAgentRuntime === 'standard') {
      inner = createAgentHarness(options)
      return
    }
    const client = createEnhancedRendererClient({
      ...runtime.api,
      subscribe: (_id, listener) => runtime.api.onEvent(listener),
    })
    inner = new EnhancedAgentRuntime(client).createSession({
      ...options,
      host: runtime.host,
      document: { id: documentId, generation },
    }) as unknown as AgentHarness<TSnapshot>
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
        inner?.dispose()
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
          void createSelected(++activation).catch(() =>
            options.events?.onError?.('enhanced_document_unavailable'),
          )
      }
    },
    deactivate() {
      inner?.dispose()
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
