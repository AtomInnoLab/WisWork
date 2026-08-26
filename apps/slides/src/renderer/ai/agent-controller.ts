import { createAgentHarness, type AgentHarness } from '@wiswork/agent-harness'
import type { AgentLoopOptions } from '@wiswork/agent-core'
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
): LifecycleAgentController<TSnapshot> => {
  let inner: AgentHarness<TSnapshot> | null = createAgentHarness(options)
  let terminal = false
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
      inner?.reset()
    },
    restore(messages) {
      inner?.restore(messages)
    },
    activate() {
      if (!terminal && !inner) inner = createAgentHarness(options)
    },
    deactivate() {
      inner?.dispose()
      inner = null
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
