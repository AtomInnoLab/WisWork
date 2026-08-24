import { createAgentHarness, type AgentHarness } from '@wiswork/agent-harness'
import type { AgentLoopOptions } from '@wiswork/agent-core'
import { useEffect, useRef } from 'react'

export interface AgentControllerRef<TSnapshot> {
  current: AgentHarness<TSnapshot> | null
}

export const createAgentController = <TSnapshot>(
  options: AgentLoopOptions<TSnapshot>,
): AgentHarness<TSnapshot> => createAgentHarness(options)

export function disposeAgentController<TSnapshot>(ref: AgentControllerRef<TSnapshot>): void {
  ref.current?.dispose()
  ref.current = null
}

export function useAgentControllerCleanup<TSnapshot>(ref: AgentControllerRef<TSnapshot>): void {
  const setupToken = useRef(0)
  useEffect(() => {
    const token = ++setupToken.current
    return () => {
      const owned = ref.current
      queueMicrotask(() => {
        if (setupToken.current === token && ref.current === owned) disposeAgentController(ref)
      })
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
  markHistoryActive,
  run,
}: {
  beginHistoryBatch: () => Promise<boolean>
  markHistoryActive: () => void
  run: () => void
}): Promise<void> {
  if (await beginHistoryBatch()) markHistoryActive()
  run()
}

export async function completeSlidesHostRun({
  cancelled,
  finishHistoryBatch,
  hasQcPages,
  clearQcPages,
  runQc,
  setBusy,
}: {
  cancelled: boolean
  finishHistoryBatch: () => Promise<unknown>
  hasQcPages: () => boolean
  clearQcPages: () => void
  runQc: () => void
  setBusy: (busy: boolean) => void
}): Promise<void> {
  try {
    await finishHistoryBatch()
  } finally {
    setBusy(false)
    if (cancelled) clearQcPages()
    else if (hasQcPages()) runQc()
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
