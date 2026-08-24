import { createAgentHarness, type AgentHarness } from '@wiswork/agent-harness'
import type { AgentLoopOptions, AgentMessage } from '@wiswork/agent-core'
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

export const selectSheetsExecution = (agentConfigured: boolean): 'agent' | 'planner' =>
  agentConfigured ? 'agent' : 'planner'

export function restoreSheetsSession<TSnapshot>(
  controller: AgentHarness<TSnapshot> | null,
  loadedSessionId: string | number | undefined,
  currentSessionId: () => string | number | undefined,
  messages: readonly AgentMessage[],
): boolean {
  if (!controller || loadedSessionId !== currentSessionId()) return false
  controller.restore(messages)
  return true
}

export async function settleSheetsApplyPromises(
  applies: Promise<boolean>[],
  autosave: () => Promise<void>,
): Promise<boolean> {
  const pending = applies.splice(0)
  if (pending.length === 0) return false
  const results = await Promise.all(pending)
  if (!results.some(Boolean)) return false
  await autosave()
  return true
}
