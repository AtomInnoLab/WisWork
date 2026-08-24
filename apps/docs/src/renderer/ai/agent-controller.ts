import { createAgentHarness, type AgentHarness } from '@wiswork/agent-harness'
import type { AgentLoopOptions } from '@wiswork/agent-core'
import { useEffect } from 'react'

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
  useEffect(() => () => disposeAgentController(ref), [ref])
}
