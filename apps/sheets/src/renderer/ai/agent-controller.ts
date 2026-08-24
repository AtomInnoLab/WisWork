import { createAgentHarness, type AgentHarness } from '@wiswork/agent-harness'
import type { AgentLoopOptions, AgentMessage } from '@wiswork/agent-core'
import { useEffect, useRef } from 'react'

interface LifecycleAgentController<TSnapshot> extends AgentHarness<TSnapshot> {
  activate(): void
  deactivate(): void
}

export interface AgentControllerRef<TSnapshot> {
  current: AgentHarness<TSnapshot> | null
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

export const selectSheetsExecution = (agentConfigured: boolean): 'agent' | 'planner' =>
  agentConfigured ? 'agent' : 'planner'

export function bindSheetsSession<TSnapshot>(
  controller: AgentHarness<TSnapshot>,
  binding: { current: string | number | undefined },
  sessionId: string | number | undefined,
): void {
  if (binding.current !== sessionId) controller.reset()
  binding.current = sessionId
}

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
