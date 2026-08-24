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
    subscribe: (listener) => inner?.subscribe(listener) ?? (() => undefined),
    run: (instruction, images) => inner?.run(instruction, images) ?? false,
    stop: () => inner?.stop(),
    reset: () => inner?.reset(),
    restore: (messages) => inner?.restore(messages),
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

export function createAgentLaunchOwner() {
  let generation = 0
  return {
    invalidate: () => {
      generation++
    },
    async launch<T>(
      collect: (isCurrent: () => boolean) => Promise<T>,
      run: (value: T) => boolean,
    ): Promise<boolean> {
      const launchGeneration = ++generation
      const isCurrent = () => launchGeneration === generation
      let value: T
      try {
        value = await collect(isCurrent)
      } catch (error) {
        if (!isCurrent()) return false
        throw error
      }
      return launchGeneration === generation ? run(value) : false
    },
  }
}

export const shouldResetAgentSession = (
  previousPath: string | null | undefined,
  nextPath: string | null | undefined,
): boolean => previousPath != null && previousPath !== nextPath

export function createAgentRunStartingGuard() {
  let active = 0
  let sequence = 0
  return {
    begin(): number | null {
      if (active !== 0) return null
      active = ++sequence
      return active
    },
    end(token: number): void {
      if (active === token) active = 0
    },
    clear(): void {
      active = 0
    },
    isActive: () => active !== 0,
  }
}
