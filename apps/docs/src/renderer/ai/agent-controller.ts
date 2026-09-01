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

export const createAgentController = <TSnapshot>(
  options: AgentLoopOptions<TSnapshot>,
  runtime?: { readonly host: PcEnhancedHost; readonly api: PcHostCodexApi },
): LifecycleAgentController<TSnapshot> => {
  let inner: AgentHarness<TSnapshot> | null = runtime ? null : createAgentHarness(options)
  let terminal = false
  let activation = 0
  let generation = 0
  const documentId = `${runtime?.host ?? 'standard'}:${crypto.randomUUID()}`
  const createEnhanced = async (token: number) => {
    if (!runtime) return
    const status = await runtime.api.status()
    if (terminal || token !== activation) return
    if (status.activeAgentRuntime === 'standard') {
      inner = createAgentHarness(options)
      return
    }
    const client = createEnhancedRendererClient({
      ...runtime.api,
      subscribe: (_documentId, listener) => runtime.api.onEvent(listener),
    })
    const selected = new EnhancedAgentRuntime(client).createSession({
      ...options,
      host: runtime.host,
      document: { id: documentId, generation },
    })
    inner = selected as unknown as AgentHarness<TSnapshot>
  }
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
    reset: () => {
      if (runtime && (inner as unknown as { mode?: string } | null)?.mode === 'enhanced') {
        const previous = generation
        inner?.dispose()
        inner = null
        generation += 1
        const token = ++activation
        void runtime.api
          .unregister(documentId, previous)
          .catch(() => undefined)
          .then(() => createEnhanced(token))
          .catch(() => options.events?.onError?.('enhanced_document_unavailable'))
      } else inner?.reset()
    },
    restore: (messages) => inner?.restore(messages),
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
          void createEnhanced(++activation).catch(() =>
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
