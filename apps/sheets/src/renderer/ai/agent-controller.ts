import { createAgentHarness, type AgentHarness } from '@wiswork/agent-harness'
import type {
  AgentLoopOptions,
  AgentMessage,
  ToolExecutionOutcome,
  ToolExecutionSuspension,
} from '@wiswork/agent-core'
import {
  createEnhancedRendererClient,
  EnhancedAgentRuntime,
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
          void createSelected(++activation).catch((error) => {
            if (isPcHostCodexUnavailable(error) && !terminal) inner = createAgentHarness(options)
            else options.events?.onError?.('enhanced_document_unavailable')
          })
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

export const selectSheetsExecution = (agentConfigured: boolean): 'agent' | 'planner' =>
  agentConfigured ? 'agent' : 'planner'

export interface AsyncGenerationGate {
  begin(): number
  invalidate(): void
  commit(token: number, effect: () => void): boolean
}

export function createAsyncGenerationGate(): AsyncGenerationGate {
  let generation = 0
  return {
    begin: () => ++generation,
    invalidate: () => {
      generation += 1
    },
    commit: (token, effect) => {
      if (token !== generation) return false
      effect()
      return true
    },
  }
}

export const createSheetsChatLoadCoordinator = createAsyncGenerationGate

export function classifySheetsDocumentTransition(
  previous: string | number | undefined,
  current: string | number | undefined,
): 'open' | 'rebind' {
  return current !== undefined && previous === current ? 'rebind' : 'open'
}

export function getSheetsDocumentIdentity(file: {
  path?: string | undefined
  sessionId: string | number
  documentInstanceId: string | number
}): string | number {
  return file.documentInstanceId
}

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
