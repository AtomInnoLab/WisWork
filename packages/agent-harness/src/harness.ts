import {
  AgentLoop,
  type AgentImage,
  type AgentLoopEvents,
  type AgentLoopOptions,
  type AgentMessage,
} from '@wiswork/agent-core'

export type AgentHarnessStatus = 'idle' | 'running' | 'done' | 'cancelled' | 'error'

export interface AgentHarnessSnapshot {
  status: AgentHarnessStatus
  busy: boolean
  generation: number
  error?: string
}

export interface AgentHarness<_TSnapshot> {
  readonly snapshot: AgentHarnessSnapshot
  readonly messages: readonly AgentMessage[]
  subscribe(listener: () => void): () => void
  run(instruction: string, images?: AgentImage[]): boolean
  stop(): void
  reset(): void
  restore(messages: readonly AgentMessage[]): void
  dispose(): void
  /** Opaque loop-owned authority used by reviewed Enhanced mutation controllers. */
  suspendToolExecution(
    result: Promise<import('@wiswork/agent-core').ToolExecution>,
  ): import('@wiswork/agent-core').ToolExecutionSuspension
}

export function createAgentHarness<TSnapshot>(
  options: AgentLoopOptions<TSnapshot>,
): AgentHarness<TSnapshot> {
  const hostEvents = options.events
  const loopOptions: AgentLoopOptions<TSnapshot> = { ...options }
  const loop = new AgentLoop(loopOptions)
  const listeners = new Set<() => void>()
  let currentSnapshot: AgentHarnessSnapshot = {
    status: 'idle',
    busy: false,
    generation: 0,
  }
  let disposed = false
  let launchPending = false

  const publish = (snapshot: AgentHarnessSnapshot): void => {
    currentSnapshot = snapshot
    for (const listener of [...listeners]) {
      try {
        listener()
      } catch {
        // One presentation subscriber must not prevent lifecycle settlement.
      }
    }
  }

  const invoke = (callback: (() => void) | undefined): void => {
    try {
      callback?.()
    } catch {
      // Host presentation failures must not strand the lifecycle as busy.
    }
  }

  const isCurrent = (generation: number): boolean =>
    !disposed && generation === currentSnapshot.generation

  const eventsFor = (generation: number): AgentLoopEvents<TSnapshot> => ({
    onText: (text) => {
      if (!isCurrent(generation)) return
      invoke(() => hostEvents?.onText?.(text))
    },
    onToolStart: (call) => {
      if (!isCurrent(generation)) return
      invoke(() => hostEvents?.onToolStart?.(call))
    },
    onToolExecuted: (event) => {
      if (!isCurrent(generation)) return
      invoke(() => hostEvents?.onToolExecuted?.(event))
    },
    onTurnEnd: () => {
      if (!isCurrent(generation)) return
      invoke(() => hostEvents?.onTurnEnd?.())
    },
    onDone: (result) => {
      if (!isCurrent(generation)) return
      invoke(() => hostEvents?.onDone?.(result))
      if (!isCurrent(generation)) return
      publish({
        status: result.cancelled ? 'cancelled' : 'done',
        busy: false,
        generation,
      })
    },
    onError: (error) => {
      if (!isCurrent(generation)) return
      invoke(() => hostEvents?.onError?.(error))
      if (!isCurrent(generation)) return
      publish({ status: 'error', busy: false, generation, error })
    },
    onPresentationClarify: (event) => {
      if (!isCurrent(generation)) return
      invoke(() => hostEvents?.onPresentationClarify?.(event))
    },
    onPresentationPlan: (event) => {
      if (!isCurrent(generation)) return
      invoke(() => hostEvents?.onPresentationPlan?.(event))
    },
    onPresentationCorrection: (event) => {
      if (!isCurrent(generation)) return
      invoke(() => hostEvents?.onPresentationCorrection?.(event))
    },
    onPresentationReceipt: (event) => {
      if (!isCurrent(generation)) return undefined
      try {
        return hostEvents?.onPresentationReceipt?.(event)
      } catch {
        return undefined
      }
    },
    onAbandonedPresentationCompletion: (event) => {
      if (!isCurrent(generation)) return
      invoke(() => hostEvents?.onAbandonedPresentationCompletion?.(event))
    },
  })

  return {
    get snapshot() {
      return currentSnapshot
    },
    get messages() {
      return loop.messages
    },
    subscribe(listener) {
      if (disposed) return () => undefined
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    run(instruction, images) {
      if (disposed || launchPending || loop.busy || !instruction) return false
      const generation = currentSnapshot.generation + 1
      loopOptions.events = eventsFor(generation)
      launchPending = true
      publish({ status: 'running', busy: true, generation })
      if (!isCurrent(generation) || !launchPending) return false
      launchPending = false
      try {
        loop.run(instruction, images)
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error)
        invoke(() => hostEvents?.onError?.(message))
        if (isCurrent(generation)) {
          publish({ status: 'error', busy: false, generation, error: message })
        }
      }
      return true
    },
    stop() {
      if (disposed) return
      if (launchPending) {
        launchPending = false
        loopOptions.events?.onDone?.({ text: '', cancelled: true, turnLimit: false })
        return
      }
      if (!loop.busy) return
      loop.cancel()
    },
    reset() {
      if (disposed) return
      launchPending = false
      const generation = currentSnapshot.generation + 1
      loop.reset()
      publish({ status: 'idle', busy: false, generation })
    },
    restore(messages) {
      if (disposed) return
      const before = loop.messages
      loop.restore(messages)
      if (loop.messages !== before) publish({ ...currentSnapshot })
    },
    dispose() {
      if (disposed) return
      disposed = true
      launchPending = false
      loop.reset()
      currentSnapshot = {
        status: 'idle',
        busy: false,
        generation: currentSnapshot.generation + 1,
      }
      listeners.clear()
    },
    suspendToolExecution(result) {
      return loop.suspendToolExecution(result)
    },
  }
}
