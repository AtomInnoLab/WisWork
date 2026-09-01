import { AgentLoop, type AgentLoopOptions, type AgentMessage } from '@wiswork/agent-core'
import {
  createEnhancedRendererClient,
  EnhancedAgentRuntime,
  type AgentRuntimeSession,
  type PcHostCodexApi,
} from '@wiswork/agent-runtime'

export interface LatexRuntimeLoop {
  readonly busy: boolean
  readonly messages: readonly AgentMessage[]
  run(instruction: string): void
  cancel(): void
  restore(messages: readonly AgentMessage[]): void
  dispose(): void
}

export function createLatexRuntimeLoop(
  options: AgentLoopOptions,
  api?: PcHostCodexApi,
): LatexRuntimeLoop {
  let target: AgentLoop | AgentRuntimeSession | null = null
  let disposed = false
  let restored: readonly AgentMessage[] | undefined
  const documentId = `latex:${crypto.randomUUID()}`
  if (!api) target = new AgentLoop(options)
  else
    void api
      .status()
      .then((status) => {
        if (disposed) return
        if (status.activeAgentRuntime === 'standard') target = new AgentLoop(options)
        else
          target = new EnhancedAgentRuntime(
            createEnhancedRendererClient({
              ...api,
              subscribe: (_id, listener) => api.onEvent(listener),
            }),
          ).createSession({
            ...options,
            host: 'latex',
            document: { id: documentId, generation: 0 },
          })
        if (restored) target.restore(restored)
      })
      .catch(() => options.events?.onError?.('enhanced_document_unavailable'))
  return {
    get busy() {
      return target ? ('busy' in target ? target.busy : target.snapshot.busy) : true
    },
    get messages() {
      return target?.messages ?? restored ?? []
    },
    run(instruction) {
      if (target && !disposed) target.run(instruction)
    },
    cancel() {
      if (!target) return
      if ('cancel' in target) target.cancel()
      else target.stop()
    },
    restore(messages) {
      restored = messages
      target?.restore(messages)
    },
    dispose() {
      disposed = true
      if (!target) return
      if ('cancel' in target) target.cancel()
      else target.dispose()
    },
  }
}
