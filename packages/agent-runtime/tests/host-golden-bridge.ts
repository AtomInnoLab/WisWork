import type { AgentToolCall } from '@wiswork/agent-core'
import type { EnhancedRendererBridge } from '../src/renderer'

export function createHostGoldenBridge(input: {
  readonly documentId: string
  readonly generation: number
  readonly call: AgentToolCall
}): EnhancedRendererBridge {
  let toolListener: Parameters<EnhancedRendererBridge['onToolCall']>[0] | undefined
  let eventListener: Parameters<EnhancedRendererBridge['subscribe']>[1] | undefined
  return {
    status: async () => ({ activeAgentRuntime: 'enhanced', documentId: input.documentId }),
    register: async () => undefined,
    unregister: async () => undefined,
    subscribe: (_documentId, listener) => {
      eventListener = listener
      return () => undefined
    },
    onToolCall: (listener) => {
      toolListener = listener
      return () => undefined
    },
    startTurn: async () => {
      toolListener?.({
        documentId: input.documentId,
        generation: input.generation,
        call: input.call,
      })
    },
    toolResult: async (result) => {
      eventListener?.({
        type: 'tool-executed',
        event: {
          call: input.call,
          execution: result.execution,
          ...(result.snapshotBefore ? { snapshotBefore: result.snapshotBefore } : {}),
        },
      })
    },
    cancelTurn: async () => undefined,
  }
}
