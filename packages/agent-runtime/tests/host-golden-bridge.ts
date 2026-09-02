import type { AgentToolCall } from '@wiswork/agent-core'
import type { EnhancedRendererBridge } from '../src/renderer'
import { writeFileSync } from 'node:fs'

export function writeEnhancedGoldenReport(report: {
  readonly host: string
  readonly verification: string
  readonly rollback: string
}): void {
  const path = process.env.WISWORK_ENHANCED_GOLDEN_REPORT
  if (!path) return
  const encoded = JSON.stringify(report)
  if (Buffer.byteLength(encoded, 'utf8') > 512) throw new Error('enhanced_golden_report_too_large')
  writeFileSync(path, encoded, { encoding: 'utf8', flag: 'wx', mode: 0o600 })
}

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
