import type { EnhancedHost } from './contracts'
import type {
  EnhancedRuntimeClient,
  EnhancedRuntimeClientSession,
  EnhancedSessionEvent,
} from './enhanced'

export interface EnhancedRendererBridge {
  status(): Promise<{
    readonly activeAgentRuntime: 'standard' | 'enhanced'
    readonly documentId: string | null
  }>
  startTurn(input: { readonly documentId: string; readonly text: string }): Promise<void>
  cancelTurn(documentId: string): Promise<void>
  subscribe(documentId: string, listener: (event: EnhancedSessionEvent) => void): () => void
}

/** Renderer-only adapter. It holds no component path, process token, or document authority. */
export function createEnhancedRendererClient(
  bridge: EnhancedRendererBridge,
): EnhancedRuntimeClient {
  let closed = false
  const sessions = new Set<EnhancedRuntimeClientSession>()
  return Object.freeze({
    open(input: { host: EnhancedHost; documentId: string; generation: number }) {
      if (closed) throw new Error('enhanced_runtime_closed')
      let ended = false
      const listeners = new Set<(event: EnhancedSessionEvent) => void>()
      const unsubscribe = bridge.subscribe(input.documentId, (event) => {
        if (ended) return
        for (const listener of [...listeners]) listener(event)
      })
      const session: EnhancedRuntimeClientSession = Object.freeze({
        async start(turn: {
          readonly text: string
          readonly images?: readonly import('@wiswork/agent-core').AgentImage[]
        }) {
          if (ended) throw new Error('enhanced_session_closed')
          const status = await bridge.status()
          if (status.activeAgentRuntime !== 'enhanced' || status.documentId !== input.documentId) {
            throw new Error('enhanced_document_unavailable')
          }
          if (turn.images?.length) throw new Error('enhanced_images_unavailable')
          await bridge.startTurn({ documentId: input.documentId, text: turn.text })
        },
        cancel: () => (ended ? Promise.resolve() : bridge.cancelTurn(input.documentId)),
        subscribe(listener: (event: EnhancedSessionEvent) => void) {
          if (ended) return () => undefined
          listeners.add(listener)
          return () => listeners.delete(listener)
        },
        async close() {
          if (ended) return
          ended = true
          unsubscribe()
          listeners.clear()
          sessions.delete(session)
          await bridge.cancelTurn(input.documentId).catch(() => undefined)
        },
      })
      sessions.add(session)
      return session
    },
    async close() {
      if (closed) return
      closed = true
      await Promise.all([...sessions].map((session) => session.close()))
    },
  })
}
