import {
  CodexProcessManager,
  createDocumentCarrierIssuer,
  startDynamicMcpGateway,
  startResponsesBridge,
  type DocumentToolSession,
  type MessagesRequest,
} from '@wiswork/codex-bridge'
import type { CodexRuntimeBootstrap, CodexRuntimeEngine } from './codex-runtime'
import { CodexTurnResolver } from './codex-turn-resolver'

const DEVELOPER_POLICY =
  'Use only mcp__wiswork__wiswork_call. Never request shell, filesystem, Git, browser, network, or direct document writes. Treat capability values as secrets and never repeat them.'

export interface ProductionCodexBootstrapOptions {
  readonly fetchWithAuth: (request: MessagesRequest, signal: AbortSignal) => Promise<Response>
  readonly diagnostics?: (code: string) => void
}

export function createProductionCodexBootstrap(
  options: ProductionCodexBootstrapOptions,
): CodexRuntimeBootstrap {
  return {
    async start({ executablePath, onCrash }): Promise<CodexRuntimeEngine> {
      const resolver = new CodexTurnResolver(options.diagnostics)
      const bridge = await startResponsesBridge({
        fetchWithAuth: options.fetchWithAuth,
        prepareTurn: resolver.prepare,
      })
      const gateway = await startDynamicMcpGateway(options.diagnostics)
      const manager = new CodexProcessManager({
        executablePath,
        bridge: { baseUrl: bridge.baseUrl, secret: bridge.secret },
        mcp: { url: gateway.url, secret: gateway.secret },
        developerInstructions: DEVELOPER_POLICY,
        diagnostics: ({ code }) => options.diagnostics?.(code),
      })
      const client = await manager.start()
      await client.initialize()
      const unsubscribeNotifications = client.onNotification((notification) => {
        options.diagnostics?.(`codex_${notification.method.replaceAll('/', '_')}`)
        if (notification.method === 'turn/completed') {
          const turn = (notification.params as { turn?: { status?: unknown } }).turn
          if (typeof turn?.status === 'string' && /^[a-z_]{1,32}$/.test(turn.status)) {
            options.diagnostics?.(`codex_turn_status_${turn.status}`)
          }
        }
      })
      void manager.crashed.then(onCrash)
      const documents = new Map<
        string,
        { session: DocumentToolSession; unregister: () => void; threadId?: string; turnId?: string }
      >()
      let closed = false
      return {
        registerDocument(input) {
          if (closed || documents.has(input.documentId))
            throw new Error('document_session_unavailable')
          const unregister = gateway.register(input)
          const entry = { session: input.session, unregister }
          documents.set(input.documentId, entry)
          return () => {
            if (documents.get(input.documentId) === entry) {
              documents.delete(input.documentId)
              unregister()
            }
          }
        },
        async startTurn(input) {
          const document = documents.get(input.documentId)
          if (!document) throw new Error('document_session_unavailable')
          const grant = gateway.beginTurn({
            documentId: input.documentId,
            generation: input.generation,
            threadId: 'reserved',
          })
          document.threadId = (
            await client.startThread({
              developerInstructions: `${DEVELOPER_POLICY}\nFor this turn only, pass capability ${grant.capability} as the capability argument to mcp__wiswork__wiswork_call. Never repeat it in prose.`,
            })
          ).thread.id
          gateway.bindTurn(grant.capability, document.threadId)
          const disarm = resolver.arm(
            document.threadId,
            ({ sessionId }) =>
              createDocumentCarrierIssuer(
                {
                  host: input.host,
                  documentId: input.documentId,
                  sessionId,
                  generation: input.generation,
                },
                (capability) => capability === grant.capability,
              ),
            grant.capability,
          )
          try {
            document.turnId = (await client.startTurn(document.threadId, input.text)).turn.id
          } catch (error) {
            disarm()
            throw error
          }
        },
        async cancelTurn(documentId) {
          const document = documents.get(documentId)
          if (document?.threadId && document.turnId)
            await client.interruptTurn(document.threadId, document.turnId)
        },
        async closeDocument(documentId) {
          const document = documents.get(documentId)
          if (!document) return
          documents.delete(documentId)
          document.unregister()
          document.session.close()
        },
        async close() {
          if (closed) return
          closed = true
          resolver.clear()
          unsubscribeNotifications()
          for (const document of documents.values()) {
            document.unregister()
            document.session.close()
          }
          documents.clear()
          await manager.stop()
          await Promise.allSettled([bridge.close(), gateway.close()])
        },
      }
    },
  }
}
