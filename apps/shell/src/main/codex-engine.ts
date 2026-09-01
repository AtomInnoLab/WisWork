import {
  CodexProcessManager,
  createDocumentCarrierIssuer,
  startDynamicMcpGateway,
  startResponsesBridge,
  type DocumentToolSession,
  type MessagesRequest,
} from '@wiswork/codex-bridge'
import type {
  CodexRuntimeBootstrap,
  CodexRuntimeEngine,
  CodexRuntimeEngineEvent,
} from './codex-runtime'
import { CodexTurnResolver } from './codex-turn-resolver'

const DEVELOPER_POLICY =
  'Use only mcp__wiswork__wiswork_call. Never request shell, filesystem, Git, browser, network, or direct document writes. Treat capability values as secrets and never repeat them.'
const TURN_TERMINAL_TIMEOUT_MS = 120_000

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
      let bridge: Awaited<ReturnType<typeof startResponsesBridge>> | undefined
      let gateway: Awaited<ReturnType<typeof startDynamicMcpGateway>> | undefined
      let manager: CodexProcessManager | undefined
      let client
      try {
        bridge = await startResponsesBridge({
          fetchWithAuth: options.fetchWithAuth,
          prepareTurn: resolver.prepare,
        })
        gateway = await startDynamicMcpGateway(options.diagnostics)
        manager = new CodexProcessManager({
          executablePath,
          bridge: { baseUrl: bridge.baseUrl, secret: bridge.secret },
          mcp: { url: gateway.url, secret: gateway.secret },
          developerInstructions: DEVELOPER_POLICY,
          diagnostics: ({ code }) => options.diagnostics?.(code),
        })
        client = await manager.start()
        await client.initialize()
      } catch (error) {
        resolver.clear()
        await manager?.stop().catch(() => undefined)
        await Promise.allSettled([bridge?.close(), gateway?.close()])
        throw error
      }
      type ActiveTurn = {
        readonly capability: string
        threadId?: string
        turnId?: string
        disarm?: () => void
        cancelled: boolean
        readonly settle: (status: 'completed' | 'cancelled' | 'failed', error?: Error) => void
      }
      const documents = new Map<
        string,
        {
          session: DocumentToolSession
          unregister: () => void
          onEvent?: (event: CodexRuntimeEngineEvent) => void
          active?: ActiveTurn
        }
      >()
      const emit = (
        listener: ((event: CodexRuntimeEngineEvent) => void) | undefined,
        event: CodexRuntimeEngineEvent,
      ): void => {
        try {
          listener?.(event)
        } catch {}
      }
      const unsubscribeNotifications = client.onNotification((notification) => {
        options.diagnostics?.(`codex_${notification.method.replaceAll('/', '_')}`)
        const params = notification.params as {
          threadId?: string
          turnId?: string
          delta?: string
          turn?: { id?: string; status?: unknown }
        }
        const document = [...documents.values()].find(
          (candidate) => candidate.active?.threadId === params.threadId,
        )
        const active = document?.active
        if (!document || !active) return
        if (notification.method === 'item/agentMessage/delta') {
          if (params.turnId === active.turnId && typeof params.delta === 'string') {
            emit(document.onEvent, { type: 'text', text: params.delta })
          }
          return
        }
        if (notification.method === 'turn/completed') {
          const turn = params.turn
          if (turn?.id !== active.turnId) return
          if (typeof turn?.status === 'string' && /^[a-z_]{1,32}$/.test(turn.status)) {
            options.diagnostics?.(`codex_turn_status_${turn.status}`)
          }
          active.settle(
            turn?.status === 'interrupted'
              ? 'cancelled'
              : turn?.status === 'failed'
                ? 'failed'
                : 'completed',
            turn?.status === 'failed' ? new Error('enhanced_turn_failed') : undefined,
          )
        }
      })
      void manager.crashed.then(onCrash)
      let closed = false
      return {
        registerDocument(input) {
          if (closed || documents.has(input.documentId))
            throw new Error('document_session_unavailable')
          const unregister = gateway.register({
            ...input,
            onToolEvent: (event) => emit(input.onEvent, event),
          })
          const entry: {
            session: DocumentToolSession
            unregister: () => void
            onEvent?: (event: CodexRuntimeEngineEvent) => void
            active?: ActiveTurn
          } = { session: input.session, unregister, onEvent: input.onEvent }
          documents.set(input.documentId, entry)
          return () => {
            if (documents.get(input.documentId) === entry) {
              documents.delete(input.documentId)
              if (entry.active) {
                entry.active.settle('failed', new Error('document_session_unavailable'))
              }
              unregister()
            }
          }
        },
        async startTurn(input) {
          const document = documents.get(input.documentId)
          if (!document) throw new Error('document_session_unavailable')
          if (document.active) throw new Error('turn_already_active')
          const grant = gateway.beginTurn({
            documentId: input.documentId,
            generation: input.generation,
            threadId: 'reserved',
          })
          let resolve!: () => void
          let reject!: (error: Error) => void
          let timer: ReturnType<typeof setTimeout> | undefined
          let settled = false
          const terminal = new Promise<void>((onResolve, onReject) => {
            resolve = onResolve
            reject = onReject
          })
          const active: ActiveTurn = {
            capability: grant.capability,
            cancelled: false,
            settle(status, error) {
              if (settled) return
              settled = true
              if (timer) clearTimeout(timer)
              gateway.revokeTurn(grant.capability)
              active.disarm?.()
              if (document.active === active) document.active = undefined
              emit(document.onEvent, { type: 'terminal', status })
              if (error) reject(error)
              else resolve()
            },
          }
          document.active = active
          timer = setTimeout(
            () => active.settle('failed', new Error('enhanced_turn_timeout')),
            TURN_TERMINAL_TIMEOUT_MS,
          )
          timer.unref()
          try {
            active.threadId = (
              await client.startThread({
                developerInstructions: `${DEVELOPER_POLICY}\nFor this turn only, pass capability ${grant.capability} as the capability argument to mcp__wiswork__wiswork_call. Never repeat it in prose.`,
              })
            ).thread.id
            if (active.cancelled) return await terminal
            gateway.bindTurn(grant.capability, active.threadId)
            active.disarm = resolver.arm(
              active.threadId,
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
            if (active.cancelled) return await terminal
            active.turnId = (await client.startTurn(active.threadId, input.text)).turn.id
            if (active.cancelled) return await terminal
            await terminal
          } catch (error) {
            if (active.cancelled) return await terminal
            active.settle(
              'failed',
              error instanceof Error ? error : new Error('enhanced_turn_failed'),
            )
            await terminal
          }
        },
        async cancelTurn(documentId) {
          const document = documents.get(documentId)
          const active = document?.active
          if (!active) return
          active.cancelled = true
          gateway.revokeTurn(active.capability)
          if (active.threadId && active.turnId) {
            await client.interruptTurn(active.threadId, active.turnId).catch(() => undefined)
          }
          active.settle('cancelled')
        },
        async closeDocument(documentId) {
          const document = documents.get(documentId)
          if (!document) return
          documents.delete(documentId)
          if (document.active) {
            document.active.settle('failed', new Error('document_session_unavailable'))
          }
          document.unregister()
          document.session.close()
        },
        async close() {
          if (closed) return
          closed = true
          resolver.clear()
          unsubscribeNotifications()
          for (const document of documents.values()) {
            if (document.active) {
              document.active.settle('failed', new Error('enhanced_runtime_closed'))
            }
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
