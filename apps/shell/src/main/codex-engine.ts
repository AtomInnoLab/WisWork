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
  "Use mcp__wiswork__wiswork_read only for read tools and mcp__wiswork__wiswork_propose only for mutation proposals. A proposal never changes the document; only the host UI can confirm it later. Every carrier call MUST use exactly {capability,callId,toolName,input}: capability is the latest private wiswork_turn_capability, callId is a new short unique string, toolName is one exact semantic tool name from the document catalog, and input is that semantic tool's argument object. Never flatten semantic arguments into the carrier object. Never request shell, filesystem, Git, browser, network, or direct document writes. Never repeat a capability in prose."
const TURN_IDLE_TIMEOUT_MS = 60_000
const INTERRUPT_TIMEOUT_MS = 2_000

export function safeTurnFailure(params: unknown): string {
  const error =
    typeof params === 'object' && params !== null
      ? (params as { error?: { codexErrorInfo?: unknown } }).error
      : undefined
  const info = error?.codexErrorInfo
  if (info === 'unauthorized') return 'enhanced_auth_required'
  if (info === 'usageLimitExceeded' || info === 'sessionBudgetExceeded')
    return 'enhanced_usage_limit'
  if (info === 'contextWindowExceeded') return 'enhanced_context_limit'
  if (info === 'badRequest' || info === 'cyberPolicy') return 'enhanced_request_rejected'
  if (info === 'serverOverloaded' || info === 'internalServerError')
    return 'enhanced_service_unavailable'
  if (typeof info === 'object' && info !== null) {
    const detail = Object.values(info as Record<string, unknown>)[0]
    const status =
      typeof detail === 'object' && detail !== null
        ? (detail as { httpStatusCode?: unknown }).httpStatusCode
        : undefined
    if (status === 401 || status === 403) return 'enhanced_auth_required'
    if (status === 408 || status === 429 || (typeof status === 'number' && status >= 500)) {
      return 'enhanced_service_unavailable'
    }
    return 'enhanced_connection_failed'
  }
  return 'enhanced_turn_failed'
}

export function createTurnIdleDeadline(onExpire: () => void, timeoutMs = TURN_IDLE_TIMEOUT_MS) {
  let timer: ReturnType<typeof setTimeout> | undefined
  const touch = () => {
    if (timer) clearTimeout(timer)
    timer = setTimeout(onExpire, timeoutMs)
    timer.unref()
  }
  const disarm = () => {
    if (timer) clearTimeout(timer)
    timer = undefined
  }
  touch()
  return Object.freeze({ touch, disarm })
}

/** Detached, bounded cleanup: callers settle user-visible cancellation before invoking this. */
export function startBestEffortCodexInterrupt(interrupt: () => Promise<unknown>): void {
  let timer: ReturnType<typeof setTimeout> | undefined
  const timeout = new Promise<void>((resolve) => {
    timer = setTimeout(resolve, INTERRUPT_TIMEOUT_MS)
    timer.unref()
  })
  void Promise.race([Promise.resolve().then(interrupt), timeout])
    .catch(() => undefined)
    .finally(() => {
      if (timer) clearTimeout(timer)
    })
}

export interface ProductionCodexBootstrapOptions {
  readonly fetchWithAuth: (request: MessagesRequest, signal: AbortSignal) => Promise<Response>
  readonly diagnostics?: (code: string) => void
  readonly onProtocolRecording?: (
    recording: import('@wiswork/codex-bridge').ProtocolRecording,
    outcome: import('@wiswork/codex-bridge').ProtocolRecordingOutcome,
  ) => void
}

const DOCUMENT_CATALOG_MAX_BYTES = 64 * 1024

export function buildDocumentToolInstructions(session: DocumentToolSession): string {
  const tools = session.listTools(session.credentials).map((tool) => ({
    name: tool.name,
    description: tool.description,
    inputSchema: tool.inputSchema,
    carrier:
      tool.annotations.readOnlyHint && !tool.annotations.destructiveHint
        ? 'wiswork_read'
        : 'wiswork_propose',
  }))
  const catalog = JSON.stringify(tools)
  if (Buffer.byteLength(catalog, 'utf8') > DOCUMENT_CATALOG_MAX_BYTES)
    throw new Error('document_tool_catalog_too_large')
  return `\nDocument semantic tool catalog (JSON): ${catalog}\nFor every call, put the selected tool's arguments only inside input.`
}

export function createProductionCodexBootstrap(
  options: ProductionCodexBootstrapOptions,
): CodexRuntimeBootstrap {
  return {
    async start({ executablePath, onCrash }): Promise<CodexRuntimeEngine> {
      const resolver = new CodexTurnResolver(options.diagnostics)
      let rejectDeterministicFailure: (code: string) => void = () => undefined
      let bridge: Awaited<ReturnType<typeof startResponsesBridge>> | undefined
      let gateway: Awaited<ReturnType<typeof startDynamicMcpGateway>> | undefined
      let manager: CodexProcessManager | undefined
      let client
      try {
        bridge = await startResponsesBridge({
          fetchWithAuth: options.fetchWithAuth,
          prepareTurn: resolver.prepare,
          diagnostics: options.diagnostics,
          onProtocolRecording: options.onProtocolRecording,
          onDeterministicFailure: (code) => rejectDeterministicFailure(code),
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
        readonly pendingProposals: Set<string>
        lastFailure?: Error
        deferredTerminal?: { status: 'completed' | 'cancelled' | 'failed'; error?: Error }
        proposalFailure?: 'cancelled' | 'failed'
        readonly touch: () => void
        readonly settle: (status: 'completed' | 'cancelled' | 'failed', error?: Error) => void
        readonly requestSettle: (
          status: 'completed' | 'cancelled' | 'failed',
          error?: Error,
        ) => void
      }
      const documents = new Map<
        string,
        {
          session: DocumentToolSession
          unregister: () => void
          onEvent?: (event: CodexRuntimeEngineEvent) => void
          instructions?: string
          threadId?: string
          active?: ActiveTurn
        }
      >()
      rejectDeterministicFailure = () => {
        const activeTurns = [...documents.values()].flatMap((document) =>
          document.active ? [document.active] : [],
        )
        // The bridge has no document identity. A single active turn is unambiguous; with
        // concurrent turns, let each app-server stream fail independently.
        if (activeTurns.length !== 1) return
        const active = activeTurns[0]!
        active.cancelled = true
        active.settle('failed', new Error('enhanced_response_incompatible'))
        if (active.threadId && active.turnId) {
          startBestEffortCodexInterrupt(() =>
            client.interruptTurn(active.threadId!, active.turnId!),
          )
        }
      }
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
        if (notification.method === 'error') {
          const failure = new Error(safeTurnFailure(notification.params))
          active.lastFailure = failure
          const willRetry = (notification.params as { willRetry?: unknown }).willRetry
          if (willRetry === false) {
            active.settle('failed', failure)
            queueMicrotask(onCrash)
          }
          return
        }
        if (notification.method === 'item/agentMessage/delta') {
          if (params.turnId === active.turnId && typeof params.delta === 'string') {
            active.touch()
            emit(document.onEvent, { type: 'text', text: params.delta })
          }
          return
        }
        if (notification.method === 'turn/completed') {
          const turn = params.turn
          if (turn?.id !== active.turnId) return
          active.touch()
          if (typeof turn?.status === 'string' && /^[a-z_]{1,32}$/.test(turn.status)) {
            options.diagnostics?.(`codex_turn_status_${turn.status}`)
          }
          active.requestSettle(
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
            onToolEvent: (event) => {
              documents.get(input.documentId)?.active?.touch()
              emit(input.onEvent, event)
            },
            onProposal: (proposal) => {
              const active = documents.get(input.documentId)?.active
              if (!active) return
              active.touch()
              active.pendingProposals.add(proposal.proposalId)
              void proposal.settled.then(
                (execution) => {
                  if (execution.isError) {
                    if (execution.output === 'mutation_cancelled') {
                      active.proposalFailure ??= 'cancelled'
                    } else {
                      active.proposalFailure = 'failed'
                    }
                  }
                  active.pendingProposals.delete(proposal.proposalId)
                  const deferred = active.deferredTerminal
                  if (deferred && active.pendingProposals.size === 0) {
                    active.deferredTerminal = undefined
                    const failure = active.proposalFailure
                    active.settle(
                      failure ?? deferred.status,
                      failure === 'failed' ? new Error('enhanced_proposal_failed') : deferred.error,
                    )
                  }
                },
                () => {
                  active.proposalFailure = 'failed'
                  active.pendingProposals.delete(proposal.proposalId)
                  if (active.deferredTerminal && active.pendingProposals.size === 0) {
                    active.deferredTerminal = undefined
                    active.settle('failed', new Error('enhanced_proposal_failed'))
                  }
                },
              )
              const { settled: _settled, ...publicProposal } = proposal
              emit(input.onEvent, { type: 'proposal', ...publicProposal })
            },
          })
          const entry: {
            session: DocumentToolSession
            unregister: () => void
            onEvent?: (event: CodexRuntimeEngineEvent) => void
            instructions?: string
            threadId?: string
            active?: ActiveTurn
          } = {
            session: input.session,
            unregister,
            onEvent: input.onEvent,
            instructions: `${input.instructions ?? ''}${buildDocumentToolInstructions(input.session)}`,
          }
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
            threadId: document.threadId ?? 'reserved',
          })
          let resolve!: () => void
          let reject!: (error: Error) => void
          let settled = false
          const terminal = new Promise<void>((onResolve, onReject) => {
            resolve = onResolve
            reject = onReject
          })
          const active: ActiveTurn = {
            capability: grant.capability,
            cancelled: false,
            pendingProposals: new Set(),
            touch: () => deadline.touch(),
            requestSettle(status, error) {
              if (active.pendingProposals.size > 0 && status === 'completed') {
                active.deferredTerminal = { status, error }
                return
              }
              active.settle(status, error)
            },
            settle(status, error) {
              if (settled) return
              settled = true
              deadline.disarm()
              gateway.revokeTurn(grant.capability, status !== 'completed')
              active.disarm?.()
              if (document.active === active) document.active = undefined
              emit(document.onEvent, { type: 'terminal', status })
              if (error) reject(error)
              else resolve()
            },
          }
          const deadline = createTurnIdleDeadline(() => {
            active.cancelled = true
            active.settle('failed', new Error('enhanced_turn_timeout'))
            if (active.threadId && active.turnId) {
              startBestEffortCodexInterrupt(() =>
                client.interruptTurn(active.threadId!, active.turnId!),
              )
            }
          })
          document.active = active
          try {
            if (!document.threadId) {
              document.threadId = (
                await client.startThread({
                  developerInstructions: `${DEVELOPER_POLICY}\n${document.instructions ?? ''}`,
                })
              ).thread.id
              gateway.bindTurn(grant.capability, document.threadId)
            }
            active.threadId = document.threadId
            active.touch()
            if (active.cancelled) return await terminal
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
            active.turnId = (
              await client.startTurn(
                active.threadId,
                `<wiswork_turn_capability>${grant.capability}</wiswork_turn_capability>\n\n${input.text}`,
              )
            ).turn.id
            active.touch()
            if (active.cancelled) return await terminal
            await terminal
          } catch (error) {
            if (active.cancelled) return await terminal
            active.settle(
              'failed',
              active.lastFailure ??
                (error instanceof Error ? error : new Error('enhanced_turn_failed')),
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
          active.settle('cancelled')
          if (active.threadId && active.turnId) {
            startBestEffortCodexInterrupt(() =>
              client.interruptTurn(active.threadId!, active.turnId!),
            )
          }
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
