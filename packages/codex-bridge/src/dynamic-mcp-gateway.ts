import { randomBytes } from 'node:crypto'
import {
  isToolExecutionSuspension,
  type AgentToolCall,
  type ToolExecution,
} from '@wiswork/agent-core'
import type { DocumentToolSession } from './tool-router.js'
import { startTrustedMcpTransport, TrustedMcpTransportDenied } from './mcp-server.js'

const MAX_CALLS = 8
const DEFAULT_TTL_MS = 10 * 60_000
const MAX_ACTIVE_GRANTS = 64
const MAX_PROPOSAL_TTL_MS = 30_000

export interface DynamicGatewayDocument {
  readonly ownerId: string
  readonly documentId: string
  readonly generation: number
  readonly session: DocumentToolSession
  readonly onProposal?: (
    proposal: Readonly<{
      proposalId: string
      call: AgentToolCall
      expiresAt: number
    }>,
  ) => void
  readonly onToolEvent?: (
    event:
      | Readonly<{ type: 'tool-start'; callId: string; toolName: string }>
      | Readonly<{ type: 'tool-complete'; callId: string; toolName: string; isError: boolean }>,
  ) => void
}

export interface DynamicMcpGateway {
  readonly url: string
  readonly secret: string
  register(document: DynamicGatewayDocument): () => void
  beginTurn(input: {
    readonly documentId: string
    readonly generation: number
    readonly threadId: string
    readonly ttlMs?: number
  }): { readonly capability: string }
  bindTurn(capability: string, threadId: string): void
  revokeTurn(capability: string, cancelPending?: boolean): void
  close(): Promise<void>
}

interface TurnGrant {
  readonly document: DynamicGatewayDocument
  threadId: string
  readonly expiresAt: number
  readonly calls: Set<string>
  remaining: number
  turnId?: string
  bound: boolean
}

function exactRecord(value: unknown, keys: readonly string[]): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value))
    throw new Error('invalid')
  if (
    Object.getPrototypeOf(value) !== Object.prototype ||
    Object.getOwnPropertySymbols(value).length
  )
    throw new Error('invalid')
  const descriptors = Object.getOwnPropertyDescriptors(value)
  if (
    Object.keys(descriptors).length !== keys.length ||
    keys.some((key) => !('value' in (descriptors[key] ?? {})))
  )
    throw new Error('invalid')
  return Object.fromEntries(keys.map((key) => [key, descriptors[key]!.value]))
}

export async function startDynamicMcpGateway(
  diagnostics?: (code: string) => void,
): Promise<DynamicMcpGateway> {
  const diagnostic = (code: string): void => {
    try {
      diagnostics?.(code)
    } catch {}
  }
  const emitTool = (
    document: DynamicGatewayDocument,
    event: Parameters<NonNullable<DynamicGatewayDocument['onToolEvent']>>[0],
  ): void => {
    try {
      document.onToolEvent?.(event)
    } catch {}
  }
  const documents = new Map<string, DynamicGatewayDocument>()
  const grants = new Map<string, TurnGrant>()
  const revokeGrant = (capability: string, cancelPending = true): void => {
    const grant = grants.get(capability)
    grants.delete(capability)
    if (grant && cancelPending) {
      try {
        grant.document.session.cancelAll(grant.document.session.credentials)
      } catch {}
    }
  }
  const sweepExpired = (): void => {
    const now = Date.now()
    for (const [capability, grant] of grants) {
      if (grant.expiresAt <= now) revokeGrant(capability)
    }
  }
  let closed = false
  const credentials = Object.freeze({
    sessionId: randomBytes(32).toString('base64url'),
    secret: randomBytes(32).toString('base64url'),
  })
  const proxySession: DocumentToolSession = {
    identity: Object.freeze({
      ownerId: 'shell',
      host: 'docs',
      documentId: 'dynamic',
      sessionId: credentials.sessionId,
      generation: 0,
    }),
    credentials,
    catalogDigest: '0'.repeat(64),
    mutationAuthority: {
      claimNext: () => undefined,
      settle: () => {
        throw new Error('unsupported')
      },
      reject: () => {
        throw new Error('unsupported')
      },
    },
    authorize(candidate) {
      if (candidate.sessionId !== credentials.sessionId || candidate.secret !== credentials.secret)
        throw new Error('tool_unauthorized')
    },
    listTools() {
      diagnostic('gateway_tools_list')
      return [
        {
          name: 'wiswork_read',
          description: 'Execute one authorized read-only WisWork document tool.',
          annotations: { readOnlyHint: true, destructiveHint: false },
          inputSchema: {
            type: 'object',
            additionalProperties: false,
            required: ['capability', 'callId', 'toolName', 'input'],
            properties: {
              capability: { type: 'string' },
              callId: { type: 'string' },
              toolName: { type: 'string' },
              input: { type: 'object' },
            },
          },
        },
        {
          name: 'wiswork_propose',
          description: 'Create an opaque pending WisWork proposal without changing the document.',
          annotations: { readOnlyHint: true, destructiveHint: false },
          inputSchema: {
            type: 'object',
            additionalProperties: false,
            required: ['capability', 'callId', 'toolName', 'input'],
            properties: {
              capability: { type: 'string' },
              callId: { type: 'string' },
              toolName: { type: 'string' },
              input: { type: 'object' },
            },
          },
        },
      ]
    },
    async callTool(_candidate, call) {
      let started:
        Readonly<{ document: DynamicGatewayDocument; callId: string; toolName: string }> | undefined
      try {
        diagnostic('gateway_tool_call_received')
        if (call.name !== 'wiswork_read' && call.name !== 'wiswork_propose')
          throw new Error('denied')
        const args = exactRecord(call.input, ['capability', 'callId', 'toolName', 'input'])
        if (
          typeof args.capability !== 'string' ||
          typeof args.callId !== 'string' ||
          typeof args.toolName !== 'string'
        )
          throw new Error('denied')
        sweepExpired()
        const grant = grants.get(args.capability)
        // Capability is consumed/budgeted before any document or host lookup.
        if (
          !grant ||
          !grant.bound ||
          Date.now() > grant.expiresAt ||
          grant.remaining <= 0 ||
          grant.calls.has(args.callId)
        )
          throw new Error('denied')
        grant.remaining -= 1
        grant.calls.add(args.callId)
        const documentCall: AgentToolCall = {
          id: args.callId,
          name: args.toolName,
          input: args.input as Record<string, unknown>,
        }
        const definition = grant.document.session
          .listTools(grant.document.session.credentials)
          .find((tool) => tool.name === documentCall.name)
        const isRead =
          definition?.annotations?.readOnlyHint === true &&
          definition.annotations.destructiveHint !== true
        if ((call.name === 'wiswork_read') !== isRead) throw new Error('denied')
        emitTool(grant.document, {
          type: 'tool-start',
          callId: documentCall.id,
          toolName: documentCall.name,
        })
        started = {
          document: grant.document,
          callId: documentCall.id,
          toolName: documentCall.name,
        }
        const outcome = grant.document.session.callTool(
          grant.document.session.credentials,
          documentCall,
        )
        if (call.name === 'wiswork_propose') {
          if (outcome instanceof Promise) throw new Error('denied')
          if (!isToolExecutionSuspension(outcome)) throw new Error('denied')
          const proposalId = randomBytes(32).toString('base64url')
          if (!grant.document.onProposal) throw new Error('denied')
          grant.document.onProposal({
            proposalId,
            call: documentCall,
            expiresAt: Math.min(grant.expiresAt, Date.now() + MAX_PROPOSAL_TTL_MS),
          })
          const execution: ToolExecution = {
            output: JSON.stringify({ proposalId, status: 'pending_confirmation' }),
            summary: 'Proposal pending confirmation',
            mutated: false,
          }
          emitTool(grant.document, {
            type: 'tool-complete',
            callId: documentCall.id,
            toolName: documentCall.name,
            isError: false,
          })
          diagnostic('gateway_proposal_created')
          diagnostic('gateway_tool_call_completed')
          started = undefined
          return execution
        }
        const execution: ToolExecution =
          outcome instanceof Promise
            ? await outcome
            : isToolExecutionSuspension(outcome)
              ? await outcome.result
              : outcome
        emitTool(grant.document, {
          type: 'tool-complete',
          callId: documentCall.id,
          toolName: documentCall.name,
          isError: execution.isError === true,
        })
        diagnostic('gateway_tool_call_completed')
        started = undefined
        return execution
      } catch {
        if (started) {
          emitTool(started.document, {
            type: 'tool-complete',
            callId: started.callId,
            toolName: started.toolName,
            isError: true,
          })
        }
        diagnostic('gateway_tool_call_denied')
        throw new TrustedMcpTransportDenied('turn_capability_denied')
      }
    },
    issueCarrier() {
      throw new Error('unsupported')
    },
    cancel: () => false,
    cancelAll: () => 0,
    close: () => undefined,
  }
  const transport = await startTrustedMcpTransport(proxySession, { diagnostics })
  return Object.freeze({
    url: transport.url,
    secret: transport.secret,
    register(document: DynamicGatewayDocument) {
      if (closed || documents.has(document.documentId))
        throw new Error('document_session_unavailable')
      documents.set(document.documentId, document)
      return () => {
        if (documents.get(document.documentId) === document) documents.delete(document.documentId)
        for (const [token, grant] of grants) if (grant.document === document) revokeGrant(token)
      }
    },
    beginTurn(input: {
      readonly documentId: string
      readonly generation: number
      readonly threadId: string
      readonly ttlMs?: number
    }) {
      if (closed) throw new Error('gateway_closed')
      sweepExpired()
      if (grants.size >= MAX_ACTIVE_GRANTS) throw new Error('turn_capability_limit')
      const document = documents.get(input.documentId)
      if (!document || document.generation !== input.generation)
        throw new Error('document_session_unavailable')
      const ttl = input.ttlMs ?? DEFAULT_TTL_MS
      if (!Number.isSafeInteger(ttl) || ttl <= 0 || ttl > DEFAULT_TTL_MS)
        throw new Error('invalid_turn_capability')
      const capability = randomBytes(32).toString('base64url')
      grants.set(capability, {
        document,
        threadId: input.threadId,
        expiresAt: Date.now() + ttl,
        calls: new Set(),
        remaining: MAX_CALLS,
        bound: input.threadId !== 'reserved',
      })
      return Object.freeze({ capability })
    },
    bindTurn(capability: string, threadId: string) {
      sweepExpired()
      const grant = grants.get(capability)
      if (!grant || grant.bound || !threadId || Buffer.byteLength(threadId) > 256)
        throw new Error('invalid_turn_capability')
      grant.threadId = threadId
      grant.bound = true
    },
    revokeTurn(capability: string, cancelPending = true) {
      revokeGrant(capability, cancelPending)
    },
    async close() {
      if (closed) return
      closed = true
      for (const capability of [...grants.keys()]) revokeGrant(capability)
      documents.clear()
      await transport.close()
    },
  })
}
