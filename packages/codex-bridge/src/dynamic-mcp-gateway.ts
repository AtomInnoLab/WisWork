import { randomBytes, timingSafeEqual } from 'node:crypto'
import { createServer } from 'node:http'
import type { AddressInfo } from 'node:net'
import {
  isToolExecutionSuspension,
  type AgentToolCall,
  type ToolExecution,
} from '@wiswork/agent-core'
import type { DocumentToolSession } from './tool-router.js'

const HOST = '127.0.0.1'
const MAX_BODY = 1_000_000
const MAX_CALLS = 16
const DEFAULT_TTL_MS = 10 * 60_000

export interface DynamicGatewayDocument {
  readonly ownerId: string
  readonly documentId: string
  readonly generation: number
  readonly session: DocumentToolSession
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
  close(): Promise<void>
}

interface TurnGrant {
  readonly document: DynamicGatewayDocument
  readonly threadId: string
  readonly expiresAt: number
  readonly calls: Set<string>
  remaining: number
  turnId?: string
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

function authorized(header: string | undefined, expected: Buffer): boolean {
  if (!header?.startsWith('Bearer ')) return false
  const raw = header.slice(7)
  if (!/^[A-Za-z0-9_-]{43}$/.test(raw)) return false
  const value = Buffer.from(raw, 'base64url')
  return value.length === expected.length && timingSafeEqual(value, expected)
}

export async function startDynamicMcpGateway(): Promise<DynamicMcpGateway> {
  const secretBytes = randomBytes(32)
  const secret = secretBytes.toString('base64url')
  const documents = new Map<string, DynamicGatewayDocument>()
  const grants = new Map<string, TurnGrant>()
  let closed = false
  const server = createServer((request, response) => {
    void (async () => {
      const send = (status: number, value: unknown) => {
        const body = Buffer.from(JSON.stringify(value))
        response.writeHead(status, {
          'content-type': 'application/json',
          'content-length': body.length,
        })
        response.end(body)
      }
      if (closed) return send(503, { error: 'gateway_closed' })
      if (!authorized(request.headers.authorization, secretBytes)) {
        request.resume()
        return send(401, { error: 'unauthorized' })
      }
      const chunks: Buffer[] = []
      let bytes = 0
      for await (const chunk of request) {
        const part = Buffer.from(chunk)
        bytes += part.length
        if (bytes > MAX_BODY) return send(413, { error: 'body_limit' })
        chunks.push(part)
      }
      let rpc: Record<string, unknown>
      try {
        const parsed = JSON.parse(Buffer.concat(chunks).toString('utf8')) as unknown
        if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed))
          throw new Error()
        const descriptors = Object.getOwnPropertyDescriptors(parsed)
        const keys = Object.keys(descriptors)
        if (
          Object.getPrototypeOf(parsed) !== Object.prototype ||
          Object.getOwnPropertySymbols(parsed).length ||
          keys.some((key) => !['jsonrpc', 'id', 'method', 'params'].includes(key)) ||
          !['jsonrpc', 'method', 'params'].every((key) => 'value' in (descriptors[key] ?? {}))
        )
          throw new Error()
        rpc = Object.fromEntries(keys.map((key) => [key, descriptors[key]!.value]))
        if (rpc.jsonrpc !== '2.0' || typeof rpc.method !== 'string') throw new Error()
      } catch {
        return send(400, { error: 'invalid_request' })
      }
      const ok = (result: unknown) => send(200, { jsonrpc: '2.0', id: rpc.id, result })
      if (rpc.method === 'initialize')
        return ok({
          protocolVersion: '2025-06-18',
          capabilities: { tools: {} },
          serverInfo: { name: 'wiswork', version: '1' },
        })
      if (rpc.method === 'notifications/initialized') return send(202, {})
      if (rpc.method === 'tools/list')
        return ok({
          tools: [
            {
              name: 'wiswork_call',
              description: 'Execute one authorized WisWork document tool call.',
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
          ],
        })
      if (rpc.method !== 'tools/call') return send(404, { error: 'method_not_found' })
      try {
        const params = exactRecord(rpc.params, ['name', 'arguments'])
        if (params.name !== 'wiswork_call') throw new Error('denied')
        const args = exactRecord(params.arguments, ['capability', 'callId', 'toolName', 'input'])
        if (
          typeof args.capability !== 'string' ||
          typeof args.callId !== 'string' ||
          typeof args.toolName !== 'string'
        )
          throw new Error('denied')
        const grant = grants.get(args.capability)
        // Capability is consumed/budgeted before any document or host lookup.
        if (
          !grant ||
          Date.now() > grant.expiresAt ||
          grant.remaining <= 0 ||
          grant.calls.has(args.callId)
        )
          throw new Error('denied')
        grant.remaining -= 1
        grant.calls.add(args.callId)
        const call: AgentToolCall = {
          id: args.callId,
          name: args.toolName,
          input: args.input as Record<string, unknown>,
        }
        const outcome = grant.document.session.callTool(grant.document.session.credentials, call)
        const execution: ToolExecution =
          outcome instanceof Promise
            ? await outcome
            : isToolExecutionSuspension(outcome)
              ? await outcome.result
              : outcome
        return ok({
          content: [{ type: 'text', text: execution.output }],
          isError: execution.isError === true,
        })
      } catch {
        return send(403, { error: 'turn_capability_denied' })
      }
    })().catch(() => response.destroy())
  })
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject)
    server.listen(0, HOST, resolve)
  })
  const url = `http://${HOST}:${(server.address() as AddressInfo).port}/mcp`
  return Object.freeze({
    url,
    secret,
    register(document: DynamicGatewayDocument) {
      if (closed || documents.has(document.documentId))
        throw new Error('document_session_unavailable')
      documents.set(document.documentId, document)
      return () => {
        if (documents.get(document.documentId) === document) documents.delete(document.documentId)
        for (const [token, grant] of grants) if (grant.document === document) grants.delete(token)
      }
    },
    beginTurn(input: {
      readonly documentId: string
      readonly generation: number
      readonly threadId: string
      readonly ttlMs?: number
    }) {
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
      })
      return Object.freeze({ capability })
    },
    async close() {
      if (closed) return
      closed = true
      grants.clear()
      documents.clear()
      await new Promise<void>((resolve) => server.close(() => resolve()))
    },
  })
}
