import { createServer, type IncomingMessage, type ServerResponse } from 'node:http'
import type { AddressInfo } from 'node:net'
import type { AgentToolCall } from '@wiswork/agent-core'
import type { DocumentCarrierHandle, DocumentCarrierTurnContext } from './types.js'
import {
  createDocumentToolSession,
  ToolRouterError,
  type DocumentToolRegistration,
  type DocumentToolSession,
} from './tool-router.js'

const HOST = '127.0.0.1'
const MAX_BODY = 1_000_000
const MAX_RPC_CALLS = 1_024
const PROTOCOL = '2025-06-18'
type RpcId = string | number

export interface DocumentMcpSession {
  readonly url: string
  readonly secret: string
  issueCarrier(turn: Omit<DocumentCarrierTurnContext, 'capability'>): DocumentCarrierHandle
  close(): void
}
export interface DocumentMcpServer {
  readonly baseUrl: string
  register(registration: DocumentToolRegistration): DocumentMcpSession
  close(): Promise<void>
}
export interface DocumentMcpServerOptions {
  readonly maxBodyBytes?: number
  readonly maxRpcCalls?: number
  readonly diagnostics?: (code: string) => void
}

function record(value: unknown): value is Record<string, unknown> {
  return (
    typeof value === 'object' &&
    value !== null &&
    !Array.isArray(value) &&
    (Object.getPrototypeOf(value) === Object.prototype || Object.getPrototypeOf(value) === null)
  )
}
function only(value: Record<string, unknown>, keys: readonly string[]): boolean {
  return Object.keys(value).every((key) => keys.includes(key))
}
function validId(value: unknown): value is RpcId {
  return (
    (typeof value === 'string' && value.length > 0 && Buffer.byteLength(value) <= 256) ||
    (typeof value === 'number' && Number.isSafeInteger(value))
  )
}
function send(response: ServerResponse, status: number, value: unknown): void {
  const body = Buffer.from(JSON.stringify(value))
  response.writeHead(status, {
    'cache-control': 'no-store',
    'content-type': 'application/json; charset=utf-8',
    'content-length': String(body.length),
    'x-content-type-options': 'nosniff',
  })
  response.end(body)
}
function rpcError(response: ServerResponse, id: RpcId | null, code: number, message: string): void {
  send(response, 200, { jsonrpc: '2.0', id, error: { code, message } })
}
async function body(request: IncomingMessage, maximum: number): Promise<Buffer | undefined> {
  const declared = request.headers['content-length']
  if (declared !== undefined && (!/^\d+$/.test(declared) || Number(declared) > maximum)) {
    request.resume()
    return undefined
  }
  const chunks: Buffer[] = []
  let bytes = 0
  for await (const chunk of request) {
    const buffer = Buffer.from(chunk)
    bytes += buffer.length
    if (bytes > maximum) {
      request.resume()
      return undefined
    }
    chunks.push(buffer)
  }
  return Buffer.concat(chunks, bytes)
}

export async function startDocumentMcpServer(
  options: DocumentMcpServerOptions = {},
): Promise<DocumentMcpServer> {
  const maxBodyBytes = options.maxBodyBytes ?? MAX_BODY
  const maxRpcCalls = options.maxRpcCalls ?? MAX_RPC_CALLS
  if (!Number.isSafeInteger(maxBodyBytes) || maxBodyBytes <= 0 || maxBodyBytes > MAX_BODY)
    throw new TypeError('invalid_mcp_body_limit')
  if (!Number.isSafeInteger(maxRpcCalls) || maxRpcCalls <= 0 || maxRpcCalls > MAX_RPC_CALLS)
    throw new TypeError('invalid_mcp_call_limit')
  const sessions = new Map<
    string,
    { session: DocumentToolSession; state: 'new' | 'initialized' | 'ready'; ids: Set<string> }
  >()
  const sockets = new Set<import('node:net').Socket>()
  let closed = false
  const diagnostic = (code: string): void => {
    try {
      options.diagnostics?.(code)
    } catch {
      /* detached */
    }
  }
  const server = createServer((request, response) => {
    void (async () => {
      if (closed) return send(response, 503, { error: 'mcp_closed' })
      if (request.method !== 'POST') {
        request.resume()
        return send(response, 405, { error: 'method_not_allowed' })
      }
      const match = request.url?.match(/^\/mcp\/([A-Za-z0-9_-]{43})$/)
      const auth = request.headers.authorization
      const entry = match ? sessions.get(match[1]!) : undefined
      const session = entry?.session
      const credentials =
        session && typeof auth === 'string' && auth.startsWith('Bearer ')
          ? { sessionId: match![1]!, secret: auth.slice(7) }
          : undefined
      try {
        if (!session || !credentials) throw new ToolRouterError('tool_unauthorized')
        session.authorize(credentials)
      } catch {
        request.resume()
        return send(response, 401, { error: 'unauthorized' })
      }
      if (request.headers['content-type']?.toLowerCase() !== 'application/json') {
        request.resume()
        return send(response, 415, { error: 'unsupported_media_type' })
      }
      const raw = await body(request, maxBodyBytes)
      if (!raw) {
        diagnostic('mcp_body_limit')
        return send(response, 413, { error: 'body_limit' })
      }
      let message: unknown
      try {
        message = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(raw))
      } catch {
        diagnostic('mcp_invalid_json')
        return rpcError(response, null, -32700, 'parse_error')
      }
      if (
        !record(message) ||
        !only(message, ['jsonrpc', 'id', 'method', 'params']) ||
        message.jsonrpc !== '2.0' ||
        typeof message.method !== 'string' ||
        (message.id !== undefined && !validId(message.id)) ||
        (message.params !== undefined && !record(message.params))
      )
        return rpcError(response, null, -32600, 'invalid_request')
      const id = message.id as RpcId | undefined
      const params = (message.params ?? {}) as Record<string, unknown>
      if (id !== undefined) {
        const idKey = `${typeof id}:${String(id)}`
        if (entry!.ids.has(idKey)) return rpcError(response, id, -32600, 'request_id_consumed')
        if (entry!.ids.size >= maxRpcCalls) {
          session.close()
          sessions.delete(session.credentials.sessionId)
          return rpcError(response, id, -32600, 'session_call_limit')
        }
        entry!.ids.add(idKey)
      }
      if (message.method === 'notifications/initialized') {
        if (
          entry!.state !== 'initialized' ||
          id !== undefined ||
          !only(params, []) ||
          Object.keys(params).length !== 0
        )
          return rpcError(response, id ?? null, -32600, 'invalid_request')
        entry!.state = 'ready'
        response.writeHead(202, { 'cache-control': 'no-store', 'content-length': '0' })
        return response.end()
      }
      if (message.method === 'notifications/cancelled') {
        if (
          entry!.state !== 'ready' ||
          id !== undefined ||
          !only(params, ['requestId', 'reason']) ||
          !validId(params.requestId)
        )
          return rpcError(response, id ?? null, -32600, 'invalid_request')
        session.cancel(credentials, String(params.requestId))
        response.writeHead(202, { 'content-length': '0' })
        return response.end()
      }
      if (id === undefined) return rpcError(response, null, -32600, 'invalid_request')
      if (message.method === 'initialize') {
        const capabilities = params.capabilities
        const clientInfo = params.clientInfo
        if (
          !only(params, ['protocolVersion', 'capabilities', 'clientInfo']) ||
          Object.keys(params).length !== 3 ||
          params.protocolVersion !== PROTOCOL ||
          !record(capabilities) ||
          !only(capabilities, ['elicitation']) ||
          !record(capabilities.elicitation) ||
          !only(capabilities.elicitation, ['form', 'url']) ||
          !record(capabilities.elicitation.form) ||
          !record(capabilities.elicitation.url) ||
          !record(clientInfo) ||
          !only(clientInfo, ['name', 'title', 'version']) ||
          clientInfo.name !== 'codex-mcp-client' ||
          clientInfo.title !== 'Codex' ||
          clientInfo.version !== '0.147.0' ||
          entry!.state !== 'new'
        )
          return rpcError(response, id, -32602, 'invalid_params')
        entry!.state = 'initialized'
        return send(response, 200, {
          jsonrpc: '2.0',
          id,
          result: {
            protocolVersion: PROTOCOL,
            capabilities: { tools: {} },
            serverInfo: { name: 'wiswork', version: '0.1.0' },
          },
        })
      }
      if (message.method === 'tools/list') {
        if (entry!.state !== 'ready') return rpcError(response, id, -32600, 'not_initialized')
        if (!only(params, ['cursor', '_meta']))
          return rpcError(response, id, -32602, 'invalid_params')
        return send(response, 200, {
          jsonrpc: '2.0',
          id,
          result: { tools: session.listTools(credentials) },
        })
      }
      if (message.method === 'tools/call') {
        if (entry!.state !== 'ready') return rpcError(response, id, -32600, 'not_initialized')
        if (
          !only(params, ['name', 'arguments', '_meta']) ||
          typeof params.name !== 'string' ||
          !record(params.arguments ?? {})
        )
          return rpcError(response, id, -32602, 'invalid_params')
        const controller = new AbortController()
        const abort = (): void => controller.abort()
        request.once('aborted', abort)
        response.once('close', abort)
        const call: AgentToolCall = {
          id: String(id),
          name: params.name,
          input: (params.arguments ?? {}) as Record<string, unknown>,
        }
        const execution = await session.callTool(credentials, call, controller.signal)
        request.off('aborted', abort)
        response.off('close', abort)
        if (!response.destroyed)
          send(response, 200, {
            jsonrpc: '2.0',
            id,
            result: {
              content: [{ type: 'text', text: execution.output }],
              isError: execution.isError === true,
            },
          })
        return
      }
      return rpcError(response, id, -32601, 'method_not_found')
    })().catch(() => {
      diagnostic('mcp_request_failed')
      if (!response.headersSent) send(response, 500, { error: 'internal_error' })
      else response.end()
    })
  })
  server.headersTimeout = 10_000
  server.requestTimeout = 30_000
  server.keepAliveTimeout = 5_000
  server.maxRequestsPerSocket = 100
  server.on('connection', (socket) => {
    sockets.add(socket)
    socket.setTimeout(30_000, () => socket.destroy())
    socket.once('close', () => sockets.delete(socket))
  })
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject)
    server.listen(0, HOST, resolve)
  })
  const baseUrl = `http://${HOST}:${(server.address() as AddressInfo).port}`
  const api: DocumentMcpServer = {
    baseUrl,
    register(registration: DocumentToolRegistration) {
      if (closed) throw new ToolRouterError('mcp_server_closed')
      const session = createDocumentToolSession(registration)
      sessions.set(session.credentials.sessionId, { session, state: 'new', ids: new Set() })
      let ended = false
      return Object.freeze({
        url: `${baseUrl}/mcp/${session.credentials.sessionId}`,
        secret: session.credentials.secret,
        issueCarrier(turn: Omit<DocumentCarrierTurnContext, 'capability'>) {
          return session.issueCarrier(session.credentials, turn)
        },
        close() {
          if (ended) return
          ended = true
          session.close()
          sessions.delete(session.credentials.sessionId)
        },
      })
    },
    async close() {
      if (closed) return
      closed = true
      for (const entry of sessions.values()) entry.session.close()
      sessions.clear()
      for (const socket of sockets) socket.destroy()
      await new Promise<void>((resolve) => server.close(() => resolve()))
    },
  }
  return Object.freeze(api)
}
