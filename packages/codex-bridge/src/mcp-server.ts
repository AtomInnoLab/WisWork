import { createServer, type IncomingMessage, type ServerResponse } from 'node:http'
import type { AddressInfo } from 'node:net'
import type { AgentToolCall } from '@wiswork/agent-core'
import {
  DocumentToolRouter,
  ToolRouterError,
  type ToolSessionCredentials,
  type ToolSessionRegistration,
} from './tool-router.js'

const DEFAULT_BODY_LIMIT = 1_000_000
const PROTOCOL_VERSION = '2025-06-18'

type JsonRpcId = string | number
type JsonRecord = Record<string, unknown>

export interface McpServerDiagnostic {
  readonly code: string
}

export interface DocumentMcpServerOptions {
  readonly router?: DocumentToolRouter
  readonly maxBodyBytes?: number
  readonly diagnostics?: (diagnostic: McpServerDiagnostic) => void
}

export interface DocumentMcpSession {
  readonly url: string
  readonly secret: string
  close(): void
}

export interface DocumentMcpServer {
  readonly baseUrl: string
  register(registration: ToolSessionRegistration): DocumentMcpSession
  close(): Promise<void>
}

function isRecord(value: unknown): value is JsonRecord {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

function hasOnlyKeys(value: JsonRecord, allowed: readonly string[]): boolean {
  const permitted = new Set(allowed)
  return Object.keys(value).every((key) => permitted.has(key))
}

function validId(value: unknown): value is JsonRpcId {
  return (
    (typeof value === 'string' && value !== '' && Buffer.byteLength(value) <= 256) ||
    (typeof value === 'number' && Number.isSafeInteger(value))
  )
}

function sendJson(response: ServerResponse, status: number, body: unknown): void {
  const encoded = Buffer.from(JSON.stringify(body))
  response.writeHead(status, {
    'content-type': 'application/json',
    'content-length': String(encoded.length),
    'cache-control': 'no-store',
  })
  response.end(encoded)
}

function sendError(
  response: ServerResponse,
  id: JsonRpcId | null,
  code: number,
  message: string,
): void {
  sendJson(response, 200, { jsonrpc: '2.0', id, error: { code, message } })
}

function sendAccepted(response: ServerResponse): void {
  response.writeHead(202, { 'cache-control': 'no-store', 'content-length': '0' })
  response.end()
}

function parseCredentials(request: IncomingMessage): ToolSessionCredentials | undefined {
  const match = request.url?.match(/^\/mcp\/([A-Za-z0-9_-]{43})$/)
  const authorization = request.headers.authorization
  if (!match || typeof authorization !== 'string' || !authorization.startsWith('Bearer ')) {
    return undefined
  }
  const token = authorization.slice('Bearer '.length)
  if (!/^[A-Za-z0-9_-]{43}$/.test(token)) return undefined
  return { sessionId: match[1]!, token }
}

async function readBody(request: IncomingMessage, limit: number): Promise<Buffer | undefined> {
  const contentLength = request.headers['content-length']
  if (
    typeof contentLength === 'string' &&
    (!/^\d+$/.test(contentLength) || Number(contentLength) > limit)
  ) {
    request.resume()
    return undefined
  }
  const chunks: Buffer[] = []
  let bytes = 0
  for await (const chunk of request) {
    const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk)
    bytes += buffer.length
    if (bytes > limit) {
      request.resume()
      return undefined
    }
    chunks.push(buffer)
  }
  return Buffer.concat(chunks)
}

export async function startDocumentMcpServer(
  options: DocumentMcpServerOptions = {},
): Promise<DocumentMcpServer> {
  const maxBodyBytes = options.maxBodyBytes ?? DEFAULT_BODY_LIMIT
  if (!Number.isSafeInteger(maxBodyBytes) || maxBodyBytes <= 0) {
    throw new TypeError('invalid_mcp_body_limit')
  }
  const router = options.router ?? new DocumentToolRouter()
  let closed = false
  const emitDiagnostic = (code: string): void => {
    try {
      options.diagnostics?.({ code })
    } catch {
      // Diagnostics must never affect request handling.
    }
  }

  const server = createServer((request, response) => {
    void (async () => {
      response.setHeader('x-content-type-options', 'nosniff')
      if (closed) {
        sendJson(response, 503, { error: 'mcp_closed' })
        return
      }
      if (request.method !== 'POST') {
        sendJson(response, 405, { error: 'method_not_allowed' })
        return
      }
      const sessionCredentials = parseCredentials(request)
      if (!sessionCredentials) {
        request.resume()
        sendJson(response, 401, { error: 'unauthorized' })
        return
      }
      try {
        router.authorize(sessionCredentials)
      } catch {
        request.resume()
        sendJson(response, 401, { error: 'unauthorized' })
        return
      }
      if (!request.headers['content-type']?.toLowerCase().startsWith('application/json')) {
        sendJson(response, 415, { error: 'unsupported_media_type' })
        return
      }
      const raw = await readBody(request, maxBodyBytes)
      if (!raw) {
        emitDiagnostic('mcp_body_limit')
        sendJson(response, 413, { error: 'body_limit' })
        return
      }
      let message: unknown
      try {
        message = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(raw))
      } catch {
        emitDiagnostic('mcp_invalid_json')
        sendError(response, null, -32700, 'parse_error')
        return
      }
      if (
        !isRecord(message) ||
        !hasOnlyKeys(message, ['jsonrpc', 'id', 'method', 'params']) ||
        message.jsonrpc !== '2.0' ||
        typeof message.method !== 'string' ||
        (message.id !== undefined && !validId(message.id)) ||
        (message.params !== undefined && !isRecord(message.params))
      ) {
        sendError(response, null, -32600, 'invalid_request')
        return
      }
      const id = message.id as JsonRpcId | undefined
      const params = (message.params ?? {}) as JsonRecord
      try {
        if (message.method === 'notifications/initialized') {
          if (id !== undefined || !hasOnlyKeys(params, [])) {
            sendError(response, id ?? null, -32600, 'invalid_request')
          } else sendAccepted(response)
          return
        }
        if (message.method === 'notifications/cancelled') {
          if (
            id !== undefined ||
            !hasOnlyKeys(params, ['requestId', 'reason']) ||
            !validId(params.requestId) ||
            (params.reason !== undefined && typeof params.reason !== 'string')
          ) {
            sendError(response, id ?? null, -32600, 'invalid_request')
          } else {
            router.cancel(sessionCredentials, String(params.requestId))
            sendAccepted(response)
          }
          return
        }
        if (id === undefined) {
          sendError(response, null, -32600, 'invalid_request')
          return
        }
        if (message.method === 'initialize') {
          const capabilities = params.capabilities
          const clientInfo = params.clientInfo
          if (
            !hasOnlyKeys(params, ['protocolVersion', 'capabilities', 'clientInfo']) ||
            params.protocolVersion !== PROTOCOL_VERSION ||
            !isRecord(capabilities) ||
            !hasOnlyKeys(capabilities, ['elicitation']) ||
            !isRecord(capabilities.elicitation) ||
            !hasOnlyKeys(capabilities.elicitation, ['form', 'url']) ||
            !isRecord(capabilities.elicitation.form) ||
            !isRecord(capabilities.elicitation.url) ||
            !isRecord(clientInfo) ||
            !hasOnlyKeys(clientInfo, ['name', 'title', 'version']) ||
            clientInfo.name !== 'codex-mcp-client' ||
            clientInfo.title !== 'Codex' ||
            clientInfo.version !== '0.147.0'
          ) {
            sendError(response, id, -32602, 'invalid_params')
            return
          }
          router.listTools(sessionCredentials)
          sendJson(response, 200, {
            jsonrpc: '2.0',
            id,
            result: {
              protocolVersion: PROTOCOL_VERSION,
              capabilities: { tools: {} },
              serverInfo: { name: 'wiswork', version: '0.1.0' },
            },
          })
          return
        }
        if (message.method === 'tools/list') {
          if (!hasOnlyKeys(params, ['cursor', '_meta'])) {
            sendError(response, id, -32602, 'invalid_params')
            return
          }
          sendJson(response, 200, {
            jsonrpc: '2.0',
            id,
            result: { tools: router.listTools(sessionCredentials) },
          })
          return
        }
        if (message.method === 'tools/call') {
          if (
            !hasOnlyKeys(params, ['name', 'arguments', '_meta']) ||
            typeof params.name !== 'string' ||
            !isRecord(params.arguments ?? {})
          ) {
            sendError(response, id, -32602, 'invalid_params')
            return
          }
          const call: AgentToolCall = {
            id: String(id),
            name: params.name,
            input: (params.arguments ?? {}) as Record<string, unknown>,
          }
          const controller = new AbortController()
          const onDisconnect = (): void => controller.abort()
          request.once('aborted', onDisconnect)
          response.once('close', onDisconnect)
          const execution = await router.callTool(sessionCredentials, call, controller.signal)
          request.off('aborted', onDisconnect)
          response.off('close', onDisconnect)
          if (response.destroyed) return
          sendJson(response, 200, {
            jsonrpc: '2.0',
            id,
            result: {
              content: [{ type: 'text', text: execution.output }],
              isError: execution.isError === true,
            },
          })
          return
        }
        sendError(response, id, -32601, 'method_not_found')
      } catch (error) {
        if (error instanceof ToolRouterError && error.code === 'mcp_unauthorized') {
          sendJson(response, 401, { error: 'unauthorized' })
        } else {
          emitDiagnostic('mcp_request_failed')
          sendError(response, id ?? null, -32603, 'internal_error')
        }
      }
    })().catch(() => {
      emitDiagnostic('mcp_request_failed')
      if (!response.headersSent) sendJson(response, 500, { error: 'internal_error' })
      else response.end()
    })
  })

  await new Promise<void>((resolve, reject) => {
    const onError = (): void => reject(new Error('mcp_bind_failed'))
    server.once('error', onError)
    server.listen(0, '127.0.0.1', () => {
      server.off('error', onError)
      resolve()
    })
  })
  const address = server.address() as AddressInfo
  const baseUrl = `http://127.0.0.1:${address.port}`

  return {
    baseUrl,
    register(registration) {
      if (closed) throw new ToolRouterError('mcp_server_closed')
      const sessionCredentials = router.register(registration)
      let sessionClosed = false
      return {
        url: `${baseUrl}/mcp/${sessionCredentials.sessionId}`,
        secret: sessionCredentials.token,
        close() {
          if (sessionClosed) return
          sessionClosed = true
          try {
            router.close(sessionCredentials)
          } catch {
            // Already closed by server teardown.
          }
        },
      }
    },
    async close() {
      if (closed) return
      closed = true
      router.closeAll()
      await new Promise<void>((resolve, reject) => {
        server.close((error) => (error ? reject(new Error('mcp_close_failed')) : resolve()))
      })
    },
  }
}
