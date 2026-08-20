import type {
  AgentMessage,
  AgentStreamCallbacks,
  AgentStreamRequest,
  AgentToolCall,
  AgentTransport,
} from '@wiswork/agent-core'
import {
  WISWORK_DEFAULT_MODEL,
  WISWORK_MESSAGES_URL,
  WISWORK_REQUEST_LOCATION,
} from '@wiswork/ai-provider'
import type { BrowserAuth } from '../auth/browser-auth.js'
import type { RuntimeConfig } from '../config.js'

const MAX_TOKENS = 8192
export const MAX_STREAM_TOOL_INPUT_LENGTH = 16 * 1024
export const MAX_REQUEST_BODY_LENGTH = 256 * 1024
const MAX_SSE_LINE_LENGTH = 64 * 1024
const MAX_PENDING_TOOL_CALLS = 16

type TransportErrorCode =
  | 'transport_http'
  | 'transport_invalid_stream'
  | 'transport_stream_error'
  | 'transport_stream_too_large'
  | 'transport_tool_input_too_large'
  | 'transport_request_too_large'

class TransportError extends Error {
  constructor(
    readonly code: TransportErrorCode,
    readonly status?: number,
  ) {
    super(code)
    this.name = 'TransportError'
  }

  publicMessage(): string {
    return this.code === 'transport_http' ? `transport_http_${this.status ?? 0}` : this.code
  }
}

function trustedMessagesUrl(value: string): boolean {
  try {
    return new URL(value).href === new URL(WISWORK_MESSAGES_URL).href
  } catch {
    return false
  }
}

function messagesForProvider(messages: AgentMessage[]): unknown[] {
  return messages.map((message) => {
    if (message.role === 'user') return { role: 'user', content: message.text }
    if (message.role === 'assistant') {
      return {
        role: 'assistant',
        content: [
          ...(message.text ? [{ type: 'text', text: message.text }] : []),
          ...(message.toolCalls ?? []).map((call) => ({
            type: 'tool_use',
            id: call.id,
            name: call.name,
            input: call.input,
          })),
        ],
      }
    }
    return {
      role: 'user',
      content: message.results.map((result) => ({
        type: 'tool_result',
        tool_use_id: result.id,
        content: result.output,
        ...(result.isError ? { is_error: true } : {}),
      })),
    }
  })
}

function safeError(error: unknown): string {
  if (error && typeof error === 'object' && 'code' in error) {
    const code = Reflect.get(error, 'code')
    if (code === 'unauthenticated' || code === 'refresh_failed' || code === 'unauthorized') {
      return 'transport_auth'
    }
    if (code === 'invalid_destination') return 'transport_destination'
  }
  return 'transport_network'
}

async function* boundedSseLines(body: ReadableStream<Uint8Array>): AsyncGenerator<string> {
  const decoder = new TextDecoder()
  const reader = body.getReader()
  let buffer = ''
  while (true) {
    const { done, value } = await reader.read()
    if (done) break
    buffer += decoder.decode(value, { stream: true })
    let newline = buffer.indexOf('\n')
    while (newline !== -1) {
      const line = buffer.slice(0, newline).replace(/\r$/, '')
      if (line.length > MAX_SSE_LINE_LENGTH) throw new TransportError('transport_stream_too_large')
      yield line
      buffer = buffer.slice(newline + 1)
      newline = buffer.indexOf('\n')
    }
    if (buffer.length > MAX_SSE_LINE_LENGTH) {
      throw new TransportError('transport_stream_too_large')
    }
  }
  buffer += decoder.decode()
  if (buffer.length > MAX_SSE_LINE_LENGTH) throw new TransportError('transport_stream_too_large')
  if (buffer) yield buffer.replace(/\r$/, '')
}

async function consumeStream(response: Response, callbacks: AgentStreamCallbacks): Promise<void> {
  if (!response.ok || !response.body) throw new TransportError('transport_http', response.status)
  const pending = new Map<number, { id: string; name: string; json: string }>()
  let stopReason: string | undefined
  for await (const line of boundedSseLines(response.body)) {
    if (!line.startsWith('data:')) continue
    const payload = line.slice(5).trim()
    if (!payload || payload === '[DONE]') continue
    let event: {
      type?: string
      index?: number
      content_block?: { type?: string; id?: string; name?: string }
      delta?: { type?: string; text?: string; partial_json?: string; stop_reason?: string }
      error?: unknown
    }
    try {
      event = JSON.parse(payload) as typeof event
    } catch {
      throw new TransportError('transport_invalid_stream')
    }
    if (event.type === 'error' || event.error) throw new TransportError('transport_stream_error')
    const index = event.index ?? 0
    if (event.type === 'content_block_start' && event.content_block?.type === 'tool_use') {
      if (!pending.has(index) && pending.size >= MAX_PENDING_TOOL_CALLS) {
        throw new TransportError('transport_stream_too_large')
      }
      pending.set(index, {
        id: event.content_block.id ?? crypto.randomUUID(),
        name: event.content_block.name ?? '',
        json: '',
      })
    } else if (event.type === 'content_block_delta' && event.delta?.type === 'text_delta') {
      if (event.delta.text) callbacks.onDelta(event.delta.text)
    } else if (event.type === 'content_block_delta' && event.delta?.type === 'input_json_delta') {
      const tool = pending.get(index)
      const fragment = event.delta.partial_json ?? ''
      if (tool) {
        if (tool.json.length + fragment.length > MAX_STREAM_TOOL_INPUT_LENGTH) {
          throw new TransportError('transport_tool_input_too_large')
        }
        tool.json += fragment
      }
    } else if (event.type === 'content_block_stop') {
      const tool = pending.get(index)
      if (tool) {
        pending.delete(index)
        let call: AgentToolCall
        try {
          const input: unknown = tool.json.trim() ? JSON.parse(tool.json) : {}
          if (!input || typeof input !== 'object' || Array.isArray(input)) throw new Error()
          call = { id: tool.id, name: tool.name, input: input as Record<string, unknown> }
        } catch {
          call = { id: tool.id, name: tool.name, input: {}, inputError: 'invalid_tool_input' }
        }
        callbacks.onToolCall(call)
      }
    } else if (event.type === 'message_delta' && event.delta?.stop_reason) {
      stopReason = event.delta.stop_reason
    }
  }
  if (pending.size) throw new TransportError('transport_invalid_stream')
  if (stopReason) callbacks.onStopReason?.(stopReason)
}

export function createOfficeAgentTransport(
  config: RuntimeConfig,
  auth: BrowserAuth,
): AgentTransport {
  if (!trustedMessagesUrl(config.messagesUrl)) throw new Error('transport_unavailable')

  return {
    stream(request: AgentStreamRequest, callbacks: AgentStreamCallbacks) {
      const controller = new AbortController()
      let completed = false
      const done = () => {
        if (completed) return
        completed = true
        callbacks.onDone()
      }
      void (async () => {
        try {
          const body = JSON.stringify({
            model: WISWORK_DEFAULT_MODEL,
            max_tokens: MAX_TOKENS,
            system: request.system,
            messages: messagesForProvider(request.messages),
            tools: request.tools.map((tool) => ({
              name: tool.name,
              description: tool.description,
              input_schema: tool.inputSchema,
            })),
            stream: true,
          })
          if (body.length > MAX_REQUEST_BODY_LENGTH) {
            throw new TransportError('transport_request_too_large')
          }
          const response = await auth.authenticatedFetch(WISWORK_MESSAGES_URL, {
            method: 'POST',
            signal: controller.signal,
            headers: {
              'content-type': 'application/json',
              'x-req-location': WISWORK_REQUEST_LOCATION,
            },
            body,
          })
          await consumeStream(response, callbacks)
        } catch (error) {
          if (!controller.signal.aborted) {
            callbacks.onError(
              error instanceof TransportError ? error.publicMessage() : safeError(error),
            )
          }
        } finally {
          done()
        }
      })()
      return { cancel: () => controller.abort() }
    },
  }
}
