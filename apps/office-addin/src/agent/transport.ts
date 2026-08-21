import type {
  AgentMessage,
  AgentStreamCallbacks,
  AgentStreamRequest,
  AgentToolCall,
  AgentTransport,
} from '@wiswork/agent-core'
import { WISWORK_DEFAULT_MODEL } from '@wiswork/ai-provider'

const MAX_TOKENS = 8192
export const MAX_STREAM_TOOL_INPUT_LENGTH = 16 * 1024
export const MAX_REQUEST_BODY_LENGTH = 256 * 1024
export const MAX_STREAM_TEXT_LENGTH = 128 * 1024
export const MAX_COMPLETED_TOOL_CALLS = 32
export const STREAM_RESPONSE_TIMEOUT_MS = 60_000
const MAX_STREAM_RESPONSE_BYTES = 1024 * 1024
const MAX_STREAM_EVENTS = 4096
const MAX_SSE_LINE_LENGTH = 64 * 1024
const MAX_PENDING_TOOL_CALLS = 16

type TransportErrorCode =
  | 'transport_http'
  | 'transport_invalid_stream'
  | 'transport_stream_error'
  | 'transport_stream_too_large'
  | 'transport_tool_input_too_large'
  | 'transport_request_too_large'
  | 'transport_stream_budget_exceeded'
  | 'transport_timeout'

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

async function* boundedSseLines(
  body: ReadableStream<Uint8Array>,
  signal: AbortSignal,
): AsyncGenerator<string> {
  const decoder = new TextDecoder()
  const reader = body.getReader()
  let readerCancelled = false
  let cancellation: Promise<void> | undefined
  let finished = false
  const abort = () => {
    readerCancelled = true
    cancellation = reader.cancel().catch(() => undefined)
  }
  signal.addEventListener('abort', abort, { once: true })
  let responseBytes = 0
  let buffer = ''
  try {
    while (true) {
      const { done, value } = await reader.read()
      if (signal.aborted) throw new DOMException('Aborted', 'AbortError')
      if (done) {
        finished = true
        break
      }
      responseBytes += value.byteLength
      if (responseBytes > MAX_STREAM_RESPONSE_BYTES) {
        throw new TransportError('transport_stream_budget_exceeded')
      }
      buffer += decoder.decode(value, { stream: true })
      let newline = buffer.indexOf('\n')
      while (newline !== -1) {
        const line = buffer.slice(0, newline).replace(/\r$/, '')
        if (line.length > MAX_SSE_LINE_LENGTH) {
          throw new TransportError('transport_stream_too_large')
        }
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
  } finally {
    signal.removeEventListener('abort', abort)
    if (cancellation) await cancellation
    else if (!finished && !readerCancelled) await reader.cancel().catch(() => undefined)
    reader.releaseLock()
  }
}

async function consumeStream(
  response: Response,
  callbacks: AgentStreamCallbacks,
  signal: AbortSignal,
): Promise<void> {
  if (!response.ok) {
    await response.body?.cancel().catch(() => undefined)
    throw new TransportError('transport_http', response.status)
  }
  if (!response.body) throw new TransportError('transport_http', response.status)
  const pending = new Map<number, { id: string; name: string; json: string }>()
  let stopReason: string | undefined
  let eventCount = 0
  let textLength = 0
  let completedToolCalls = 0
  for await (const line of boundedSseLines(response.body, signal)) {
    if (!line.startsWith('data:')) continue
    const payload = line.slice(5).trim()
    if (!payload || payload === '[DONE]') continue
    eventCount += 1
    if (eventCount > MAX_STREAM_EVENTS) {
      throw new TransportError('transport_stream_budget_exceeded')
    }
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
      if (event.delta.text) {
        textLength += event.delta.text.length
        if (textLength > MAX_STREAM_TEXT_LENGTH) {
          throw new TransportError('transport_stream_budget_exceeded')
        }
        callbacks.onDelta(event.delta.text)
      }
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
        completedToolCalls += 1
        if (completedToolCalls > MAX_COMPLETED_TOOL_CALLS) {
          throw new TransportError('transport_stream_budget_exceeded')
        }
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

/** PC-backed transport: provider credentials remain in WisWork PC. */
export function createPcBridgeAgentTransport(bridge: {
  authenticatedFetch(path: '/v1/office/messages', init: RequestInit): Promise<Response>
}): AgentTransport {
  return createTransport((init) => bridge.authenticatedFetch('/v1/office/messages', init))
}

function createTransport(fetchMessages: (init: RequestInit) => Promise<Response>): AgentTransport {
  return {
    stream(request: AgentStreamRequest, callbacks: AgentStreamCallbacks) {
      const controller = new AbortController()
      let timeout: ReturnType<typeof setTimeout> | undefined
      let cancelListener: (() => void) | undefined
      let completed = false
      const done = () => {
        if (!completed) {
          completed = true
          callbacks.onDone()
        }
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
          if (body.length > MAX_REQUEST_BODY_LENGTH)
            throw new TransportError('transport_request_too_large')
          const operation = fetchMessages({
            method: 'POST',
            signal: controller.signal,
            headers: { 'content-type': 'application/json' },
            body,
          }).then((response) => consumeStream(response, callbacks, controller.signal))
          const expired = new Promise<never>((_resolve, reject) => {
            timeout = setTimeout(() => {
              reject(new TransportError('transport_timeout'))
              controller.abort()
            }, STREAM_RESPONSE_TIMEOUT_MS)
          })
          const cancelled = new Promise<never>((_resolve, reject) => {
            cancelListener = () => reject(new DOMException('Aborted', 'AbortError'))
            controller.signal.addEventListener('abort', cancelListener, { once: true })
          })
          await Promise.race([operation, expired, cancelled])
        } catch (error) {
          if (error instanceof TransportError && error.code === 'transport_timeout')
            callbacks.onError(error.publicMessage())
          else if (!controller.signal.aborted)
            callbacks.onError(
              error instanceof TransportError ? error.publicMessage() : safeError(error),
            )
        } finally {
          if (timeout !== undefined) clearTimeout(timeout)
          if (cancelListener) controller.signal.removeEventListener('abort', cancelListener)
          done()
        }
      })()
      return { cancel: () => controller.abort() }
    },
  }
}
