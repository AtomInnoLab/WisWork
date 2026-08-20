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
  sseLines,
} from '@wiswork/ai-provider'
import type { BrowserAuth } from '../auth/browser-auth.js'
import type { RuntimeConfig } from '../config.js'

const MAX_TOKENS = 8192

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

async function consumeStream(response: Response, callbacks: AgentStreamCallbacks): Promise<void> {
  if (!response.ok || !response.body) throw new Error(`transport_http_${response.status}`)
  const pending = new Map<number, { id: string; name: string; json: string }>()
  let stopReason: string | undefined
  for await (const line of sseLines(response.body)) {
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
      throw new Error('transport_invalid_stream')
    }
    if (event.type === 'error' || event.error) throw new Error('transport_stream_error')
    const index = event.index ?? 0
    if (event.type === 'content_block_start' && event.content_block?.type === 'tool_use') {
      pending.set(index, {
        id: event.content_block.id ?? crypto.randomUUID(),
        name: event.content_block.name ?? '',
        json: '',
      })
    } else if (event.type === 'content_block_delta' && event.delta?.type === 'text_delta') {
      if (event.delta.text) callbacks.onDelta(event.delta.text)
    } else if (event.type === 'content_block_delta' && event.delta?.type === 'input_json_delta') {
      const tool = pending.get(index)
      if (tool) tool.json += event.delta.partial_json ?? ''
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
  if (pending.size) throw new Error('transport_invalid_stream')
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
          const response = await auth.authenticatedFetch(WISWORK_MESSAGES_URL, {
            method: 'POST',
            signal: controller.signal,
            headers: {
              'content-type': 'application/json',
              'x-req-location': WISWORK_REQUEST_LOCATION,
            },
            body: JSON.stringify({
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
            }),
          })
          await consumeStream(response, callbacks)
        } catch (error) {
          if (!controller.signal.aborted) {
            const message = error instanceof Error ? error.message : ''
            callbacks.onError(message.startsWith('transport_') ? message : safeError(error))
          }
        } finally {
          done()
        }
      })()
      return { cancel: () => controller.abort() }
    },
  }
}
