import type {
  AgentMessage,
  AgentStreamCallbacks,
  AgentStreamRequest,
  AgentToolCall,
  AgentTransport,
  ToolDisplay,
} from '@wiswork/agent-core'
import { WISWORK_DEFAULT_MODEL } from '@wiswork/ai-provider'

const MAX_TOKENS = 8192
export const MAX_STREAM_TOOL_INPUT_LENGTH = 16 * 1024
export const MAX_REQUEST_BODY_LENGTH = 256 * 1024
export const MAX_STREAM_TEXT_LENGTH = 128 * 1024
export const MAX_COMPLETED_TOOL_CALLS = 32
// Enhanced streams span a whole document session, not a single provider response.
export const MAX_OBSERVED_TOOL_CALLS = 1024
// Standard responses keep their absolute budget. Enhanced responses contain an entire
// tool loop: renew on actual progress, bounded by a separate whole-turn deadline.
export const STREAM_RESPONSE_TIMEOUT_MS = 280_000
export const ENHANCED_TURN_TIMEOUT_MS = 30 * 60_000
const MAX_STREAM_RESPONSE_BYTES = 1024 * 1024
const MAX_STREAM_EVENTS = 4096
const MAX_ENHANCED_STREAM_RESPONSE_BYTES = 8 * 1024 * 1024
const MAX_ENHANCED_STREAM_EVENTS = 32 * 1024
const ENHANCED_STREAM_BUDGET_WINDOW_MS = 60_000
const MAX_SSE_LINE_LENGTH = 64 * 1024
const MAX_PENDING_TOOL_CALLS = 16

export interface OfficeToolActivity {
  readonly callId: string
  readonly toolName: string
  readonly state: 'running' | 'complete' | 'error'
  readonly startedAt: number
  readonly query?: string
  readonly summary?: string
  readonly resultCount?: number
  readonly display?: ToolDisplay
}
export interface OfficeAgentTransport extends AgentTransport {
  setToolActivityHandler?(handler: ((event: OfficeToolActivity) => void) | undefined): void
}

function parseToolActivity(
  event: Record<string, unknown>,
  generation: number | undefined,
): OfficeToolActivity {
  const keys = [
    'type',
    'generation',
    'call_id',
    'tool_name',
    'state',
    'started_at',
    'query',
    'summary',
    'result_count',
    'display',
  ]
  const invalid = () => {
    throw new TransportError('transport_invalid_stream')
  }
  if (
    new TextEncoder().encode(JSON.stringify(event)).byteLength > 12 * 1024 ||
    Object.keys(event).some((key) => !keys.includes(key)) ||
    generation === undefined ||
    event.generation !== generation ||
    typeof event.call_id !== 'string' ||
    !/^call_[A-Za-z0-9_-]{8,123}$/.test(event.call_id) ||
    typeof event.tool_name !== 'string' ||
    !/^[A-Za-z0-9_-]{1,128}$/.test(event.tool_name) ||
    !['running', 'complete', 'error'].includes(String(event.state)) ||
    !Number.isSafeInteger(event.started_at) ||
    Number(event.started_at) < 0 ||
    (event.query !== undefined && (typeof event.query !== 'string' || event.query.length > 240)) ||
    (event.summary !== undefined &&
      (typeof event.summary !== 'string' || event.summary.length > 160)) ||
    (event.result_count !== undefined &&
      (!Number.isSafeInteger(event.result_count) ||
        Number(event.result_count) < 0 ||
        Number(event.result_count) > 20))
  )
    invalid()
  let display: ToolDisplay | undefined
  if (event.display !== undefined) {
    const value = event.display as ToolDisplay
    if (
      !value ||
      typeof value !== 'object' ||
      !['images', 'links'].includes(value.kind) ||
      Object.keys(value).some((key) => !['kind', 'items'].includes(key)) ||
      !Array.isArray(value.items) ||
      value.items.length > 8
    )
      invalid()
    const items = value.items!.map((item) => {
      if (
        !item ||
        typeof item !== 'object' ||
        Object.keys(item).some((key) => !['url', 'title'].includes(key)) ||
        typeof item.url !== 'string' ||
        item.url.length > 2048 ||
        (item.title !== undefined && (typeof item.title !== 'string' || item.title.length > 160))
      )
        invalid()
      let url: URL
      try {
        url = new URL(item.url)
      } catch {
        return invalid()
      }
      if (
        url.protocol !== 'https:' ||
        url.username ||
        url.password ||
        url.port ||
        url.search ||
        url.hash ||
        !url.hostname.includes('.') ||
        /^[\d.]+$/.test(url.hostname) ||
        url.hostname.includes(':') ||
        /\.(?:localhost|local)$/.test(url.hostname)
      )
        invalid()
      return { url: url.href, ...(item.title === undefined ? {} : { title: item.title }) }
    })
    display = { kind: value.kind, items }
  }
  return {
    callId: event.call_id as string,
    toolName: event.tool_name as OfficeToolActivity['toolName'],
    state: event.state as OfficeToolActivity['state'],
    startedAt: Number(event.started_at),
    ...(event.query === undefined ? {} : { query: event.query as string }),
    ...(event.summary === undefined ? {} : { summary: event.summary as string }),
    ...(event.result_count === undefined ? {} : { resultCount: Number(event.result_count) }),
    ...(display ? { display } : {}),
  }
}

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
  if (error instanceof Error) {
    const known: Readonly<Record<string, string>> = {
      relay_session_expired: 'session_expired',
      relay_auth_required: 'auth_required',
      relay_request_timeout: 'request_timeout',
      relay_timeout: 'request_timeout',
      relay_upstream_error: 'provider_unavailable',
      relay_disconnected: 'network_error',
    }
    if (Object.hasOwn(known, error.message)) return known[error.message]!
  }
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
  isEnhanced: () => boolean,
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
  let responseWindowStartedAt = Date.now()
  let buffer = ''
  try {
    while (true) {
      const { done, value } = await reader.read()
      if (signal.aborted) throw new DOMException('Aborted', 'AbortError')
      if (done) {
        finished = true
        break
      }
      const enhanced = isEnhanced()
      if (enhanced && Date.now() - responseWindowStartedAt >= ENHANCED_STREAM_BUDGET_WINDOW_MS) {
        responseBytes = 0
        responseWindowStartedAt = Date.now()
      }
      responseBytes += value.byteLength
      const maxResponseBytes = isEnhanced()
        ? MAX_ENHANCED_STREAM_RESPONSE_BYTES
        : MAX_STREAM_RESPONSE_BYTES
      if (responseBytes > maxResponseBytes) {
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
  handleControl?: (event: Record<string, unknown>, signal: AbortSignal) => Promise<void>,
  handleActivity?: (event: Record<string, unknown>) => void,
  onProgress?: () => void,
  isEnhanced: () => boolean = () => false,
): Promise<void> {
  if (!response.ok) {
    await response.body?.cancel().catch(() => undefined)
    throw new TransportError('transport_http', response.status)
  }
  if (!response.body) throw new TransportError('transport_http', response.status)
  const pending = new Map<number, { id: string; name: string; json: string }>()
  let stopReason: string | undefined
  let eventCount = 0
  let eventWindowStartedAt = Date.now()
  let textLength = 0
  let completedToolCalls = 0
  for await (const line of boundedSseLines(response.body, signal, isEnhanced)) {
    if (!line.startsWith('data:')) continue
    const payload = line.slice(5).trim()
    if (!payload || payload === '[DONE]') continue
    const enhanced = isEnhanced()
    if (enhanced && Date.now() - eventWindowStartedAt >= ENHANCED_STREAM_BUDGET_WINDOW_MS) {
      eventCount = 0
      eventWindowStartedAt = Date.now()
    }
    eventCount += 1
    const maxEvents = enhanced ? MAX_ENHANCED_STREAM_EVENTS : MAX_STREAM_EVENTS
    if (eventCount > maxEvents) {
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
    if (event.type === 'wiswork_tool_activity' || event.type === 'wiswork_tool_lifecycle') {
      handleActivity?.(event as Record<string, unknown>)
      onProgress?.()
      continue
    }
    if (event.type === 'wiswork_tool_call') {
      if (!handleControl) throw new TransportError('transport_invalid_stream')
      onProgress?.()
      await handleControl(event as Record<string, unknown>, signal)
      onProgress?.()
      continue
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
        onProgress?.()
      }
    } else if (event.type === 'content_block_delta' && event.delta?.type === 'input_json_delta') {
      const tool = pending.get(index)
      const fragment = event.delta.partial_json ?? ''
      if (tool) {
        if (tool.json.length + fragment.length > MAX_STREAM_TOOL_INPUT_LENGTH) {
          throw new TransportError('transport_tool_input_too_large')
        }
        tool.json += fragment
        if (fragment) onProgress?.()
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
  handleToolFrame?(event: Record<string, unknown>, signal: AbortSignal): Promise<void>
  snapshot?(): { enhanced?: { session_generation: number } }
}): OfficeAgentTransport {
  return createTransport(
    (init) => bridge.authenticatedFetch('/v1/office/messages', init),
    bridge.handleToolFrame?.bind(bridge),
    () => bridge.snapshot?.().enhanced?.session_generation,
  )
}

function createTransport(
  fetchMessages: (init: RequestInit) => Promise<Response>,
  handleControl?: (event: Record<string, unknown>, signal: AbortSignal) => Promise<void>,
  enhancedGeneration: () => number | undefined = () => undefined,
): OfficeAgentTransport {
  let activityHandler: ((event: OfficeToolActivity) => void) | undefined
  let activityEpoch = 0
  return {
    setToolActivityHandler(handler) {
      activityHandler = handler
      activityEpoch += 1
    },
    stream(request: AgentStreamRequest, callbacks: AgentStreamCallbacks) {
      const controller = new AbortController()
      const allowed = new Set(request.tools.map((tool) => tool.name))
      // Tool-free review subturns must not observe or supersede the user-facing stream.
      const epoch = allowed.size ? ++activityEpoch : undefined
      let generation = enhancedGeneration()
      const pendingActivities = new Map<string, OfficeToolActivity>()
      const completedActivities = new Set<string>()
      const observe = (event: OfficeToolActivity) => {
        if (epoch === activityEpoch) activityHandler?.(event)
      }
      const closeActivities = () => {
        for (const activity of pendingActivities.values())
          observe({ ...activity, state: 'error', summary: 'Tool interrupted' })
        pendingActivities.clear()
      }
      const handleActivity = (event: Record<string, unknown>) => {
        if (epoch === undefined || controller.signal.aborted || epoch !== activityEpoch) return
        // The first request may promote a standard Relay session before its first SSE event.
        // Bind that newly negotiated generation once; never accept a later replacement.
        const negotiatedGeneration = enhancedGeneration()
        generation ??= negotiatedGeneration
        const activity = parseToolActivity(event, generation)
        if (negotiatedGeneration !== generation || !allowed.has(activity.toolName))
          throw new TransportError('transport_invalid_stream')
        const pending = pendingActivities.get(activity.callId)
        if (
          completedActivities.has(activity.callId) ||
          (activity.state === 'running'
            ? pending !== undefined
            : !pending ||
              pending.toolName !== activity.toolName ||
              pending.startedAt !== activity.startedAt) ||
          (activity.state === 'running' &&
            pendingActivities.size + completedActivities.size >= MAX_OBSERVED_TOOL_CALLS)
        )
          throw new TransportError('transport_invalid_stream')
        if (activity.state === 'running') pendingActivities.set(activity.callId, activity)
        else {
          pendingActivities.delete(activity.callId)
          completedActivities.add(activity.callId)
        }
        observe(activity)
      }
      let timeout: ReturnType<typeof setTimeout> | undefined
      const startedAt = Date.now()
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
          let renewDeadline = () => {}
          const expired = new Promise<never>((_resolve, reject) => {
            const expire = () => {
              reject(new TransportError('transport_timeout'))
              controller.abort()
            }
            timeout = setTimeout(expire, STREAM_RESPONSE_TIMEOUT_MS)
            renewDeadline = () => {
              const negotiated = enhancedGeneration()
              generation ??= negotiated
              if (
                controller.signal.aborted ||
                completed ||
                generation === undefined ||
                negotiated !== generation
              )
                return
              clearTimeout(timeout)
              timeout = setTimeout(
                expire,
                Math.max(
                  0,
                  Math.min(
                    STREAM_RESPONSE_TIMEOUT_MS,
                    ENHANCED_TURN_TIMEOUT_MS - (Date.now() - startedAt),
                  ),
                ),
              )
            }
          })
          const operation = fetchMessages({
            method: 'POST',
            signal: controller.signal,
            headers: { 'content-type': 'application/json' },
            body,
          }).then((response) =>
            consumeStream(
              response,
              callbacks,
              controller.signal,
              handleControl,
              handleActivity,
              renewDeadline,
              () => enhancedGeneration() !== undefined,
            ),
          )
          const cancelled = new Promise<never>((_resolve, reject) => {
            cancelListener = () => reject(new DOMException('Aborted', 'AbortError'))
            controller.signal.addEventListener('abort', cancelListener, { once: true })
          })
          await Promise.race([operation, expired, cancelled])
        } catch (error) {
          // Settle visible cards before onError makes the session inactive.
          closeActivities()
          if (error instanceof TransportError && error.code === 'transport_timeout')
            callbacks.onError(error.publicMessage())
          else if (!controller.signal.aborted)
            callbacks.onError(
              error instanceof TransportError ? error.publicMessage() : safeError(error),
            )
        } finally {
          if (timeout !== undefined) clearTimeout(timeout)
          if (cancelListener) controller.signal.removeEventListener('abort', cancelListener)
          closeActivities()
          done()
        }
      })()
      return { cancel: () => controller.abort() }
    },
  }
}
