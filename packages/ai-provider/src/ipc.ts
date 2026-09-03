import { AiCreditsError, streamForProvider } from './stream'
import { chatForProvider } from './chat'
import { AiProviderError, isAuthRequiredError } from './errors'
import { resolveWisworkMainRequest, sanitizeWisworkSettings } from './main-config'
import { defaultAiSettings } from './providers'
import { AiTimeoutError } from './watchdog'
import type {
  AiChatRequest,
  AiProviderConfig,
  AiProviderId,
  AiSettings,
  AiServiceDiagnostic,
  AiStreamChunk,
  AiStreamRequest,
  WisworkFetchWithAuth,
} from './types'

export const AI_IPC_LIMITS = {
  maxTokens: 32_768,
  maxMessages: 500,
  maxTools: 50,
  maxImages: 20,
  maxSystemChars: 1_000_000,
  maxMessageChars: 1_000_000,
  maxImageChars: 8_000_000,
  maxToolDescriptionChars: 32_000,
  maxTotalChars: 16_000_000,
  maxJsonDepth: 20,
  maxJsonNodes: 20_000,
} as const

export type AiIpcErrorCode =
  | 'untrusted_sender'
  | 'invalid_payload'
  | 'payload_too_large'
  | 'invalid_request_id'
  | 'duplicate_request_id'
  | 'request_owner_mismatch'

export class AiIpcError extends Error {
  readonly code: AiIpcErrorCode

  constructor(code: AiIpcErrorCode) {
    super(code)
    this.name = 'AiIpcError'
    this.code = code
  }
}

export interface WisworkIpcSender {
  id: number
  isDestroyed(): boolean
  send(channel: string, value: unknown): void
}

export interface WisworkIpcEvent {
  sender: WisworkIpcSender
}

export interface IpcMainLike {
  handle(channel: string, handler: (event: WisworkIpcEvent, ...args: unknown[]) => unknown): void
}

export interface WisworkModelIpcChannels {
  getSettings: string
  setSettings: string
  stream: string
  streamChunk: string
  cancel: string
  chat?: string
}

export interface RegisterWisworkModelIpcOptions {
  ipcMain: IpcMainLike
  channels: WisworkModelIpcChannels
  isTrustedSender(senderId: number): boolean
  loadSettings(): unknown
  saveSettings(settings: AiSettings): void
  getAccessToken(): Promise<string | null>
  fetchWithAuth: WisworkFetchWithAuth
}

const PROVIDER_IDS: ReadonlySet<string> = new Set<AiProviderId>([
  'wiswork',
  'anthropic',
  'gemini',
  'deepseek',
  'openai',
  'custom',
])
const REQUEST_ID = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i

function record(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new AiIpcError('invalid_payload')
  }
  const prototype = Object.getPrototypeOf(value)
  if (prototype !== Object.prototype && prototype !== null) {
    throw new AiIpcError('invalid_payload')
  }
  return value as Record<string, unknown>
}

function exact(value: Record<string, unknown>, allowed: readonly string[]): void {
  if (Object.keys(value).some((key) => !allowed.includes(key))) {
    throw new AiIpcError('invalid_payload')
  }
}

function boundedString(value: unknown, max: number, state?: { chars: number }): string {
  if (typeof value !== 'string') throw new AiIpcError('invalid_payload')
  if (value.length > max) throw new AiIpcError('payload_too_large')
  if (state) {
    state.chars += value.length
    if (state.chars > AI_IPC_LIMITS.maxTotalChars) throw new AiIpcError('payload_too_large')
  }
  return value
}

function optionalString(value: unknown, max: number): string | undefined {
  return value === undefined ? undefined : boundedString(value, max)
}

function validateRequestId(value: unknown): string {
  const id = boundedString(value, 64)
  if (!REQUEST_ID.test(id)) throw new AiIpcError('invalid_request_id')
  return id
}

function measureJson(value: unknown, state: { chars: number; nodes: number }, depth = 0): void {
  if (depth > AI_IPC_LIMITS.maxJsonDepth) throw new AiIpcError('payload_too_large')
  state.nodes += 1
  if (state.nodes > AI_IPC_LIMITS.maxJsonNodes) throw new AiIpcError('payload_too_large')
  if (value === null || typeof value === 'boolean' || typeof value === 'number') return
  if (typeof value === 'string') {
    boundedString(value, AI_IPC_LIMITS.maxMessageChars, state)
    return
  }
  if (Array.isArray(value)) {
    for (const item of value) measureJson(item, state, depth + 1)
    return
  }
  const object = record(value)
  for (const [key, item] of Object.entries(object)) {
    boundedString(key, 256, state)
    measureJson(item, state, depth + 1)
  }
}

export function validateAiSettings(value: unknown): AiSettings {
  const settings = record(value)
  exact(settings, ['provider', 'providers'])
  if (typeof settings.provider !== 'string' || !PROVIDER_IDS.has(settings.provider)) {
    throw new AiIpcError('invalid_payload')
  }
  const providers = record(settings.providers)
  if (!('wiswork' in providers) || Object.keys(providers).length > PROVIDER_IDS.size) {
    throw new AiIpcError('invalid_payload')
  }
  const validated = {} as AiSettings['providers']
  for (const [id, rawConfig] of Object.entries(providers)) {
    if (!PROVIDER_IDS.has(id)) throw new AiIpcError('invalid_payload')
    const config = record(rawConfig)
    exact(config, ['apiKey', 'model', 'baseUrl'])
    validated[id as AiProviderId] = {
      apiKey: boundedString(config.apiKey, 16_384),
      model: boundedString(config.model, 256),
      baseUrl: optionalString(config.baseUrl, 2_048),
    } as AiProviderConfig
  }
  return { provider: settings.provider as AiProviderId, providers: validated }
}

function validateMessages(value: unknown, state: { chars: number; nodes: number }) {
  if (!Array.isArray(value) || value.length > AI_IPC_LIMITS.maxMessages) {
    throw new AiIpcError(value instanceof Array ? 'payload_too_large' : 'invalid_payload')
  }
  let imageCount = 0
  for (const raw of value) {
    const message = record(raw)
    if (message.role === 'user') {
      exact(message, ['role', 'text', 'images'])
      boundedString(message.text, AI_IPC_LIMITS.maxMessageChars, state)
      if (message.images !== undefined) {
        if (!Array.isArray(message.images)) throw new AiIpcError('invalid_payload')
        imageCount += message.images.length
        if (imageCount > AI_IPC_LIMITS.maxImages) throw new AiIpcError('payload_too_large')
        for (const rawImage of message.images) {
          const image = record(rawImage)
          exact(image, ['base64', 'mime'])
          boundedString(image.base64, AI_IPC_LIMITS.maxImageChars, state)
          boundedString(image.mime, 64, state)
        }
      }
    } else if (message.role === 'assistant') {
      exact(message, ['role', 'text', 'toolCalls'])
      boundedString(message.text, AI_IPC_LIMITS.maxMessageChars, state)
      if (message.toolCalls !== undefined) {
        if (
          !Array.isArray(message.toolCalls) ||
          message.toolCalls.length > AI_IPC_LIMITS.maxTools
        ) {
          throw new AiIpcError('invalid_payload')
        }
        for (const rawCall of message.toolCalls) {
          const call = record(rawCall)
          exact(call, ['id', 'name', 'input', 'inputError', 'truncated'])
          boundedString(call.id, 256, state)
          boundedString(call.name, 256, state)
          measureJson(call.input, state)
        }
      }
    } else if (message.role === 'tool') {
      exact(message, ['role', 'results'])
      if (!Array.isArray(message.results) || message.results.length > AI_IPC_LIMITS.maxTools) {
        throw new AiIpcError('invalid_payload')
      }
      for (const rawResult of message.results) {
        const result = record(rawResult)
        exact(result, ['id', 'name', 'output', 'isError', 'content'])
        boundedString(result.id, 256, state)
        boundedString(result.name, 256, state)
        boundedString(result.output, AI_IPC_LIMITS.maxMessageChars, state)
        if (result.isError !== undefined && typeof result.isError !== 'boolean') {
          throw new AiIpcError('invalid_payload')
        }
        if (result.content !== undefined) {
          if (!Array.isArray(result.content) || result.content.length > 4)
            throw new AiIpcError('payload_too_large')
          for (const rawBlock of result.content) {
            const block = record(rawBlock)
            exact(block, ['type', 'image'])
            if (block.type !== 'image') throw new AiIpcError('invalid_payload')
            const image = record(block.image)
            exact(image, ['base64', 'mime'])
            boundedString(image.base64, AI_IPC_LIMITS.maxImageChars, state)
            boundedString(image.mime, 64, state)
            if (!['image/png', 'image/jpeg', 'image/webp'].includes(image.mime as string))
              throw new AiIpcError('invalid_payload')
          }
        }
      }
    } else {
      throw new AiIpcError('invalid_payload')
    }
  }
  return value as AiStreamRequest['messages']
}

function validateTools(value: unknown, state: { chars: number; nodes: number }) {
  if (value === undefined) return []
  if (!Array.isArray(value) || value.length > AI_IPC_LIMITS.maxTools) {
    throw new AiIpcError(value instanceof Array ? 'payload_too_large' : 'invalid_payload')
  }
  for (const raw of value) {
    const tool = record(raw)
    exact(tool, ['name', 'description', 'inputSchema'])
    boundedString(tool.name, 256, state)
    boundedString(tool.description, AI_IPC_LIMITS.maxToolDescriptionChars, state)
    measureJson(tool.inputSchema, state)
  }
  return value as NonNullable<AiStreamRequest['tools']>
}

export function validateAiStreamRequest(value: unknown): AiStreamRequest {
  const request = record(value)
  exact(request, ['requestId', 'settings', 'system', 'messages', 'tools', 'maxTokens'])
  const state = { chars: 0, nodes: 0 }
  const maxTokens = request.maxTokens ?? AI_IPC_LIMITS.maxTokens
  if (
    typeof maxTokens !== 'number' ||
    !Number.isInteger(maxTokens) ||
    maxTokens < 1 ||
    maxTokens > AI_IPC_LIMITS.maxTokens
  ) {
    throw new AiIpcError('invalid_payload')
  }
  return {
    requestId: validateRequestId(request.requestId),
    settings: validateAiSettings(request.settings),
    system: boundedString(request.system, AI_IPC_LIMITS.maxSystemChars, state),
    messages: validateMessages(request.messages, state),
    tools: validateTools(request.tools, state),
    maxTokens,
  }
}

export function validateAiSearchArgs(
  query: unknown,
  maxResults: unknown,
  fallback: number,
): { query: string; maxResults: number } {
  const validatedQuery = boundedString(query, 4_096)
  const validatedMax = maxResults === undefined ? fallback : maxResults
  if (
    typeof validatedMax !== 'number' ||
    !Number.isInteger(validatedMax) ||
    validatedMax < 1 ||
    validatedMax > 20
  ) {
    throw new AiIpcError('invalid_payload')
  }
  return { query: validatedQuery, maxResults: validatedMax }
}

export function validateAiChatRequest(value: unknown): AiChatRequest {
  const request = record(value)
  exact(request, ['settings', 'system', 'user'])
  const state = { chars: 0, nodes: 0 }
  return {
    settings: validateAiSettings(request.settings),
    system: boundedString(request.system, AI_IPC_LIMITS.maxSystemChars, state),
    user: boundedString(request.user, AI_IPC_LIMITS.maxMessageChars, state),
  }
}

function safeSettings(value: unknown): AiSettings {
  try {
    return sanitizeWisworkSettings(validateAiSettings(value))
  } catch {
    return sanitizeWisworkSettings(defaultAiSettings())
  }
}

type StableStreamErrorCode = NonNullable<AiStreamChunk['errorCode']>

function stableErrorCode(error: unknown): StableStreamErrorCode {
  if (isAuthRequiredError(error)) return 'auth_required'
  if (error instanceof AiTimeoutError) return 'timeout'
  if (error instanceof AiCreditsError) return 'credits'
  if (error instanceof AiProviderError) return error.code
  return 'model_upstream_unavailable'
}

function stableDiagnostic(error: unknown): AiServiceDiagnostic {
  if (error instanceof AiProviderError) {
    return {
      stage: error.stage ?? (error.status ? 'response' : 'stream'),
      ...(error.status ? { httpStatus: error.status } : {}),
    }
  }
  return { stage: isAuthRequiredError(error) ? 'auth' : 'stream' }
}

export function registerWisworkModelIpc(options: RegisterWisworkModelIpcOptions): void {
  const active = new Map<string, { senderId: number; controller: AbortController }>()

  const trusted = (event: WisworkIpcEvent) => {
    if (!options.isTrustedSender(event.sender.id)) throw new AiIpcError('untrusted_sender')
  }
  const settingsForRequest = async (settings: AiSettings) => {
    const accessToken = await options.getAccessToken()
    return resolveWisworkMainRequest(accessToken !== null, settings.providers.wiswork)
  }
  const send = (event: WisworkIpcEvent, chunk: AiStreamChunk) => {
    if (!event.sender.isDestroyed()) event.sender.send(options.channels.streamChunk, chunk)
  }

  options.ipcMain.handle(options.channels.getSettings, (event, ...args) => {
    trusted(event)
    if (args.length !== 0) throw new AiIpcError('invalid_payload')
    return safeSettings(options.loadSettings())
  })

  options.ipcMain.handle(options.channels.setSettings, (event, ...args) => {
    trusted(event)
    if (args.length !== 1) throw new AiIpcError('invalid_payload')
    const settings = sanitizeWisworkSettings(validateAiSettings(args[0]))
    options.saveSettings(settings)
  })

  options.ipcMain.handle(options.channels.stream, async (event, ...args) => {
    trusted(event)
    if (args.length !== 1) throw new AiIpcError('invalid_payload')
    const request = validateAiStreamRequest(args[0])
    if (active.has(request.requestId)) throw new AiIpcError('duplicate_request_id')
    const controller = new AbortController()
    active.set(request.requestId, { senderId: event.sender.id, controller })
    let lastPing = 0
    try {
      const resolved = await settingsForRequest(request.settings)
      if (!resolved.ok) {
        send(event, {
          requestId: request.requestId,
          type: 'error',
          error: resolved.errorCode,
          errorCode: resolved.errorCode,
        })
        return
      }
      let stopReason: string | undefined
      await streamForProvider(
        resolved.provider,
        resolved.config,
        request.system,
        request.messages,
        request.tools ?? [],
        request.maxTokens ?? AI_IPC_LIMITS.maxTokens,
        {
          signal: controller.signal,
          onDelta: (text) => send(event, { requestId: request.requestId, type: 'delta', text }),
          onToolCall: (toolCall) =>
            send(event, { requestId: request.requestId, type: 'tool-call', toolCall }),
          onActivity: () => {
            const now = Date.now()
            if (now - lastPing < 5_000) return
            lastPing = now
            send(event, { requestId: request.requestId, type: 'ping' })
          },
          onStopReason: (reason) => {
            stopReason = reason
          },
        },
        options.fetchWithAuth,
      )
      send(event, {
        requestId: request.requestId,
        type: 'done',
        ...(stopReason ? { stopReason } : {}),
      })
    } catch (error) {
      if (controller.signal.aborted) {
        send(event, { requestId: request.requestId, type: 'done' })
      } else {
        const code = stableErrorCode(error)
        send(event, {
          requestId: request.requestId,
          type: 'error',
          error: code,
          errorCode: code,
          diagnostic: stableDiagnostic(error),
        })
      }
    } finally {
      const current = active.get(request.requestId)
      if (current?.controller === controller) active.delete(request.requestId)
    }
  })

  options.ipcMain.handle(options.channels.cancel, (event, ...args) => {
    trusted(event)
    if (args.length !== 1) throw new AiIpcError('invalid_payload')
    const requestId = validateRequestId(args[0])
    const current = active.get(requestId)
    if (!current) return false
    if (current.senderId !== event.sender.id) throw new AiIpcError('request_owner_mismatch')
    current.controller.abort()
    return true
  })

  if (options.channels.chat) {
    options.ipcMain.handle(options.channels.chat, async (event, ...args) => {
      trusted(event)
      if (args.length !== 1) throw new AiIpcError('invalid_payload')
      const request = validateAiChatRequest(args[0])
      try {
        const resolved = await settingsForRequest(request.settings)
        if (!resolved.ok) {
          return { ok: false, error: resolved.errorCode, errorCode: resolved.errorCode }
        }
        const result = await chatForProvider(
          resolved.provider,
          resolved.config,
          request.system,
          request.user,
          undefined,
          options.fetchWithAuth,
        )
        if (result.ok) return result
        const code = result.errorCode ?? 'model_upstream_unavailable'
        return { ok: false, error: code, errorCode: code, diagnostic: result.diagnostic }
      } catch (error) {
        const code = stableErrorCode(error)
        return { ok: false, error: code, errorCode: code, diagnostic: stableDiagnostic(error) }
      }
    })
  }
}
