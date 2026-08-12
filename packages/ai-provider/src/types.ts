import type { AgentMessage, AgentToolCall, AgentToolDef } from '@wiswork/agent-core'

export type AiProviderId = 'wiswork' | 'anthropic' | 'gemini' | 'deepseek' | 'openai' | 'custom'

export interface AiProviderConfig {
  apiKey: string
  model: string
  /** only used by the custom (OpenAI-compatible) provider */
  baseUrl?: string | undefined
}

/** Main-process-only authenticated request boundary used by the managed provider. */
export type WisworkFetchWithAuth = (
  request: (accessToken: string) => Promise<Response>,
) => Promise<Response>

export interface AiProviderMeta {
  id: AiProviderId
  label: string
  models: string[]
  defaultModel: string
  keyPlaceholder: string
  needsBaseUrl?: boolean
}

export interface AiSettings {
  provider: AiProviderId
  providers: Record<AiProviderId, AiProviderConfig>
}

/** pre-provider settings shape (single OpenAI-compatible endpoint); migrated into "custom" */
export interface LegacyAiSettings {
  baseUrl?: string
  apiKey?: string
  model?: string
}

export interface AiChatRequest {
  settings: AiSettings
  system: string
  user: string
}

export interface AiChatResponse {
  ok: boolean
  content?: string
  error?: string
  errorCode?:
    | 'auth_required'
    | 'model_credentials_missing'
    | 'model_rate_limited'
    | 'model_upstream_unavailable'
    | 'model_invalid_response'
}

export interface AiStreamRequest {
  requestId: string
  settings: AiSettings
  system: string
  messages: AgentMessage[]
  tools?: AgentToolDef[]
  maxTokens?: number
}

export interface AiStreamChunk {
  requestId: string
  /** 'ping' = wire-level keepalive so the renderer can tell a live stream from a dead one */
  type: 'delta' | 'tool-call' | 'done' | 'error' | 'ping'
  text?: string
  /** complete parsed tool call (emitted once its arguments finish streaming) */
  toolCall?: AgentToolCall
  error?: string
  /** machine-readable error cause ('timeout', exhausted 'credits'); lets the renderer localize the message */
  errorCode?:
    | 'timeout'
    | 'credits'
    | 'auth_required'
    | 'model_credentials_missing'
    | 'model_rate_limited'
    | 'model_upstream_unavailable'
    | 'model_invalid_response'
  /** normalized stop reason carried on 'done' ('max_tokens' = output cut off by the token limit) */
  stopReason?: string
}
