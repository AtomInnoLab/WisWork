import type { AgentRuntimeMode, EnhancedHost } from '@wiswork/agent-runtime'

export const CODEX_RUNTIME_CHANNELS = {
  status: 'codex:runtime:status',
  startTurn: 'codex:runtime:start-turn',
  cancelTurn: 'codex:runtime:cancel-turn',
} as const

export type CodexRuntimePublicState = 'standard' | 'starting' | 'ready' | 'failed_safe'

export interface CodexRuntimeStatus {
  readonly activeAgentRuntime: AgentRuntimeMode
  readonly state: CodexRuntimePublicState
  readonly documentId: string | null
}

export interface CodexRuntimeStartRequest {
  readonly documentId: string
  readonly text: string
}

export interface CodexDocumentIdentity {
  readonly documentId: string
  readonly host: EnhancedHost
  readonly generation: number
}

export interface CodexRuntimeApi {
  status(): Promise<CodexRuntimeStatus>
  startTurn(request: CodexRuntimeStartRequest): Promise<void>
  cancelTurn(documentId: string): Promise<void>
}
