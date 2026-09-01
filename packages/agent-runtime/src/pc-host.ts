import type { AgentToolCall, AgentToolDef, ToolExecution } from '@wiswork/agent-core'
import type { EnhancedHost } from './contracts'
import type { EnhancedSessionEvent } from './enhanced'

export const PC_HOST_CODEX_CHANNELS = Object.freeze({
  status: 'codex:pc-host:status',
  register: 'codex:pc-host:register',
  unregister: 'codex:pc-host:unregister',
  toolResult: 'codex:pc-host:tool-result',
  confirmProposal: 'codex:pc-host:confirm-proposal',
  cancelProposal: 'codex:pc-host:cancel-proposal',
  event: 'codex:pc-host:event',
  toolCall: 'codex:pc-host:tool-call',
  proposal: 'codex:pc-host:proposal',
})

export type PcEnhancedHost = Extract<EnhancedHost, 'latex' | 'slides' | 'docs' | 'sheets'>

export interface PcHostRegistration {
  readonly host: PcEnhancedHost
  readonly documentId: string
  readonly generation: number
  readonly systemPrompt: string
  readonly tools: readonly AgentToolDef[]
  readonly mutatingTools: readonly string[]
}

export interface PcHostToolRequest {
  readonly documentId: string
  readonly generation: number
  readonly call: AgentToolCall
}

export interface PcHostToolResult {
  readonly documentId: string
  readonly generation: number
  readonly callId: string
  readonly execution: ToolExecution
  /** Opaque renderer-local snapshot binding; never contains document state. */
  readonly snapshotBefore?: string
}

export interface PcHostProposalRequest {
  readonly proposalId: string
  readonly documentId: string
  readonly generation: number
  readonly toolName: string
  readonly summary: PcHostProposalSummary
  readonly expiresAt: number
}

export interface PcHostProposalSummary {
  readonly operation: 'insert' | 'replace' | 'delete' | 'format' | 'restructure' | 'compile'
  readonly target:
    | 'document'
    | 'selection'
    | 'blocks'
    | 'cells'
    | 'sheet'
    | 'slides'
    | 'elements'
    | 'project-files'
  readonly scope: 'single' | 'selection' | 'bounded-set' | 'whole-document'
  readonly count?: number
}

export interface PcHostCodexApi {
  status(): Promise<{
    readonly activeAgentRuntime: 'standard' | 'enhanced'
    readonly documentId: string | null
  }>
  register(input: PcHostRegistration): Promise<void>
  unregister(documentId: string, generation: number): Promise<void>
  startTurn(input: { readonly documentId: string; readonly text: string }): Promise<void>
  cancelTurn(documentId: string): Promise<void>
  toolResult(input: PcHostToolResult): Promise<void>
  confirmProposal(documentId: string, generation: number, proposalId: string): Promise<void>
  cancelProposal(documentId: string, generation: number, proposalId: string): Promise<void>
  onEvent(listener: (event: EnhancedSessionEvent) => void): () => void
  onToolCall(listener: (request: PcHostToolRequest) => void): () => void
  onProposal(listener: (request: PcHostProposalRequest) => void): () => void
}

export interface PcHostIpcRenderer {
  invoke(channel: string, ...args: unknown[]): Promise<unknown>
  on(channel: string, listener: (event: unknown, value: unknown) => void): void
  removeListener(channel: string, listener: (event: unknown, value: unknown) => void): void
}

/** Electron's stable missing-handler failure for standalone editor processes. */
export function isPcHostCodexUnavailable(error: unknown): boolean {
  const message = error instanceof Error ? error.message : String(error)
  return /No handler registered for ['"]codex:pc-host:status['"]/.test(message)
}

/** Preload-only bounded bridge; it exposes no generic IPC primitive. */
export function createPcHostCodexApi(ipc: PcHostIpcRenderer): PcHostCodexApi {
  const listen = <T>(channel: string, handler: (value: T) => void) => {
    const listener = (_event: unknown, value: unknown) => handler(value as T)
    ipc.on(channel, listener)
    return () => ipc.removeListener(channel, listener)
  }
  const api: PcHostCodexApi = {
    status: () => ipc.invoke(PC_HOST_CODEX_CHANNELS.status) as ReturnType<PcHostCodexApi['status']>,
    register: (input: PcHostRegistration) =>
      ipc.invoke(PC_HOST_CODEX_CHANNELS.register, input).then(() => undefined),
    unregister: (documentId: string, generation: number) =>
      ipc.invoke(PC_HOST_CODEX_CHANNELS.unregister, documentId, generation).then(() => undefined),
    startTurn: (input: { readonly documentId: string; readonly text: string }) =>
      ipc.invoke('codex:runtime:start-turn', input).then(() => undefined),
    cancelTurn: (documentId: string) =>
      ipc.invoke('codex:runtime:cancel-turn', documentId).then(() => undefined),
    toolResult: (input: PcHostToolResult) =>
      ipc.invoke(PC_HOST_CODEX_CHANNELS.toolResult, input).then(() => undefined),
    confirmProposal: (documentId, generation, proposalId) =>
      ipc
        .invoke(PC_HOST_CODEX_CHANNELS.confirmProposal, documentId, generation, proposalId)
        .then(() => undefined),
    cancelProposal: (documentId, generation, proposalId) =>
      ipc
        .invoke(PC_HOST_CODEX_CHANNELS.cancelProposal, documentId, generation, proposalId)
        .then(() => undefined),
    onEvent: (listener: (event: EnhancedSessionEvent) => void) =>
      listen(PC_HOST_CODEX_CHANNELS.event, listener),
    onToolCall: (listener: (request: PcHostToolRequest) => void) =>
      listen(PC_HOST_CODEX_CHANNELS.toolCall, listener),
    onProposal: (listener: (request: PcHostProposalRequest) => void) =>
      listen(PC_HOST_CODEX_CHANNELS.proposal, listener),
  }
  return Object.freeze(api)
}
