import type { AgentToolCall, AgentToolDef, ToolExecution } from '@wiswork/agent-core'

export const CODEX_TOOL_CHANNELS = {
  register: 'codex:tools:register',
  unregister: 'codex:tools:unregister',
  request: 'codex:tools:request',
  response: 'codex:tools:response',
  cancel: 'codex:tools:cancel',
} as const

export type CodexToolMutability = 'read' | 'mutate'

export interface CodexToolRegistrationRequest {
  readonly documentId: string
  readonly skill: {
    readonly id: string
    readonly systemPrompt: string
    readonly tools: readonly AgentToolDef[]
  }
  readonly policy: Readonly<Record<string, CodexToolMutability>>
}

interface CodexToolRequestBase {
  readonly requestId: string
  readonly documentId: string
}

export type CodexToolRequest =
  | (CodexToolRequestBase & { readonly type: 'revision' })
  | (CodexToolRequestBase & {
      readonly type: 'approval'
      readonly call: AgentToolCall
      readonly expectedRevision: string
    })
  | (CodexToolRequestBase & {
      readonly type: 'snapshot'
      readonly call: AgentToolCall
      readonly expectedRevision: string
    })
  | (CodexToolRequestBase & {
      readonly type: 'execute'
      readonly call: AgentToolCall
    })
  | (CodexToolRequestBase & {
      readonly type: 'executeMutation'
      readonly call: AgentToolCall
      /**
       * The renderer must atomically verify cancellation and expectedRevision immediately before
       * commit, then validate and record undo state before acknowledging this single transaction.
       */
      readonly guard: { readonly expectedRevision: string; readonly snapshotId: string }
    })

export type CodexToolResponse =
  | { readonly requestId: string; readonly ok: false; readonly code: string }
  | {
      readonly requestId: string
      readonly ok: true
      readonly type: 'revision'
      readonly revision: string
    }
  | {
      readonly requestId: string
      readonly ok: true
      readonly type: 'approval'
      readonly approved: boolean
    }
  | {
      readonly requestId: string
      readonly ok: true
      readonly type: 'snapshot'
      readonly snapshotId: string
    }
  | {
      readonly requestId: string
      readonly ok: true
      readonly type: 'execution'
      readonly execution: ToolExecution
    }

export interface CodexToolCancel {
  readonly requestId: string
  readonly documentId: string
}

export interface CodexToolApi {
  register(request: CodexToolRegistrationRequest): Promise<{ readonly registered: true }>
  unregister(documentId: string): Promise<void>
  respond(response: CodexToolResponse): Promise<boolean>
  onRequest(handler: (request: CodexToolRequest) => void): () => void
  onCancel(handler: (cancel: CodexToolCancel) => void): () => void
}
