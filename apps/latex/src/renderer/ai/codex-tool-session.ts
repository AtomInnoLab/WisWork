import type { AgentSkill, AgentToolCall, ToolExecution } from '@wiswork/agent-core'
import type {
  CodexToolApi,
  CodexToolMutability,
  CodexToolRequest,
  CodexToolResponse,
} from '../../../../shell/src/shared/codex-api.js'
import {
  loadProposalForReview,
  proposalForSelection,
  type ReviewProposal,
} from './proposal-review.js'

type IpcResult<T> = { ok: true; value: T } | { ok: false; error: { message: string } }

interface PreparedMutation {
  readonly call: AgentToolCall
  readonly proposalId: string
  readonly preparationId: string
  readonly snapshotId: string
  readonly expectedRevision: string
  cleanup?: Promise<void>
}

type SuccessResponseWithoutId =
  Extract<CodexToolResponse, { ok: true }> extends infer T
    ? T extends unknown
      ? Omit<T, 'requestId'>
      : never
    : never

export interface LatexCodexMutationResult {
  readonly proposalId: string
  readonly snapshotId: string
  readonly compile: {
    readonly ok: boolean
    readonly error?: string
    readonly result?: { readonly diagnostics: readonly unknown[] }
  }
}

export interface LatexCodexDomainApi {
  getProposal(request: { projectId: string; proposalId: string }): Promise<IpcResult<unknown>>
  proposeProjectEdits(request: {
    projectId: string
    files: Array<{ path: string; afterText: string }>
  }): Promise<IpcResult<unknown>>
  discardProposal(request: { projectId: string; proposalId: string }): Promise<IpcResult<unknown>>
  getCodexMutationRevision(request: { projectId: string }): Promise<IpcResult<{ revision: string }>>
  prepareCodexProposalMutation(request: {
    projectId: string
    documentId: string
    callId: string
    proposalId: string
    expectedRevision: string
  }): Promise<IpcResult<{ preparationId: string; snapshotId: string }>>
  executeCodexProposalMutation(request: {
    projectId: string
    documentId: string
    callId: string
    preparationId: string
    snapshotId: string
    expectedRevision: string
  }): Promise<IpcResult<LatexCodexMutationResult>>
  discardCodexProposalMutation(request: {
    projectId: string
    documentId: string
    callId: string
    preparationId: string
    snapshotId: string
  }): Promise<IpcResult<unknown>>
}

export const LATEX_CODEX_TOOL_POLICY = Object.freeze({
  list_project_files: 'read',
  search_project_text: 'read',
  read_project_text: 'read',
  get_compile_diagnostics: 'read',
  compile_project: 'read',
  // Proposal creation and apply are one guarded Codex tool transaction. The approval handler
  // renders the existing review UI before a real snapshot or project mutation can occur.
  propose_project_edits: 'mutate',
}) satisfies Readonly<Record<string, CodexToolMutability>>

export interface LatexCodexToolSessionOptions {
  readonly documentId: string
  readonly projectId: string
  readonly skill: AgentSkill
  readonly tools: CodexToolApi
  readonly domain: LatexCodexDomainApi
  readonly requestReview: (
    proposal: ReviewProposal,
    signal: AbortSignal,
  ) => Promise<ReadonlySet<string> | null>
  readonly onApplied: (result: LatexCodexMutationResult) => void | Promise<void>
}

const MAX_ID_LENGTH = 256

function validId(value: unknown): value is string {
  return typeof value === 'string' && value.length > 0 && value.length <= MAX_ID_LENGTH
}

function sameCall(left: AgentToolCall, right: AgentToolCall): boolean {
  try {
    return (
      left.id === right.id &&
      left.name === right.name &&
      JSON.stringify(left.input) === JSON.stringify(right.input)
    )
  } catch {
    return false
  }
}

function proposalSummary(value: unknown): string {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error()
  const record = value as Record<string, unknown>
  if (!validId(record.id) || typeof record.expiresAt !== 'number' || !Array.isArray(record.files)) {
    throw new Error()
  }
  return JSON.stringify({
    proposalId: record.id,
    expiresAt: record.expiresAt,
    fileCount: record.files.length,
  })
}

function errorCode(error: unknown, fallback: string): string {
  if (error instanceof Error) {
    if (error.message === 'tool_cancelled') return 'tool_cancelled'
    if (/revision|changed|conflict|stale/i.test(error.message)) return 'document_changed'
    if (/closed|unavailable/i.test(error.message)) return 'tool_session_closed'
  }
  return fallback
}

export class LatexCodexToolSession {
  readonly #options: LatexCodexToolSessionOptions
  readonly #requests = new Map<string, AbortController>()
  readonly #approved = new Map<string, { call: AgentToolCall; proposal: ReviewProposal }>()
  readonly #ownedProposalIds = new Set<string>()
  readonly #prepared = new Map<string, PreparedMutation>()
  #removeRequestListener: (() => void) | undefined
  #removeCancelListener: (() => void) | undefined
  #started = false
  #closed = false

  constructor(options: LatexCodexToolSessionOptions) {
    if (!validId(options.documentId) || !validId(options.projectId)) {
      throw new TypeError('invalid_latex_codex_session')
    }
    this.#options = options
  }

  async start(): Promise<void> {
    if (this.#closed) throw new Error('tool_session_closed')
    if (this.#started) return
    this.#removeRequestListener = this.#options.tools.onRequest((request) => {
      if (request.documentId !== this.#options.documentId || this.#closed) return
      void this.#handle(request)
    })
    this.#removeCancelListener = this.#options.tools.onCancel((cancel) => {
      if (cancel.documentId !== this.#options.documentId) return
      this.#requests.get(cancel.requestId)?.abort()
    })
    try {
      await this.#options.tools.register({
        documentId: this.#options.documentId,
        skill: {
          id: this.#options.skill.id,
          systemPrompt: this.#options.skill.systemPrompt,
          tools: this.#options.skill.tools,
        },
        policy: LATEX_CODEX_TOOL_POLICY,
      })
      this.#started = true
    } catch (error) {
      this.#removeRequestListener?.()
      this.#removeCancelListener?.()
      this.#removeRequestListener = undefined
      this.#removeCancelListener = undefined
      throw error
    }
  }

  async close(): Promise<void> {
    if (this.#closed) return
    this.#closed = true
    this.#removeRequestListener?.()
    this.#removeCancelListener?.()
    for (const controller of this.#requests.values()) controller.abort()
    this.#requests.clear()
    const prepared = [...this.#prepared.values()]
    this.#prepared.clear()
    await Promise.allSettled(prepared.map((item) => this.#discardPrepared(item)))
    const proposals = [...this.#ownedProposalIds]
    this.#approved.clear()
    this.#ownedProposalIds.clear()
    await Promise.allSettled(
      proposals.map((proposalId) =>
        this.#options.domain.discardProposal({
          projectId: this.#options.projectId,
          proposalId,
        }),
      ),
    )
    if (this.#started) {
      await this.#options.tools.unregister(this.#options.documentId).catch(() => undefined)
    }
    this.#started = false
  }

  async #handle(request: CodexToolRequest): Promise<void> {
    if (!validId(request.requestId) || this.#requests.has(request.requestId)) return
    const controller = new AbortController()
    this.#requests.set(request.requestId, controller)
    try {
      const response = await this.#dispatch(request, controller.signal)
      await this.#options.tools.respond({ requestId: request.requestId, ...response })
    } catch (error) {
      const fallback = request.type === 'snapshot' ? 'snapshot_failed' : 'tool_execution_failed'
      await this.#options.tools.respond({
        requestId: request.requestId,
        ok: false,
        code: controller.signal.aborted ? 'tool_cancelled' : errorCode(error, fallback),
      })
    } finally {
      this.#requests.delete(request.requestId)
    }
  }

  async #dispatch(
    request: CodexToolRequest,
    signal: AbortSignal,
  ): Promise<SuccessResponseWithoutId> {
    if (signal.aborted || this.#closed) throw new Error('tool_cancelled')
    switch (request.type) {
      case 'execute': {
        const execution = await this.#options.skill.executeTool(request.call, signal)
        if (signal.aborted) throw new Error('tool_cancelled')
        return { ok: true, type: 'execution', execution }
      }
      case 'revision': {
        const result = await this.#options.domain.getCodexMutationRevision({
          projectId: this.#options.projectId,
        })
        if (!result.ok || !validId(result.value.revision)) throw new Error('revision_unavailable')
        return { ok: true, type: 'revision', revision: result.value.revision }
      }
      case 'approval':
        return this.#approve(request.call, request.expectedRevision, signal)
      case 'snapshot':
        return this.#capture(request.call, request.expectedRevision, signal)
      case 'discardSnapshot': {
        const prepared = this.#prepared.get(request.snapshotId)
        if (!prepared || !sameCall(prepared.call, request.call)) throw new Error('snapshot_failed')
        await this.#discardPrepared(prepared)
        this.#prepared.delete(request.snapshotId)
        this.#approved.delete(request.call.id)
        return { ok: true, type: 'discard' }
      }
      case 'executeMutation':
        return this.#executeMutation(request.call, request.guard, signal)
    }
  }

  async #approve(
    call: AgentToolCall,
    expectedRevision: string,
    signal: AbortSignal,
  ): Promise<Omit<Extract<CodexToolResponse, { ok: true; type: 'approval' }>, 'requestId'>> {
    if (call.name !== 'propose_project_edits' || !validId(expectedRevision)) {
      throw new Error('tool_execution_failed')
    }
    const execution = await this.#options.skill.executeTool(call, signal)
    if (execution.isError || signal.aborted) throw new Error('tool_execution_failed')
    const proposal = await loadProposalForReview(
      execution.output,
      this.#options.projectId,
      (request) => this.#options.domain.getProposal(request),
    )
    this.#ownedProposalIds.add(proposal.id)
    const selected = await this.#options.requestReview(proposal, signal)
    if (signal.aborted) throw new Error('tool_cancelled')
    if (selected === null) {
      await this.#discardProposal(proposal.id)
      return { ok: true, type: 'approval', approved: false }
    }
    const owned = await proposalForSelection(proposal, selected, async (files) => {
      const result = await this.#options.domain.proposeProjectEdits({
        projectId: this.#options.projectId,
        files,
      })
      if (!result.ok) throw new Error('tool_execution_failed')
      const replacement = await loadProposalForReview(
        proposalSummary(result.value),
        this.#options.projectId,
        (request) => this.#options.domain.getProposal(request),
      )
      this.#ownedProposalIds.add(replacement.id)
      return replacement
    })
    if (owned.id !== proposal.id) await this.#discardProposal(proposal.id)
    this.#approved.set(call.id, { call, proposal: owned })
    return { ok: true, type: 'approval', approved: true }
  }

  async #capture(
    call: AgentToolCall,
    expectedRevision: string,
    signal: AbortSignal,
  ): Promise<Omit<Extract<CodexToolResponse, { ok: true; type: 'snapshot' }>, 'requestId'>> {
    const approved = this.#approved.get(call.id)
    if (!approved || !sameCall(approved.call, call) || signal.aborted) {
      throw new Error('snapshot_failed')
    }
    const result = await this.#options.domain.prepareCodexProposalMutation({
      projectId: this.#options.projectId,
      documentId: this.#options.documentId,
      callId: call.id,
      proposalId: approved.proposal.id,
      expectedRevision,
    })
    if (!result.ok || !validId(result.value.preparationId) || !validId(result.value.snapshotId)) {
      throw new Error('snapshot_failed')
    }
    const prepared: PreparedMutation = {
      call,
      proposalId: approved.proposal.id,
      preparationId: result.value.preparationId,
      snapshotId: result.value.snapshotId,
      expectedRevision,
    }
    this.#prepared.set(prepared.snapshotId, prepared)
    if (this.#closed) {
      await this.#discardPrepared(prepared)
      this.#prepared.delete(prepared.snapshotId)
      throw new Error('tool_session_closed')
    }
    // If cancellation raced capture while the session remains open, return the materialized ID.
    // The safety-critical Shell request stays pending so the router can issue its one discard.
    return { ok: true, type: 'snapshot', snapshotId: prepared.snapshotId }
  }

  async #executeMutation(
    call: AgentToolCall,
    guard: { expectedRevision: string; snapshotId: string },
    signal: AbortSignal,
  ): Promise<Omit<Extract<CodexToolResponse, { ok: true; type: 'execution' }>, 'requestId'>> {
    const prepared = this.#prepared.get(guard.snapshotId)
    if (
      !prepared ||
      !sameCall(prepared.call, call) ||
      prepared.expectedRevision !== guard.expectedRevision ||
      signal.aborted ||
      this.#closed
    ) {
      throw new Error('document_changed')
    }
    // This is the final renderer cancellation check. The main-process domain method performs the
    // matching revision check while holding its mutation lock immediately before commit.
    if (signal.aborted) throw new Error('tool_cancelled')
    const onAbort = () => {
      // This call marks cancellation in the domain even if guarded execution is already awaiting
      // its transaction lock. The same promise is reused by the router's later cleanup request.
      void this.#discardPrepared(prepared).catch(() => undefined)
    }
    signal.addEventListener('abort', onAbort, { once: true })
    if (signal.aborted) onAbort()
    const result = await this.#options.domain
      .executeCodexProposalMutation({
        projectId: this.#options.projectId,
        documentId: this.#options.documentId,
        callId: call.id,
        preparationId: prepared.preparationId,
        snapshotId: prepared.snapshotId,
        expectedRevision: guard.expectedRevision,
      })
      .finally(() => signal.removeEventListener('abort', onAbort))
    if (!result.ok || result.value.snapshotId !== prepared.snapshotId) {
      throw new Error('tool_execution_failed')
    }
    this.#prepared.delete(prepared.snapshotId)
    this.#approved.delete(call.id)
    this.#ownedProposalIds.delete(prepared.proposalId)
    await this.#options.onApplied(result.value)
    const execution: ToolExecution = {
      output: JSON.stringify({
        applied: true,
        snapshotId: result.value.snapshotId,
        compileOk: result.value.compile.ok,
        diagnostics: result.value.compile.result?.diagnostics.length ?? 0,
      }),
      summary: 'Apply reviewed LaTeX project edits',
      mutated: true,
    }
    return { ok: true, type: 'execution', execution }
  }

  async #discardPrepared(prepared: PreparedMutation): Promise<void> {
    if (!prepared.cleanup) {
      prepared.cleanup = this.#options.domain
        .discardCodexProposalMutation({
          projectId: this.#options.projectId,
          documentId: this.#options.documentId,
          callId: prepared.call.id,
          preparationId: prepared.preparationId,
          snapshotId: prepared.snapshotId,
        })
        .then((result) => {
          if (!result.ok) throw new Error('snapshot_failed')
        })
    }
    return prepared.cleanup
  }

  async #discardProposal(proposalId: string): Promise<void> {
    const result = await this.#options.domain.discardProposal({
      projectId: this.#options.projectId,
      proposalId,
    })
    if (!result.ok) throw new Error('tool_execution_failed')
    this.#ownedProposalIds.delete(proposalId)
  }
}
