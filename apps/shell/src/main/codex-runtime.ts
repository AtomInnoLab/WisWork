import { isAbsolute } from 'node:path'
import type { AgentRuntimeMode, EnhancedHost, EnhancedRolloutPolicy } from '@wiswork/agent-runtime'
import type { CodexRuntimePublicState } from '../shared/codex-api'

const MAX_DOCUMENT_ID_BYTES = 256
const MAX_TURN_TEXT_BYTES = 1_000_000
const MAX_DOCUMENTS = 16

export interface CodexOwner {
  isDestroyed(): boolean
}

export interface CodexRuntimeEngine {
  startTurn(input: {
    readonly documentId: string
    readonly host: EnhancedHost
    readonly generation: number
    readonly text: string
  }): Promise<void>
  cancelTurn(documentId: string): Promise<void>
  closeDocument(documentId: string): Promise<void>
  close(): Promise<void>
}

export interface CodexRuntimeBootstrap {
  start(input: {
    readonly executablePath: string
    readonly onCrash: () => void
  }): Promise<CodexRuntimeEngine>
}

export interface ShellCodexRuntimeOptions {
  readonly activeAgentRuntime: AgentRuntimeMode
  readonly policy: EnhancedRolloutPolicy
  readonly isSignedIn: () => Promise<boolean>
  readonly resolveExecutable: () => Promise<string>
  readonly bootstrap: CodexRuntimeBootstrap
  readonly diagnostics?: (code: string) => void
}

interface DocumentRecord {
  readonly owner: CodexOwner
  readonly documentId: string
  readonly host: EnhancedHost
  readonly generation: number
  busy: boolean
}

export class ShellCodexRuntimeError extends Error {
  constructor(readonly code: string) {
    super(code)
    this.name = 'ShellCodexRuntimeError'
  }
}

function boundedText(value: unknown, maximum: number): value is string {
  return typeof value === 'string' && value.length > 0 && Buffer.byteLength(value) <= maximum
}

export class ShellCodexRuntime {
  readonly configuredAgentRuntime: AgentRuntimeMode
  readonly #options: ShellCodexRuntimeOptions
  readonly #documents = new Map<string, DocumentRecord>()
  #state: CodexRuntimePublicState
  #engine: CodexRuntimeEngine | undefined
  #initialization: Promise<void> | undefined
  #closed = false
  #epoch = 0

  constructor(options: ShellCodexRuntimeOptions) {
    this.#options = options
    this.configuredAgentRuntime = options.activeAgentRuntime
    this.#state = options.activeAgentRuntime === 'standard' ? 'standard' : 'starting'
  }

  get state(): CodexRuntimePublicState {
    return this.#state
  }

  get activeAgentRuntime(): AgentRuntimeMode {
    return this.configuredAgentRuntime === 'enhanced' && this.#state === 'ready'
      ? 'enhanced'
      : 'standard'
  }

  initialize(): Promise<void> {
    if (this.#initialization) return this.#initialization
    this.#initialization = this.#initialize()
    return this.#initialization
  }

  async #initialize(): Promise<void> {
    if (this.configuredAgentRuntime === 'standard') return
    if (this.#closed) throw new ShellCodexRuntimeError('enhanced_runtime_closed')
    const epoch = this.#epoch
    try {
      if (!this.#options.policy.globalEnabled) {
        throw new ShellCodexRuntimeError('enhanced_runtime_blocked_by_policy')
      }
      if (!(await this.#options.isSignedIn())) throw new ShellCodexRuntimeError('auth_required')
      this.#assertEpoch(epoch)
      const executablePath = await this.#options.resolveExecutable()
      this.#assertEpoch(epoch)
      if (!isAbsolute(executablePath)) throw new ShellCodexRuntimeError('component_unverified')
      const engine = await this.#options.bootstrap.start({
        executablePath,
        onCrash: () => void this.#crash(),
      })
      if (this.#closed || this.#epoch !== epoch) {
        await engine.close()
        throw new ShellCodexRuntimeError('enhanced_runtime_closed')
      }
      this.#engine = engine
      this.#state = 'ready'
    } catch (error) {
      this.#state = this.#closed ? 'unavailable' : 'failed_safe'
      this.#diagnostic(
        error instanceof ShellCodexRuntimeError ? error.code : 'enhanced_start_failed',
      )
      throw error instanceof ShellCodexRuntimeError
        ? error
        : new ShellCodexRuntimeError('enhanced_start_failed')
    }
  }

  registerDocument(input: {
    readonly owner: CodexOwner
    readonly documentId: string
    readonly host: EnhancedHost
    readonly generation: number
  }): { close: () => Promise<void> } {
    if (
      this.#state !== 'ready' ||
      !this.#engine ||
      input.owner.isDestroyed() ||
      !boundedText(input.documentId, MAX_DOCUMENT_ID_BYTES) ||
      !Number.isSafeInteger(input.generation) ||
      input.generation < 0 ||
      !this.#options.policy.hosts[input.host]
    ) {
      throw new ShellCodexRuntimeError('enhanced_document_unavailable')
    }
    if (this.#documents.has(input.documentId)) {
      throw new ShellCodexRuntimeError('enhanced_document_exists')
    }
    if (this.#documents.size >= MAX_DOCUMENTS) {
      throw new ShellCodexRuntimeError('enhanced_document_limit')
    }
    this.#documents.set(input.documentId, { ...input, busy: false })
    return { close: () => this.closeDocument(input.documentId) }
  }

  ownsDocument(owner: CodexOwner, documentId: string): boolean {
    const document = this.#documents.get(documentId)
    return document?.owner === owner && !owner.isDestroyed()
  }

  async startTurn(owner: CodexOwner, documentId: string, text: string): Promise<void> {
    const document = this.#owned(owner, documentId)
    if (!boundedText(text, MAX_TURN_TEXT_BYTES)) {
      throw new ShellCodexRuntimeError('enhanced_invalid_turn')
    }
    if (document.busy) throw new ShellCodexRuntimeError('enhanced_turn_in_progress')
    const engine = this.#engine
    if (!engine) throw new ShellCodexRuntimeError('enhanced_runtime_unavailable')
    document.busy = true
    const epoch = this.#epoch
    try {
      await engine.startTurn({
        documentId,
        host: document.host,
        generation: document.generation,
        text,
      })
      if (
        this.#epoch !== epoch ||
        this.#engine !== engine ||
        this.#documents.get(documentId) !== document
      ) {
        throw new ShellCodexRuntimeError('enhanced_turn_stale')
      }
    } catch {
      // Never replay or cross-dispatch the same request through Standard.
      throw new ShellCodexRuntimeError('enhanced_turn_failed')
    } finally {
      if (this.#documents.get(documentId) === document) document.busy = false
    }
  }

  async cancelTurn(owner: CodexOwner, documentId: string): Promise<void> {
    const document = this.#owned(owner, documentId)
    if (!document.busy) return
    await this.#engine?.cancelTurn(documentId)
  }

  async closeDocument(documentId: string): Promise<void> {
    const document = this.#documents.get(documentId)
    if (!document) return
    this.#documents.delete(documentId)
    if (document.busy) await this.#engine?.cancelTurn(documentId).catch(() => undefined)
    await this.#engine?.closeDocument(documentId).catch(() => undefined)
  }

  async logout(): Promise<void> {
    await this.shutdown()
  }

  async shutdown(): Promise<void> {
    if (this.#closed) return
    this.#closed = true
    this.#epoch += 1
    for (const documentId of [...this.#documents.keys()]) await this.closeDocument(documentId)
    const engine = this.#engine
    this.#engine = undefined
    await engine?.close().catch(() => undefined)
    if (this.configuredAgentRuntime === 'enhanced') this.#state = 'unavailable'
  }

  #owned(owner: CodexOwner, documentId: string): DocumentRecord {
    const document = this.#documents.get(documentId)
    if (!document || document.owner !== owner || owner.isDestroyed()) {
      throw new ShellCodexRuntimeError('enhanced_document_unavailable')
    }
    return document
  }

  async #crash(): Promise<void> {
    if (this.#closed) return
    this.#epoch += 1
    this.#state = 'failed_safe'
    const documentIds = [...this.#documents.keys()]
    this.#documents.clear()
    const engine = this.#engine
    this.#engine = undefined
    for (const documentId of documentIds) {
      await engine?.closeDocument(documentId).catch(() => undefined)
    }
    await engine?.close().catch(() => undefined)
    this.#diagnostic('enhanced_runtime_crashed')
  }

  #assertEpoch(epoch: number): void {
    if (this.#closed || this.#epoch !== epoch) {
      throw new ShellCodexRuntimeError('enhanced_runtime_closed')
    }
  }

  #diagnostic(code: string): void {
    try {
      this.#options.diagnostics?.(code)
    } catch {
      // Diagnostics cannot affect runtime state.
    }
  }
}
