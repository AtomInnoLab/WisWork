import type { AgentSkill, PresentationTaskPreparation } from './skill'
import {
  parsePresentationAcceptanceContract,
  parsePresentationCompletionReceipt,
  renderPresentationCompletionFacts,
  type PresentationAcceptanceContract,
  type PresentationCompletionFacts,
} from '@wiswork/presentation-verification'
import type {
  AgentImage,
  AgentMessage,
  AgentStreamHandle,
  AgentToolCall,
  AgentToolContent,
  AgentToolResult,
  AgentTransport,
  ToolExecution,
  ToolExecutionOutcome,
  ToolExecutionSuspension,
} from './types'
import {
  isToolExecutionSuspension,
  isToolExecutionSuspensionOwnedBy,
  mintLoopToolExecutionSuspension,
} from './types'

export interface ToolExecutedEvent<TSnapshot> {
  call: AgentToolCall
  execution: ToolExecution
  /**
   * Snapshot captured just before this tool ran; present only on the first
   * mutating tool of a run (hook for one-click rollback UIs).
   */
  snapshotBefore?: TSnapshot | undefined
}

export interface AgentRunResult {
  /** final assistant text of the run ('' when cut off) */
  text: string
  cancelled: boolean
  /** true when maxTurns was reached; text is the partial answer from the no-tools finalizing turn */
  turnLimit: boolean
  /** the final turn hit the token limit (stop_reason max_tokens): text is incomplete; set only when true */
  truncated?: boolean
  /** Authoritative, receipt-derived facts for presentation mutation runs. */
  presentation?: PresentationCompletionFacts
  clarification?: true
}

export interface AgentLoopEvents<TSnapshot> {
  /** cumulative assistant text of the current turn (call per delta) */
  onText?(text: string): void
  /** a tool is about to execute (UI shows a live "running" indicator; onToolExecuted always follows) */
  onToolStart?(call: AgentToolCall): void
  onToolExecuted?(event: ToolExecutedEvent<TSnapshot>): void
  /** a turn requested tools and they ran; the loop is going back to the model */
  onTurnEnd?(): void
  onDone?(result: AgentRunResult): void
  onError?(error: string): void
  onPresentationClarify?(event: { question: string }): void
  onPresentationPlan?(event: { steps: string[]; requiresConfirmation: boolean }): void
  onPresentationCorrection?(event: { pass: number; maximum: number }): void
  onPresentationReceipt?(event: {
    receipt: import('@wiswork/presentation-verification').PresentationCompletionReceipt
    facts: PresentationCompletionFacts
  }): string | undefined
  /** Host audit sink for a proved mutation whose UI session was reset during reconciliation. */
  onAbandonedPresentationCompletion?(event: {
    documentToken: string
    sessionToken: string
    receipt: import('@wiswork/presentation-verification').PresentationCompletionReceipt
    facts: PresentationCompletionFacts
  }): void
}

/** Context compaction config (budget tracked in UTF-8 bytes rather than message count) */
export interface CompactionOptions {
  /** History size that triggers compaction (UTF-8 bytes, default 256KB) */
  maxBytes?: number
  /** Size of recent messages kept after compaction (bytes, default 96KB, cut at a user boundary) */
  keepRecentBytes?: number
  /** Disable LLM summarization and use only the mechanical digest (for tests/offline) */
  disableLlmSummary?: boolean
}

export interface AgentLoopOptions<TSnapshot = unknown> {
  transport: AgentTransport
  skill: AgentSkill
  events?: AgentLoopEvents<TSnapshot>
  /** optional hard cap on model round-trips per run; interactive Cowork runs are unbounded by default */
  maxTurns?: number
  /** history cap in messages, trimmed at user-turn boundaries (default 40) */
  maxHistory?: number
  /** Context compaction; false disables it (enabled by default with default thresholds) */
  compaction?: CompactionOptions | false
  /** capture rollback state; invoked right before tools run (see snapshotBefore) */
  captureSnapshot?(): TSnapshot
  /** wrap instruction + skill context into the user message text */
  formatUserMessage?(instruction: string, context: string): string
  /** appended to the system prompt each turn (e.g. reply-language directive following the UI language) */
  systemSuffix?(): string
}

const COMPACT_MAX_BYTES = 256 * 1024
const COMPACT_KEEP_RECENT_BYTES = 96 * 1024
/** Pre-truncation of each tool output in the summary request (the compaction request itself must not blow up on huge outputs) */
const SUMMARIZE_TOOL_OUTPUT_MAX = 2_000
const SUMMARIZE_TIMEOUT_MS = 30_000
/** When over budget mid-run, keep the last N tool messages verbatim and truncate earlier outputs to this length */
const STALE_TOOL_KEEP_RECENT = 2
const STALE_TOOL_OUTPUT_MAX = 1_000

/** Cap on consecutive tool-input parse failures (a successful parse resets it); abort beyond it (keeps the model from burning turns on bad JSON) */
const MAX_INPUT_PARSE_RETRIES = 3
/** Stop only a proven unchanged loop; varied long-running Cowork work remains unbounded. */
const MAX_IDENTICAL_TOOL_BATCHES = 4
const MAX_TOOL_CONTENT_IMAGES = 4
const MAX_TOOL_IMAGE_BYTES = 4 * 1024 * 1024
const TOOL_IMAGE_MIMES = new Set(['image/png', 'image/jpeg', 'image/webp'])
/** Review policies need denial language, not an unbounded copy of model output. */
const FINAL_RESPONSE_REVIEW_MAX_CHARS = 4_096
const FINAL_RESPONSE_REVIEW_TRUNCATION = '\n[response truncated for review]\n'
/** Static corrective guidance should remain much smaller than a normal prompt. */
const FINAL_RESPONSE_CORRECTION_MAX_CHARS = 2_000
const FINAL_RESPONSE_CORRECTION_MAX_BYTES = 4_000
const PRESENTATION_QUESTION_MAX_CHARS = 1_000
const PRESENTATION_PLAN_MAX_STEPS = 12
const PRESENTATION_PLAN_STEP_MAX_CHARS = 500

function isFinalToolExecution(value: unknown): value is ToolExecution {
  if (typeof value !== 'object' || value === null) return false
  const execution = value as Partial<ToolExecution>
  return (
    typeof execution.output === 'string' &&
    typeof execution.summary === 'string' &&
    (execution.isError === undefined || typeof execution.isError === 'boolean') &&
    (execution.mutated === undefined || typeof execution.mutated === 'boolean') &&
    (execution.stopToolBatch === undefined || typeof execution.stopToolBatch === 'boolean')
  )
}

const INVALID_TOOL_OUTPUT: ToolExecution = {
  output: 'invalid_tool_output',
  isError: true,
  summary: 'invalid tool output',
}

function boundedToolContent(content?: AgentToolContent[]): AgentToolContent[] | undefined {
  if (!content?.length) return undefined
  if (content.length > MAX_TOOL_CONTENT_IMAGES) throw new Error('invalid_tool_output')
  return content.map((block) => {
    const { base64, mime } = block.image
    const padding = base64.endsWith('==') ? 2 : base64.endsWith('=') ? 1 : 0
    const bytes = (base64.length / 4) * 3 - padding
    if (
      block.type !== 'image' ||
      !TOOL_IMAGE_MIMES.has(mime) ||
      base64.length % 4 !== 0 ||
      !/^[A-Za-z0-9+/]*={0,2}$/.test(base64) ||
      bytes <= 0 ||
      bytes > MAX_TOOL_IMAGE_BYTES
    )
      throw new Error('invalid_tool_output')
    return { type: 'image', image: { base64, mime } }
  })
}

const TURN_LIMIT_NOTE =
  '[System] The tool-call turn limit for this request has been reached; no more tools may be called this turn. ' +
  'Answer directly from the information already gathered; if the task is unfinished, briefly state what is done and what remains.'

/**
 * Terminal assistant text when tools mutated the artifact (or an edits-only
 * turn was restored) and the model returned no prose. Must be non-empty so
 * provider message converters never emit empty assistant content, which breaks
 * multi-turn follow-ups (see finishTurn / restore).
 * Exported so apps can substitute a localized / tool-derived summary in the UI.
 */
export const COMPLETED_VIA_TOOLS_TEXT = '(completed tool actions; no text reply)'

/** Locale-neutral authoritative terminal text; UIs may localize from the adjacent facts. */
export function renderPresentationCompletionText(facts: PresentationCompletionFacts): string {
  return [
    `presentation:${facts.status}`,
    `slides=${facts.affectedSlides.join(',')}`,
    `passed=${facts.passedCount}`,
    `failed=${facts.failedCount}`,
    `unavailable=${facts.unavailableCount}`,
    `corrections=${facts.correctionPasses}`,
    `rollback=${facts.rollbackAvailable}`,
    ...(facts.safeCode ? [`code=${facts.safeCode}`] : []),
  ].join(';')
}

const SUMMARIZE_SYSTEM =
  'You are a conversation compressor. Compress this editing session between the user and the AI assistant into a concise summary so later turns can continue with context. ' +
  "Keep: the user's goals and key instructions, completed changes (which files/pages/elements were modified), important facts and data, and outstanding items. " +
  'For specific figures/statistics, mark their provenance: figures from the user or from tool results (e.g. web_search) keep their source; figures the assistant produced without a source must be marked "(unverified)" so later turns do not treat them as established facts. ' +
  'Omit: pleasantries, tool-call details, and intermediate trial and error. Use a bullet list of at most 400 words. Write the summary in the same language as the conversation. Output only the summary body, with no preamble.'

/** Prefix of the synthetic user message that carries the compacted-history summary */
const COMPACT_SUMMARY_PREFIX = '[Summary of earlier conversation'
const COMPACT_SUMMARY_HEADER = '[Summary of earlier conversation (auto-compacted)]'
const COMPACT_SUMMARY_ACK = 'Understood, continuing from the progress so far.'
const PRESENTATION_ENROLLMENT_MAX_BYTES = 64 * 1024

/** Approximate UTF-8 byte count (ASCII 1 byte, CJK etc. 3; surrogate pairs count as 6 — slight overestimate is harmless) */
function utf8Size(s: string): number {
  let n = 0
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i)
    n += c < 0x80 ? 1 : c < 0x800 ? 2 : 3
  }
  return n
}

/** Keep both the opening claim and closing rationale where denials commonly appear. */
function boundedFinalResponseForReview(text: string): string {
  if (text.length <= FINAL_RESPONSE_REVIEW_MAX_CHARS) return text
  const available = FINAL_RESPONSE_REVIEW_MAX_CHARS - FINAL_RESPONSE_REVIEW_TRUNCATION.length
  const head = Math.floor(available / 2)
  const tail = available - head
  return `${text.slice(0, head)}${FINAL_RESPONSE_REVIEW_TRUNCATION}${text.slice(-tail)}`
}

function safeFinalResponseCorrection(value: unknown): string | undefined {
  if (
    typeof value !== 'string' ||
    !value.trim() ||
    value.length > FINAL_RESPONSE_CORRECTION_MAX_CHARS ||
    utf8Size(value) > FINAL_RESPONSE_CORRECTION_MAX_BYTES
  )
    return undefined
  const sanitized = sanitizeAgentPayload(value).trim()
  if (
    !sanitized ||
    sanitized.length > FINAL_RESPONSE_CORRECTION_MAX_CHARS ||
    utf8Size(sanitized) > FINAL_RESPONSE_CORRECTION_MAX_BYTES
  )
    return undefined
  return sanitized
}

function stableJson(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(stableJson).join(',')}]`
  if (value && typeof value === 'object')
    return `{${Object.entries(value as Record<string, unknown>)
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([key, item]) => `${JSON.stringify(key)}:${stableJson(item)}`)
      .join(',')}}`
  return JSON.stringify(value) ?? 'null'
}

function imagePayloadSize(image: AgentImage): number {
  const padding = image.base64.endsWith('==') ? 2 : image.base64.endsWith('=') ? 1 : 0
  const decoded = Math.max(0, (image.base64.length / 4) * 3 - padding)
  // Count both the encoded request and decoded media footprints so images
  // cannot bypass the text-oriented history budget.
  return utf8Size(image.base64) + decoded + utf8Size(image.mime) + 32
}

/** Approximate byte cost of one message (text + tool inputs/outputs + image base64) */
function messageSize(m: AgentMessage): number {
  if (m.role === 'tool') {
    return m.results.reduce(
      (size, result) =>
        size +
        utf8Size(result.output) +
        40 +
        (result.content?.reduce(
          (mediaSize, block) => mediaSize + imagePayloadSize(block.image),
          0,
        ) ?? 0),
      0,
    )
  }
  let n = utf8Size(m.text)
  if (m.role === 'user' && m.images) {
    n += m.images.reduce((size, image) => size + imagePayloadSize(image), 0)
  }
  if (m.role === 'assistant' && m.toolCalls) {
    for (const c of m.toolCalls) {
      try {
        n += utf8Size(JSON.stringify(c.input)) + 40
      } catch {
        n += 40
      }
    }
  }
  return n
}

function historySize(messages: readonly AgentMessage[]): number {
  return messages.reduce((n, m) => n + messageSize(m), 0)
}

/** Mechanical digest when LLM summarization is unavailable: bullet list of user instructions + final replies */
function mechanicalDigest(dropped: readonly AgentMessage[]): string {
  const lines: string[] = []
  for (const m of dropped) {
    if (m.role === 'user' && !m.text.startsWith(COMPACT_SUMMARY_PREFIX)) {
      lines.push(`- User: ${m.text.slice(0, 200)}`)
    } else if (m.role === 'assistant' && m.text && !m.toolCalls?.length) {
      lines.push(`  Reply: ${m.text.slice(0, 200)}`)
    }
  }
  return lines.join('\n').slice(0, 4_000) || '(earlier conversation omitted)'
}

/**
 * Generic ReAct loop: user message -> model turn (text + tool calls) ->
 * execute tools -> feed results back -> repeat until the model answers with
 * plain text. History persists across runs, so follow-up questions work.
 */
export class AgentLoop<TSnapshot = unknown> {
  private readonly options: AgentLoopOptions<TSnapshot>
  private history: AgentMessage[] = []
  private handle: AgentStreamHandle | null = null
  private running = false
  private cancelled = false
  private turns = 0
  /** Finalizing turn after hitting the turn limit: no tools, let the model answer from what it has read */
  private finalizing = false

  /** Mint a suspension bound to this loop instance; transports cannot self-authorize one. */
  suspendToolExecution(result: Promise<ToolExecution>): ToolExecutionSuspension {
    return mintLoopToolExecutionSuspension(this, result)
  }
  ownsToolExecutionSuspension(value: ToolExecutionOutcome): value is ToolExecutionSuspension {
    return isToolExecutionSuspensionOwnedBy(this, value)
  }
  /** Bounded retries let staged workflows reject more than one premature completion. */
  private completionReviewRetries = 0
  private lastCompletionReviewCorrection = ''
  private mutationSeen = false
  private presentationContract: PresentationAcceptanceContract | null = null
  private presentationCorrectionPasses = 0
  private presentationPlanEmitted = false
  private presentationCorrectionTurns = 0
  private presentationCorrectionPending = false
  private inputParseFails = 0
  private lastToolBatchSignature = ''
  private identicalToolBatches = 0
  private turnStopReason: string | null = null
  private turnText = ''
  private toolCalls: AgentToolCall[] = []
  /** user message of the in-flight run; a failed run rolls it (and everything after) back out of history */
  private runUserMsg: AgentMessage | null = null
  /** invalidates stale transport callbacks after cancel/reset */
  private generation = 0
  private readonly invocationSession = globalThis.crypto.randomUUID()
  private invocationRun = 0
  private invocationSequence = 0
  private readonly turnInvocationIds = new Map<string, string>()
  /** per-run abort: aborted on cancel(); long tools (e.g. generate_deck) use it to break internal loops */
  private abortController: AbortController | null = null
  /** Cancel stops model/tool work, but must not erase post-dispatch truth reconciliation. */
  private reconciliationController: AbortController | null = null

  constructor(options: AgentLoopOptions<TSnapshot>) {
    this.options = options
  }

  get busy(): boolean {
    return this.running
  }

  get messages(): readonly AgentMessage[] {
    return this.history
  }

  /**
   * Seed the conversation with restored history (e.g. transcript reloaded from
   * disk when a document reopens), so follow-up instructions keep their context.
   * No-op unless the loop is idle with an empty history.
   * Old messages over the compaction budget fold into a mechanical digest
   * (no LLM request on restore, guaranteeing zero latency).
   */
  restore(messages: readonly AgentMessage[]): void {
    if (this.running || this.history.length > 0 || messages.length === 0) return
    // Edits-only runs persist an assistant message with no text; give it a placeholder
    // so the turn stays paired and providers never see an empty assistant content block
    const normalized = messages.map((m) =>
      m.role === 'assistant' && !m.text ? { ...m, text: COMPLETED_VIA_TOOLS_TEXT } : m,
    )
    // Unanswered user messages (a failed or interrupted run persisted them without a
    // reply) must not re-enter the model context: trailing ones would pair with the
    // next instruction as one turn, adjacent ones read as a combined instruction
    this.history = normalized.filter(
      (m, i) => m.role !== 'user' || (normalized[i + 1] && normalized[i + 1]!.role !== 'user'),
    )
    if (this.history.length === 0) return
    if (this.compactionEnabled()) {
      const { maxBytes, keepRecentBytes } = this.compactBudget()
      if (historySize(this.history) > maxBytes) {
        const cut = this.findCompactCut(keepRecentBytes)
        if (cut > 0) {
          const digest = mechanicalDigest(this.history.slice(0, cut))
          this.history = [
            { role: 'user', text: `${COMPACT_SUMMARY_HEADER}\n${digest}` },
            { role: 'assistant', text: COMPACT_SUMMARY_ACK },
            ...this.history.slice(cut),
          ]
        }
      }
    }
    this.trimHistory()
  }

  /** images: inline attachments for this user turn (vision input; see AgentImage) */
  run(instruction: string, images?: AgentImage[]): void {
    if (this.running || !instruction) return
    this.running = true
    this.cancelled = false
    this.turns = 0
    this.finalizing = false
    this.completionReviewRetries = 0
    this.lastCompletionReviewCorrection = ''
    this.mutationSeen = false
    this.presentationContract = null
    this.presentationCorrectionPasses = 0
    this.presentationPlanEmitted = false
    this.presentationCorrectionTurns = 0
    this.presentationCorrectionPending = false
    this.inputParseFails = 0
    this.lastToolBatchSignature = ''
    this.identicalToolBatches = 0
    this.abortController = new AbortController()
    this.reconciliationController = new AbortController()
    this.invocationRun += 1
    try {
      const context = this.options.skill.buildContext?.() ?? ''
      const format =
        this.options.formatUserMessage ??
        ((instr: string, ctx: string) => (ctx ? `${instr}\n\n${ctx}` : instr))
      const userMsg: AgentMessage = {
        role: 'user',
        text: format(instruction, context),
        ...(images?.length ? { images } : {}),
      }
      void this.beginRun(userMsg)
    } catch (error) {
      this.running = false
      this.abortController = null
      throw error
    }
  }

  /** Compact (if needed), push the user message, then start the turn. Compaction failure doesn't block the run. */
  private async beginRun(userMsg: AgentMessage): Promise<void> {
    const generation = this.generation
    try {
      await this.maybeCompact()
    } catch {
      // Proceed with the run even if compaction fails (an over-budget history only costs more, it's still correct)
    }
    if (generation !== this.generation) return // reset during compaction
    if (this.cancelled) {
      this.running = false
      this.options.events?.onDone?.({ text: '', cancelled: true, turnLimit: false })
      return
    }
    if (
      this.options.skill.presentation &&
      !(await this.preparePresentationRun(userMsg.role === 'user' ? userMsg.text : '', generation))
    )
      return
    // Leftover unanswered user message (a previous run failed before replying):
    // drop it so the model never sees two adjacent user turns as one combined instruction
    while (this.history.at(-1)?.role === 'user') this.history.pop()
    this.trimHistory()
    if (userMsg.role === 'user') {
      userMsg = { ...userMsg, text: sanitizeAgentPayload(userMsg.text) }
    }
    this.runUserMsg = userMsg
    this.history.push(userMsg)
    try {
      this.startTurn()
    } catch (error) {
      this.handle = null
      this.running = false
      this.abortController = null
      this.rollbackFailedRun()
      const message = error instanceof Error ? error.message : String(error)
      try {
        this.options.events?.onError?.(message)
      } catch {
        // A presentation callback must not turn a handled launch failure into an unhandled rejection.
      }
    }
  }

  private async preparePresentationRun(instruction: string, generation: number): Promise<boolean> {
    const hooks = this.options.skill.presentation
    if (!hooks) return true
    let prepared: Awaited<ReturnType<typeof hooks.prepare>>
    try {
      prepared = await hooks.prepare(instruction, this.abortController?.signal)
    } catch {
      // Planning is additive and may fail open only before mutation dispatch.
      return generation === this.generation && !this.cancelled
    }
    if (generation !== this.generation || !this.running) return false
    if (this.cancelled) {
      this.running = false
      this.abortController = null
      this.options.events?.onDone?.({ text: '', cancelled: true, turnLimit: false })
      return false
    }
    if (prepared.kind === 'bypass') return true
    if (prepared.kind === 'clarify') {
      const question = this.boundedPresentationText(
        prepared.question,
        PRESENTATION_QUESTION_MAX_CHARS,
      )
      if (!question) return true
      this.running = false
      this.abortController = null
      this.options.events?.onPresentationClarify?.({ question })
      this.options.events?.onDone?.({
        text: '',
        cancelled: false,
        turnLimit: false,
        clarification: true,
      })
      return false
    }
    let contract: PresentationAcceptanceContract
    try {
      contract = parsePresentationAcceptanceContract(prepared.contract)
    } catch {
      // A host compiler failure is still pre-dispatch, so preserve legacy behavior.
      return true
    }
    const steps = (prepared.plan ?? [])
      .slice(0, PRESENTATION_PLAN_MAX_STEPS)
      .map((step) => this.boundedPresentationText(step, PRESENTATION_PLAN_STEP_MAX_CHARS))
      .filter((step): step is string => !!step)
    const requiresConfirmation = prepared.requiresConfirmation === true
    if (steps.length && !this.presentationPlanEmitted) {
      this.presentationPlanEmitted = true
      this.options.events?.onPresentationPlan?.({ steps, requiresConfirmation })
    }
    if (requiresConfirmation) {
      const confirmed = await (async () => {
        try {
          return (await hooks.confirm?.(contract, this.abortController?.signal)) === true
        } catch {
          return false
        }
      })()
      if (generation !== this.generation || !this.running) return false
      if (!confirmed || this.cancelled) {
        this.running = false
        this.abortController = null
        this.options.events?.onDone?.({ text: '', cancelled: this.cancelled, turnLimit: false })
        return false
      }
    }
    this.presentationContract = contract
    return true
  }

  private boundedPresentationText(value: unknown, maxChars: number): string | undefined {
    if (typeof value !== 'string' || !value.trim() || value.length > maxChars) return undefined
    const safe = sanitizeAgentPayload(value).trim()
    return safe && safe.length <= maxChars ? safe : undefined
  }

  /**
   * A run failed: remove its user message and every message after it, so the
   * failed instruction can't be silently re-executed by the next run.
   */
  private rollbackFailedRun(): void {
    const msg = this.runUserMsg
    this.runUserMsg = null
    if (!msg) return
    const i = this.history.lastIndexOf(msg)
    if (i >= 0) this.history.splice(i)
  }

  /** Settle a detached async run failure without letting it strand the loop busy. */
  private failRun(error: unknown, generation: number): void {
    if (generation !== this.generation || !this.running) return
    this.handle = null
    this.running = false
    this.abortController?.abort()
    this.abortController = null
    this.rollbackFailedRun()
    const message = error instanceof Error ? error.message : String(error)
    try {
      this.options.events?.onError?.(message)
    } catch {
      // Presentation callbacks are outside the run's trust boundary.
    }
  }

  // ── Context compaction: fold old conversation into a summary, keep recent messages verbatim ──

  private compactionEnabled(): boolean {
    return this.options.compaction !== false
  }

  private compactBudget(): { maxBytes: number; keepRecentBytes: number } {
    const opt = this.options.compaction === false ? undefined : this.options.compaction
    return {
      maxBytes: opt?.maxBytes ?? COMPACT_MAX_BYTES,
      keepRecentBytes: opt?.keepRecentBytes ?? COMPACT_KEEP_RECENT_BYTES,
    }
  }

  /**
   * Find the compaction cut at a user boundary: accumulate from the tail up to keepRecentBytes.
   * Returns the start index of the kept segment; if no suitable boundary exists,
   * fall back to keeping the last user turn.
   */
  private findCompactCut(keepRecentBytes: number): number {
    let kept = 0
    let cut = -1
    for (let i = this.history.length - 1; i >= 0; i--) {
      kept += messageSize(this.history[i]!)
      if (kept > keepRecentBytes && cut >= 0) break
      if (this.history[i]!.role === 'user') cut = i
    }
    if (cut < 0) {
      for (let i = this.history.length - 1; i >= 0; i--) {
        if (this.history[i]!.role === 'user') return i
      }
    }
    return cut
  }

  private async maybeCompact(): Promise<void> {
    if (!this.compactionEnabled()) return
    const { maxBytes, keepRecentBytes } = this.compactBudget()
    if (historySize(this.history) <= maxBytes) return
    const cut = this.findCompactCut(keepRecentBytes)
    if (cut <= 0) return // no foldable prefix
    const dropped = this.history.slice(0, cut)
    const opt = this.options.compaction === false ? undefined : this.options.compaction
    let summary: string | null = null
    if (!opt?.disableLlmSummary) summary = await this.summarizeViaLlm(dropped)
    if (!summary) summary = mechanicalDigest(dropped)
    this.history = [
      { role: 'user', text: `${COMPACT_SUMMARY_HEADER}\n${summary}` },
      { role: 'assistant', text: COMPACT_SUMMARY_ACK },
      ...this.history.slice(cut),
    ]
  }

  /** Hand the folded conversation to the model for a summary; returns null on failure/timeout (falls back to the mechanical digest). */
  private summarizeViaLlm(dropped: readonly AgentMessage[]): Promise<string | null> {
    // Slim down the summary request itself: pre-truncate tool outputs, strip images
    const slim: AgentMessage[] = dropped.map((m) => {
      if (m.role === 'tool') {
        return {
          role: 'tool' as const,
          results: m.results.map(({ content: _content, ...result }) => ({
            ...result,
            output: result.output.slice(0, SUMMARIZE_TOOL_OUTPUT_MAX),
          })),
        }
      }
      if (m.role === 'user' && m.images?.length) return { role: 'user' as const, text: m.text }
      return m
    })
    return new Promise((resolve) => {
      let text = ''
      let settled = false
      const finish = (v: string | null) => {
        if (settled) return
        settled = true
        clearTimeout(timer)
        resolve(v)
      }
      const timer = setTimeout(() => finish(null), SUMMARIZE_TIMEOUT_MS)
      try {
        // Attach to this.handle so cancel() can abort the summary request when the user clicks stop
        this.handle = this.options.transport.stream(
          {
            system: SUMMARIZE_SYSTEM,
            messages: [
              ...slim,
              { role: 'user', text: 'Compress the conversation above as instructed.' },
            ],
            tools: [],
          },
          {
            onDelta: (t) => {
              text += t
            },
            onToolCall: () => {
              /* the summary turn gets no tools */
            },
            onDone: () => finish(text.trim() || null),
            onError: () => finish(null),
          },
        )
      } catch {
        finish(null)
      }
    })
  }

  /**
   * When over budget mid-run (between tool turns), truncate stale tool outputs:
   * keep structure (tool_use/tool_result pairs intact), cut content only,
   * and keep the most recent N verbatim.
   */
  private squashStaleToolOutputs(): void {
    if (!this.compactionEnabled()) return
    const { maxBytes } = this.compactBudget()
    if (historySize(this.history) <= maxBytes) return
    // Media from prior tool rounds has already been shown to the model. Drop it
    // oldest-first while preserving the newest round for its first provider turn.
    let newestToolIndex = -1
    for (let i = this.history.length - 1; i >= 0; i--) {
      if (this.history[i]!.role !== 'tool') continue
      newestToolIndex = i
      break
    }
    for (let i = 0; i < newestToolIndex && historySize(this.history) > maxBytes; i++) {
      const message = this.history[i]!
      if (message.role !== 'tool') continue
      message.results = message.results.map((result) =>
        result.content?.length ? { ...result, content: undefined } : result,
      )
    }
    let recent = 0
    for (let i = this.history.length - 1; i >= 0; i--) {
      const m = this.history[i]!
      if (m.role !== 'tool') continue
      recent++
      if (recent <= STALE_TOOL_KEEP_RECENT) continue
      m.results = m.results.map((r) =>
        r.output.length > STALE_TOOL_OUTPUT_MAX
          ? {
              ...r,
              output: `${r.output.slice(0, STALE_TOOL_OUTPUT_MAX)}\n…(output truncated: too long)`,
            }
          : r,
      )
    }
  }

  cancel(): void {
    if (!this.running) return
    this.cancelled = true
    // abort lets long tools mid-execution (internal LLM loops etc.) stop promptly
    this.abortController?.abort()
    // the transport emits onDone after aborting, which finalizes the run
    this.handle?.cancel()
  }

  /** drop the conversation (e.g. when a different document is opened) */
  reset(): void {
    this.options.skill.presentation?.abandon?.()
    this.generation++
    this.abortController?.abort()
    if (!this.mutationSeen) this.reconciliationController?.abort()
    this.handle?.cancel()
    this.handle = null
    this.running = false
    this.cancelled = false
    this.history = []
    this.runUserMsg = null
    this.reconciliationController = null
  }

  /** Runs at run boundaries only (restore / before a new user message): a long run's tail is all assistant/tool messages, and cutting mid-run would empty the request. */
  private trimHistory(): void {
    const max = this.options.maxHistory ?? 40
    if (this.history.length <= max) return
    // cut only at a user message so tool_use/tool_result pairs stay intact
    let i = this.history.length - max
    while (i < this.history.length && this.history[i]!.role !== 'user') i++
    if (i >= this.history.length) return // no user boundary in the window: keep history over budget
    const next = this.history.slice(i)
    if (this.runUserMsg && !next.includes(this.runUserMsg)) return
    this.history = next
  }

  private startTurn(): void {
    const generation = this.generation
    this.turnText = ''
    this.toolCalls = []
    this.turnInvocationIds.clear()
    this.turnStopReason = null
    // Some transports emit an extra onDone after cancel — this turn may finalize only once
    let settled = false
    this.handle = this.options.transport.stream(
      {
        system: this.options.skill.systemPrompt + (this.options.systemSuffix?.() ?? ''),
        messages: [...this.history],
        tools: this.finalizing ? [] : this.options.skill.tools,
      },
      {
        onDelta: (text) => {
          if (generation !== this.generation || settled) return
          this.turnText += text
          this.options.events?.onText?.(this.turnText)
        },
        onToolCall: (call) => {
          if (generation !== this.generation || settled) return
          const signature = stableJson({
            id: call.id,
            name: call.name,
            input: call.input,
            inputError: call.inputError,
            truncated: call.truncated,
          })
          let invocationId = this.turnInvocationIds.get(signature)
          if (!invocationId) {
            this.invocationSequence += 1
            invocationId = `${this.invocationSession}-${this.invocationRun}-${this.invocationSequence}`
            this.turnInvocationIds.set(signature, invocationId)
          }
          const invokedCall = { ...call }
          Object.defineProperty(invokedCall, 'invocationId', {
            configurable: false,
            enumerable: false,
            writable: false,
            value: invocationId,
          })
          this.toolCalls.push(invokedCall)
        },
        onStopReason: (reason) => {
          if (generation !== this.generation || settled) return
          this.turnStopReason = reason
        },
        onDone: () => {
          if (generation !== this.generation || settled) return
          settled = true
          void this.finishTurn().catch((error) => this.failRun(error, generation))
        },
        onError: (error) => {
          if (generation !== this.generation || settled) return
          settled = true
          this.running = false
          this.rollbackFailedRun()
          this.options.events?.onError?.(error)
        },
      },
    )
  }

  private async finishTurn(): Promise<void> {
    const { events, skill, captureSnapshot } = this.options
    const toolCalls = this.toolCalls

    if (toolCalls.length > 0 && !this.cancelled && !this.finalizing) {
      const signature = stableJson(
        toolCalls.map((call) => ({
          name: call.name,
          input: call.input,
          inputError: call.inputError,
          truncated: call.truncated,
        })),
      )
      this.identicalToolBatches =
        signature === this.lastToolBatchSignature ? this.identicalToolBatches + 1 : 1
      this.lastToolBatchSignature = signature
      if (this.identicalToolBatches > MAX_IDENTICAL_TOOL_BATCHES) {
        this.running = false
        this.rollbackFailedRun()
        events?.onError?.('tool_loop_detected')
        return
      }
    }

    // Presentation mutation runs cannot cross the terminal boundary until the
    // host reconciles authoritative state into a contract-bound receipt.
    if (
      (toolCalls.length === 0 || this.cancelled || this.finalizing) &&
      this.presentationContract
    ) {
      const handled = await this.finishPresentationRun()
      if (handled) return
    }

    // A skill may reject one normal tool-free terminal response. Preserve the
    // rejected assistant prose in model history, pair it with a synthetic user
    // correction, and continue the same run without settling the UI turn.
    if (
      toolCalls.length === 0 &&
      !this.cancelled &&
      !this.finalizing &&
      this.completionReviewRetries < 3 &&
      skill.reviewFinalResponse
    ) {
      const generation = this.generation
      let correction: unknown
      try {
        correction = skill.reviewFinalResponse({
          text: boundedFinalResponseForReview(this.turnText),
          mutated: this.mutationSeen,
        })
      } catch {
        // Completion policies are advisory. A broken policy must not strand or
        // fail an otherwise valid run.
      }
      if (generation !== this.generation || !this.running) return
      const safeCorrection = safeFinalResponseCorrection(correction)
      if (
        safeCorrection &&
        safeCorrection !== this.lastCompletionReviewCorrection &&
        !this.cancelled
      ) {
        this.completionReviewRetries++
        this.lastCompletionReviewCorrection = safeCorrection
        this.history.push({
          role: 'assistant',
          text: this.turnText || COMPLETED_VIA_TOOLS_TEXT,
        })
        this.history.push({ role: 'user', text: safeCorrection })
        // The rejected prose was already streamed into the current bubble. A
        // corrective tool turn may emit no text, so clear it explicitly now.
        events?.onText?.('')
        this.startTurn()
        return
      }
    }

    // final turn: no tools requested, the user stopped the run, or the
    // no-tools finalizing turn after hitting the limit
    // (a cancelled turn drops its tool calls — no results would follow)
    if (toolCalls.length === 0 || this.cancelled || this.finalizing) {
      // Models often end a tool-using run with an empty text turn ("I'm done").
      // Leaving assistant text empty in history then poisons the next user
      // prompt: Anthropic rejects empty content arrays, Gemini rejects empty
      // parts, and OpenAI-compatible routes send content:null with no tool_calls —
      // all of which make follow-up turns fail or return empty again (see
      // upstream#12 / #22: first prompt works, second shows "no summary").
      // Same normalization as restore(), applied unconditionally: cancelled and
      // read-only empty turns poison follow-ups just the same. onDone still
      // reports the raw turn text so app UIs keep their localized fallbacks
      // instead of surfacing this English placeholder.
      this.history.push({ role: 'assistant', text: this.turnText || COMPLETED_VIA_TOOLS_TEXT })
      this.running = false
      this.runUserMsg = null
      events?.onDone?.({
        text: this.turnText,
        cancelled: this.cancelled,
        turnLimit: this.finalizing,
        // set only when true so exact-shape consumers/tests stay unaffected
        ...(this.turnStopReason === 'max_tokens' && !this.cancelled ? { truncated: true } : {}),
      })
      return
    }

    this.history.push({ role: 'assistant', text: this.turnText, toolCalls })
    const generation = this.generation
    if (skill.presentation?.enroll) {
      let enrollment: PresentationTaskPreparation
      try {
        const serializedCalls = sanitizeAgentPayload(
          JSON.stringify(
            toolCalls.map(({ id, name, input, invocationId }) => ({
              id,
              name,
              input,
              ...(invocationId ? { invocationId } : {}),
            })),
          ),
        )
        if (
          new TextEncoder().encode(serializedCalls).byteLength > PRESENTATION_ENROLLMENT_MAX_BYTES
        )
          throw new TypeError('presentation enrollment is too large')
        const enrollmentCalls = JSON.parse(serializedCalls) as AgentToolCall[]
        enrollment = await skill.presentation.enroll(
          enrollmentCalls,
          this.presentationContract ?? undefined,
          this.abortController?.signal,
        )
      } catch {
        this.failPresentationRun('presentation_enrollment_unavailable')
        return
      }
      if (generation !== this.generation || !this.running) return
      if (enrollment.kind === 'clarify') {
        const question = this.boundedPresentationText(
          enrollment.question,
          PRESENTATION_QUESTION_MAX_CHARS,
        )
        this.running = false
        this.runUserMsg = null
        this.options.events?.onPresentationClarify?.({
          question: question ?? 'presentation_scope_required',
        })
        this.options.events?.onDone?.({
          text: '',
          cancelled: false,
          turnLimit: false,
          clarification: true,
        })
        return
      }
      if (enrollment.kind === 'ready') {
        let enrolled: PresentationAcceptanceContract
        try {
          enrolled = parsePresentationAcceptanceContract(enrollment.contract)
        } catch {
          this.failPresentationRun('presentation_enrollment_invalid')
          return
        }
        if (
          this.presentationContract &&
          JSON.stringify(this.presentationContract) !== JSON.stringify(enrolled)
        ) {
          this.failPresentationRun('presentation_scope_expansion')
          return
        }
        this.presentationContract = enrolled
        const steps = (enrollment.plan ?? [])
          .slice(0, PRESENTATION_PLAN_MAX_STEPS)
          .map((step) => this.boundedPresentationText(step, PRESENTATION_PLAN_STEP_MAX_CHARS))
          .filter((step): step is string => !!step)
        if (steps.length && !this.presentationPlanEmitted) {
          this.presentationPlanEmitted = true
          this.options.events?.onPresentationPlan?.({
            steps,
            requiresConfirmation: enrollment.requiresConfirmation === true,
          })
        }
      }
    }
    const results: AgentToolResult[] = []
    let stopToolBatch = false
    for (const call of toolCalls) {
      // The user hit stop while an earlier tool was running: skip remaining tools,
      // but fill in paired error results to keep tool_use/tool_result pairs valid for the next request
      if (this.cancelled) {
        results.push({
          id: call.id,
          name: call.name,
          output: '(the user stopped the run; this tool was not executed)',
          isError: true,
        })
        continue
      }
      if (stopToolBatch) {
        results.push({
          id: call.id,
          name: call.name,
          output: '(a previous tool result stopped this tool batch; this tool was not executed)',
          isError: true,
        })
        continue
      }
      // Unusable input (truncated by the token limit, or JSON that failed to parse):
      // don't execute; feed a targeted error back so the model retries correctly
      if (call.truncated || call.inputError) {
        this.inputParseFails++
        const output = call.truncated
          ? 'Tool arguments were cut off by the output length limit; the tool was not executed. Split this operation into several smaller tool calls (less content per call) and try again.'
          : `Tool input JSON failed to parse; the tool was not executed: ${call.inputError}\nFix the arguments (make sure quotes inside strings are escaped) and call again.`
        results.push({ id: call.id, name: call.name, output, isError: true })
        events?.onToolExecuted?.({
          call,
          execution: { output, isError: true, summary: call.name },
        })
        continue
      }
      this.inputParseFails = 0
      events?.onToolStart?.(call)
      const snapshot = !this.mutationSeen ? captureSnapshot?.() : undefined
      let outcome: ToolExecutionOutcome
      try {
        outcome = await skill.executeTool(call, this.abortController?.signal)
      } catch (e) {
        outcome = {
          output: e instanceof Error ? e.message : String(e),
          isError: true,
          summary: call.name,
        }
      }
      if (generation !== this.generation) return // reset while a tool was running
      let execution: ToolExecution
      try {
        if (isToolExecutionSuspension(outcome)) {
          const suspended = await this.waitForSuspension(outcome, this.abortController?.signal)
          if (generation !== this.generation) return // reset while awaiting approval
          if (suspended === null) {
            results.push({
              id: call.id,
              name: call.name,
              output: '(the user stopped the run; this tool was not executed)',
              isError: true,
            })
            continue
          }
          execution = suspended
        } else if (
          typeof outcome === 'object' &&
          outcome !== null &&
          'kind' in outcome &&
          outcome.kind === 'tool-execution-suspension'
        ) {
          execution = INVALID_TOOL_OUTPUT
        } else {
          execution = outcome
        }
      } catch {
        execution = INVALID_TOOL_OUTPUT
      }
      if (!isFinalToolExecution(execution)) execution = INVALID_TOOL_OUTPUT
      let content: AgentToolContent[] | undefined
      try {
        content = boundedToolContent(execution.modelContent)
      } catch {
        execution = {
          output: 'invalid_tool_output',
          isError: true,
          mutated: false,
          summary: call.name,
        }
      }
      const firstMutation = !!execution.mutated && !this.mutationSeen
      if (execution.mutated) this.mutationSeen = true
      if (execution.mutated && this.presentationCorrectionPending) {
        this.presentationCorrectionPasses++
        this.presentationCorrectionPending = false
      }
      results.push({
        id: call.id,
        name: call.name,
        output: execution.output,
        isError: execution.isError,
        content,
      })
      events?.onToolExecuted?.({
        call,
        execution,
        snapshotBefore: firstMutation ? snapshot : undefined,
      })
      if (execution.stopToolBatch) stopToolBatch = true
    }
    this.history.push({ role: 'tool', results })

    // Cancelled while tools were executing: finish immediately, no further model request
    if (this.cancelled) {
      if (this.presentationContract && this.mutationSeen) {
        await this.finishPresentationRun()
      } else {
        this.running = false
        this.runUserMsg = null
        events?.onDone?.({ text: this.turnText, cancelled: true, turnLimit: false })
      }
      return
    }

    // Bad-input retries hit the cap: abort instead of burning more turns
    if (this.inputParseFails >= MAX_INPUT_PARSE_RETRIES) {
      this.running = false
      this.rollbackFailedRun()
      events?.onError?.(
        `Tool input was unusable (unparseable or truncated) ${MAX_INPUT_PARSE_RETRIES} times in a row; retries stopped, please send the request again`,
      )
      return
    }

    this.turns++
    if (this.options.maxTurns !== undefined && this.turns >= this.options.maxTurns) {
      // Don't throw away the context already gathered: append one no-tools turn for a partial answer
      this.finalizing = true
      this.history.push({ role: 'user', text: TURN_LIMIT_NOTE })
    }
    // Long runs (e.g. page-by-page generation) over budget mid-way: truncate stale tool outputs so each turn doesn't resend a huge payload
    this.squashStaleToolOutputs()
    events?.onTurnEnd?.()
    this.startTurn()
  }

  /** Returns true when the run was settled or redirected into a corrective turn. */
  private async finishPresentationRun(): Promise<boolean> {
    const contract = this.presentationContract
    const hooks = this.options.skill.presentation
    if (!contract || !hooks) return false
    const generation = this.generation
    const modelCorrectionPasses = this.presentationCorrectionPasses
    let completion: Awaited<ReturnType<typeof hooks.complete>>
    try {
      const reconciliationSignal = this.reconciliationController?.signal
      completion = await hooks.complete({
        contract,
        mutated: this.mutationSeen,
        cancelled: this.cancelled,
        correctionPasses: this.presentationCorrectionPasses,
        ...(reconciliationSignal ? { signal: reconciliationSignal } : {}),
      })
    } catch {
      if (generation !== this.generation || !this.running) return true
      this.failPresentationRun('presentation_completion_unavailable')
      return true
    }
    if (generation !== this.generation || !this.running) {
      if (completion?.kind === 'receipt') {
        try {
          const receipt = parsePresentationCompletionReceipt(completion.receipt, contract)
          if (receipt.correctionPasses < modelCorrectionPasses) throw new TypeError()
          this.options.events?.onAbandonedPresentationCompletion?.({
            documentToken: contract.documentToken,
            sessionToken: contract.sessionToken,
            receipt,
            facts: renderPresentationCompletionFacts(receipt, contract),
          })
        } catch {
          // An invalid abandoned receipt is never persisted or surfaced.
        }
      }
      return true
    }
    let completionKind: 'receipt' | 'correct'
    let completionValue: unknown
    try {
      if (completion.kind !== 'receipt' && completion.kind !== 'correct') throw new TypeError()
      completionKind = completion.kind
      completionValue = completion.kind === 'correct' ? completion.instruction : completion.receipt
    } catch {
      this.failPresentationRun('presentation_completion_invalid')
      return true
    }
    if (completionKind === 'correct') {
      let instruction: string | undefined
      try {
        instruction = this.boundedPresentationText(
          completionValue,
          FINAL_RESPONSE_CORRECTION_MAX_CHARS,
        )
      } catch {
        instruction = undefined
      }
      if (
        this.cancelled ||
        !instruction ||
        this.presentationCorrectionTurns >= contract.maxCorrectionPasses
      ) {
        this.failPresentationRun('presentation_receipt_required')
        return true
      }
      this.presentationCorrectionTurns++
      this.options.events?.onPresentationCorrection?.({
        pass: this.presentationCorrectionTurns,
        maximum: contract.maxCorrectionPasses,
      })
      this.presentationCorrectionPending = true
      this.history.push({
        role: 'assistant',
        text: this.turnText || COMPLETED_VIA_TOOLS_TEXT,
      })
      this.history.push({ role: 'user', text: instruction })
      this.options.events?.onText?.('')
      this.startTurn()
      return true
    }
    try {
      const receipt = parsePresentationCompletionReceipt(completionValue, contract)
      // A receipt may not under-report corrections already orchestrated here.
      if (receipt.correctionPasses < this.presentationCorrectionPasses)
        throw new TypeError('presentation correction count mismatch')
      const facts = renderPresentationCompletionFacts(receipt, contract)
      const localized = this.options.events?.onPresentationReceipt?.({ receipt, facts })
      const text = localized
        ? (this.boundedPresentationText(localized, FINAL_RESPONSE_CORRECTION_MAX_CHARS) ??
          renderPresentationCompletionText(facts))
        : renderPresentationCompletionText(facts)
      this.history.push({ role: 'assistant', text })
      this.running = false
      this.runUserMsg = null
      this.abortController = null
      this.reconciliationController = null
      // Replace any streamed model claim in the live bubble with receipt truth.
      this.options.events?.onText?.(text)
      this.options.events?.onDone?.({
        text,
        cancelled: this.cancelled,
        turnLimit: this.finalizing,
        presentation: facts,
        ...(this.turnStopReason === 'max_tokens' && !this.cancelled ? { truncated: true } : {}),
      })
    } catch {
      this.failPresentationRun('presentation_receipt_invalid')
    }
    return true
  }

  private failPresentationRun(error: string): void {
    // Preserve paired provider history and any mutation truth. Do not roll back
    // the run or allow free-form prose to become terminal success.
    this.history.push({ role: 'assistant', text: `presentation:error;code=${error}` })
    this.running = false
    this.runUserMsg = null
    this.abortController = null
    this.reconciliationController = null
    try {
      this.options.events?.onText?.(`presentation:error;code=${error}`)
    } catch {
      // A UI callback cannot suppress the authoritative failure event.
    }
    try {
      this.options.events?.onError?.(error)
    } catch {
      // Consumer callbacks remain outside the state machine trust boundary.
    }
  }

  private async waitForSuspension(
    suspension: ToolExecutionSuspension,
    signal?: AbortSignal,
  ): Promise<ToolExecution | null> {
    if (signal?.aborted) return null
    let removeAbortListener: (() => void) | undefined
    const aborted = new Promise<null>((resolve) => {
      if (!signal) return
      const onAbort = () => resolve(null)
      signal.addEventListener('abort', onAbort, { once: true })
      removeAbortListener = () => signal.removeEventListener('abort', onAbort)
    })
    try {
      const settled = new Promise<ToolExecution>((resolve) => {
        try {
          Promise.prototype.then.call(
            suspension.result,
            (execution) =>
              resolve(isFinalToolExecution(execution) ? execution : INVALID_TOOL_OUTPUT),
            () => resolve(INVALID_TOOL_OUTPUT),
          )
        } catch {
          resolve(INVALID_TOOL_OUTPUT)
        }
      })
      const result = await Promise.race([settled, aborted])
      return result
    } finally {
      removeAbortListener?.()
    }
  }
}

/**
 * Redact secret-looking tokens from an outgoing user message so accidentally
 * pasted API keys, URL credentials, and password assignments don't reach
 * remote model APIs verbatim.
 *
 * Imported from public PR #32 (BuiltByHarshil), with the credential pattern
 * narrowed to URL userinfo (scheme://user:pass@host) so ordinary "a:b@c"
 * prose is never rewritten.
 */
export function sanitizeAgentPayload(payload: string): string {
  return payload
    .replace(/\b(?:sk-|AIza|ghp_|secret_)[A-Za-z0-9_-]{16,}/g, '[REDACTED_API_KEY]')
    .replace(/([a-z][a-z0-9+.-]*:\/\/[^\s:@/]+):[^\s@/]+@/gi, '$1:[REDACTED_CREDENTIALS]@')
    .replace(
      /(password|passwd|secret_key|private_key)(\s*[:=]\s*)["'][^"']+["']/gi,
      '$1$2"[REDACTED_SECURE_TOKEN]"',
    )
}
