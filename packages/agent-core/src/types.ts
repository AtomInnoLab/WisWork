/** JSON-Schema described tool exposed to the model */
export interface AgentToolDef {
  name: string
  description: string
  /** JSON Schema (object) describing the tool input */
  inputSchema: Record<string, unknown>
}

export interface AgentToolCall {
  id: string
  name: string
  input: Record<string, unknown>
  /** Host-assigned identity for one actual tool invocation; transport retries reuse it. */
  invocationId?: string | undefined
  /** Parse error when the model emitted invalid input JSON; the loop feeds back an is_error result for retry instead of aborting the run */
  inputError?: string | undefined
  /** The argument stream was cut off by the token limit (stop_reason max_tokens); the loop asks the model to split the call instead of "fixing JSON" */
  truncated?: boolean | undefined
}

export interface AgentToolResult {
  id: string
  /** tool name (Gemini addresses function responses by name, not id) */
  name: string
  output: string
  isError?: boolean | undefined
  /** bounded multimodal blocks returned to the model with this tool result */
  content?: AgentToolContent[] | undefined
}

/** inline image attached to a user turn, fed to vision-capable providers as multimodal input */
export interface AgentImage {
  /** raw base64 (no data: URL prefix) */
  base64: string
  /** e.g. "image/png" */
  mime: string
}

export type AgentToolContent = { type: 'image'; image: AgentImage }

export type AgentMessage =
  | { role: 'user'; text: string; images?: AgentImage[] | undefined }
  | { role: 'assistant'; text: string; toolCalls?: AgentToolCall[] | undefined }
  | { role: 'tool'; results: AgentToolResult[] }

/**
 * Side-channel display data: UI-only, never merged into messages sent to the LLM.
 * kind='images' → image grid; kind='links' → link list; kind='text' → extra text.
 */
export interface ToolDisplay {
  kind: 'images' | 'links' | 'text'
  /** entry list for images / links modes */
  items?: Array<{ url: string; title?: string; thumb?: string }>
  /** extra text for text mode */
  text?: string
}

/** outcome of one tool execution */
export interface ToolExecution {
  /** result text fed back to the model */
  output: string
  isError?: boolean
  /** true when the tool changed the underlying artifact (document / sheet / deck) */
  mutated?: boolean
  /**
   * Stop executing later calls emitted in the same provider tool batch. The
   * loop still pairs every skipped call with an error result before returning
   * the completed batch to the provider.
   */
  stopToolBatch?: boolean
  /** short human-readable label for activity UI */
  summary: string
  /**
   * In-memory diagnostic cause. Never serialized into AgentMessage/tool output;
   * consumers may extract only allowlisted identifiers from it.
   */
  diagnosticError?: unknown
  /** bounded content sent to the model; unlike display, this enters model history */
  modelContent?: AgentToolContent[]
  /**
   * Side-channel display: for UI only, never enters the LLM context.
   * Ignored when tool results are assembled into an AgentMessage.
   */
  display?: ToolDisplay
}

/**
 * A tool can suspend its final execution while it waits for an external
 * decision (for example, user approval). The loop keeps the current tool call
 * open and does not make another provider request until `result` settles.
 */
export interface ToolExecutionSuspension extends ToolExecution {
  kind: 'tool-execution-suspension'
  result: Promise<ToolExecution>
}

export type ToolExecutionOutcome = ToolExecution | ToolExecutionSuspension

const toolExecutionSuspensions = new WeakSet<object>()
const suspensionOwners = new WeakMap<object, object>()
const legacySuspensionOwner = Object.freeze({})

function mintToolExecutionSuspension(
  owner: object,
  result: Promise<ToolExecution>,
): ToolExecutionSuspension {
  const suspension: ToolExecutionSuspension = {
    kind: 'tool-execution-suspension',
    result,
    output: 'tool_execution_suspended',
    summary: 'Awaiting tool execution',
  }
  toolExecutionSuspensions.add(suspension)
  suspensionOwners.set(suspension, owner)
  return Object.freeze(suspension)
}

/** Create the only supported, bounded shape for a suspended tool execution. */
export function suspendToolExecution(result: Promise<ToolExecution>): ToolExecutionSuspension {
  // The placeholder ToolExecution fields preserve source compatibility for
  // direct skill consumers. AgentLoop detects `kind` and never publishes this
  // placeholder to model history or execution events.
  return mintToolExecutionSuspension(legacySuspensionOwner, result)
}

/** @internal AgentLoop is the only production caller. */
export function mintLoopToolExecutionSuspension(
  owner: object,
  result: Promise<ToolExecution>,
): ToolExecutionSuspension {
  return mintToolExecutionSuspension(owner, result)
}

/** Owner-specific authority for transports that must not trust the public compatibility helper. */
export function createToolExecutionSuspensionAuthority(): Readonly<{
  suspend(result: Promise<ToolExecution>): ToolExecutionSuspension
  owns(value: ToolExecutionOutcome): value is ToolExecutionSuspension
}> {
  const owner = Object.freeze({})
  return Object.freeze({
    suspend: (result) => mintToolExecutionSuspension(owner, result),
    owns: (value): value is ToolExecutionSuspension =>
      typeof value === 'object' && value !== null && suspensionOwners.get(value) === owner,
  })
}

/** Accept only suspension objects created by suspendToolExecution in this Agent Core instance. */
export function isToolExecutionSuspension(
  value: ToolExecutionOutcome,
): value is ToolExecutionSuspension {
  return typeof value === 'object' && value !== null && toolExecutionSuspensions.has(value)
}

// ---- run phase (drives the in-progress status line in chat UIs) ----

export type AgentPhaseKind =
  /** request sent, waiting for the model's first content block */
  | 'requesting'
  | 'thinking'
  | 'responding'
  /** the model is streaming tool arguments (e.g. a full outline) with no visible text */
  | 'tool-input'
  | 'tool-running'

export interface AgentPhase {
  kind: AgentPhaseKind
  toolName?: string | undefined
}

// ---- LLM transport (how one model turn is streamed; app supplies the impl) ----

export interface AgentStreamRequest {
  system: string
  messages: AgentMessage[]
  tools: AgentToolDef[]
}

export interface AgentStreamCallbacks {
  onDelta(text: string): void
  /** complete parsed tool call (arguments finished streaming) */
  onToolCall(call: AgentToolCall): void
  /** Phase changes within the model stream (thinking / responding / tool-input); older transports may omit this */
  onPhase?(phase: AgentPhase): void
  /** normalized stop reason of the turn ('max_tokens' = cut off by the token limit); transports may omit this */
  onStopReason?(reason: string): void
  onDone(): void
  onError(error: string): void
}

export interface AgentStreamHandle {
  /** abort the in-flight turn; the transport must still emit onDone afterwards */
  cancel(): void
}

export interface AgentTransport {
  stream(request: AgentStreamRequest, callbacks: AgentStreamCallbacks): AgentStreamHandle
}
