import type { AgentToolCall, AgentToolDef, ToolExecutionOutcome } from './types'
import type {
  PresentationAcceptanceContract,
  PresentationCompletionReceipt,
} from '@wiswork/presentation-verification'

export type PresentationTaskPreparation =
  | { kind: 'bypass' }
  | { kind: 'clarify'; question: string }
  | {
      kind: 'ready'
      contract: PresentationAcceptanceContract
      plan?: string[]
      requiresConfirmation?: boolean
    }

export type PresentationTaskCompletion =
  | { kind: 'receipt'; receipt: PresentationCompletionReceipt }
  | { kind: 'correct'; instruction: string }

export interface PresentationTaskHooks {
  /** Host-owned intent compiler. Model output is never accepted as a contract. */
  prepare(
    instruction: string,
    signal?: AbortSignal,
  ): PresentationTaskPreparation | Promise<PresentationTaskPreparation>
  /** Required when prepare marks an existing host confirmation boundary. */
  confirm?(
    contract: PresentationAcceptanceContract,
    signal?: AbortSignal,
  ): boolean | Promise<boolean>
  /** Reconcile authoritative host state and either close with a receipt or request one correction. */
  complete(context: {
    contract: PresentationAcceptanceContract
    mutated: boolean
    cancelled: boolean
    correctionPasses: number
    signal?: AbortSignal
  }): PresentationTaskCompletion | Promise<PresentationTaskCompletion>
}

/**
 * Deliberately small set of run facts available to a terminal-response policy.
 * The hook never receives tool results, snapshots, or the full conversation.
 */
export interface FinalResponseReviewContext {
  /** text from the normal tool-free terminal turn */
  readonly text: string
  /** whether any tool reported a successful mutation during this run */
  readonly mutated: boolean
}

/**
 * A skill packages one capability domain for the agent loop: its system
 * prompt section, its tools, per-turn context, and the tool executor.
 * AI Docs ships a docx skill; Excel / PPT skills plug in the same way.
 */
export interface AgentSkill {
  id: string
  /** system prompt section describing this skill's rules and tools */
  systemPrompt: string
  tools: AgentToolDef[]
  /**
   * Fresh context sections attached to every user turn (e.g. document
   * skeleton + selection). Return '' when there is nothing to attach.
   */
  buildContext?(): string
  /** Optional host-owned orchestration for presentation mutation runs. */
  presentation?: PresentationTaskHooks
  /**
   * Optionally reject one normal tool-free terminal response by returning a
   * static corrective user message. Exceptions fail open in AgentLoop.
   */
  reviewFinalResponse?(context: FinalResponseReviewContext): string | undefined
  /**
   * signal: aborted when the user hits stop. Long-running tools (e.g.
   * generate_deck with internal LLM calls) should check signal.aborted in
   * their loops and stop promptly.
   */
  executeTool(
    call: AgentToolCall,
    signal?: AbortSignal,
  ): ToolExecutionOutcome | Promise<ToolExecutionOutcome>
}

/**
 * Merge several skills into one (tool names must be globally unique).
 * `intro` becomes the shared preamble of the combined system prompt.
 */
export function composeSkills(id: string, intro: string, skills: AgentSkill[]): AgentSkill {
  const owner = new Map<string, AgentSkill>()
  for (const skill of skills) {
    for (const tool of skill.tools) {
      if (owner.has(tool.name)) throw new Error(`duplicate tool name: ${tool.name}`)
      owner.set(tool.name, skill)
    }
  }
  const reviewers = skills.flatMap((skill) =>
    skill.reviewFinalResponse ? [skill.reviewFinalResponse.bind(skill)] : [],
  )
  const presentations = skills.flatMap((skill) => (skill.presentation ? [skill.presentation] : []))
  if (presentations.length > 1) throw new Error('multiple presentation orchestrators')
  return {
    id,
    systemPrompt: [intro, ...skills.map((s) => s.systemPrompt)].filter(Boolean).join('\n\n'),
    tools: skills.flatMap((s) => s.tools),
    buildContext: () =>
      skills
        .map((s) => s.buildContext?.() ?? '')
        .filter(Boolean)
        .join('\n\n'),
    ...(presentations[0] ? { presentation: presentations[0] } : {}),
    ...(reviewers.length
      ? {
          reviewFinalResponse: (context: FinalResponseReviewContext) => {
            for (const review of reviewers) {
              const correction = review(context)
              if (correction) return correction
            }
            return undefined
          },
        }
      : {}),
    executeTool: (call, signal) => {
      const skill = owner.get(call.name)
      if (!skill) {
        return { output: `Unknown tool: ${call.name}`, isError: true, summary: call.name }
      }
      return skill.executeTool(call, signal)
    },
  }
}
