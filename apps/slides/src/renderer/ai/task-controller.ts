import type {
  PresentationAcceptanceContract,
  PresentationCompletionReceipt,
} from '@wiswork/presentation-verification'
import { parsePresentationCompletionReceipt } from '@wiswork/presentation-verification'
import type { PresentationVerificationFlags } from '@wiswork/presentation-verification'
import {
  emitPresentationTelemetry,
  presentationCompletionTelemetry,
  type PresentationTelemetryEvent,
} from '@wiswork/presentation-verification'
import type { AgentToolCall, PresentationTaskHooks } from '@wiswork/agent-core'
import { runSlidesTaskReview, type SlidesTaskReviewAdapter } from './task-review'
import {
  compileSlidesAcceptance,
  type SlidesAcceptanceAuthority,
  type SlidesAcceptanceIntent,
} from './task-acceptance'

export type SlidesTaskEnrollment = {
  contract: PresentationAcceptanceContract
  plannedMutationTargets: string[]
  plan?: string[]
}

export type SlidesRecordedMutation = {
  transactionId: string
  status: 'applied' | 'unchanged' | 'conflict' | 'uncertain'
  mutatedTargetTokens: string[]
  rollbackId?: string
}

export interface SlidesTaskController {
  hooks: PresentationTaskHooks
  recordMutation(receipt: SlidesRecordedMutation): void
  reset(): void
}

const canonicalProperties = {
  color: 'color',
  fontSize: 'font_size',
  fontFamily: 'font_family',
  bold: 'bold',
  italic: 'italic',
  x: 'x',
  y: 'y',
  w: 'width',
  h: 'height',
} as const

/** Compiles only exact canonical calls. Scripts and unsupported families bypass verification. */
export function compileCanonicalSlidesCalls(input: {
  calls: readonly AgentToolCall[]
  authority: SlidesAcceptanceAuthority
  sourceTargetTokens: Readonly<Record<string, string>>
  taskId: string
}): SlidesTaskEnrollment | { kind: 'bypass' } | { kind: 'clarify'; question: string } {
  const affected = new Set<number>()
  const changes: SlidesAcceptanceIntent['changes'] = []
  for (const call of input.calls) {
    const slideIndex = call.input.slideIndex
    // An invalid model-generated index is not missing user intent. Let the
    // transactional tool reject it and give the model a chance to self-correct;
    // asking the user produced an unanswerable generic clarification state.
    if (!Number.isSafeInteger(slideIndex) || (slideIndex as number) < 0) return { kind: 'bypass' }
    const slide = (slideIndex as number) + 1
    if (!input.authority.slides.some((candidate) => candidate.number === slide))
      return { kind: 'bypass' }
    affected.add(slide)
    if (call.name === 'set_slide_background') {
      const color = call.input.color
      if (typeof color !== 'string') return { kind: 'bypass' }
      changes.push({ kind: 'set_background', slides: [slide], color })
      continue
    }
    const sourceId = call.input.sourceId
    if (typeof sourceId !== 'string') return { kind: 'bypass' }
    const targetToken = input.sourceTargetTokens[`${slide}:${sourceId}`]
    // A newly-created runtime element may not have reached the durable authority map yet.
    // This is an internal verification-availability condition, not missing user intent.
    // Let the existing transactional tool path validate the source id instead of presenting
    // an unanswerable "more information needed" state to the user.
    if (!targetToken) return { kind: 'bypass' }
    if (call.name === 'set_element_style' || call.name === 'set_element_transform') {
      let added = false
      for (const [field, property] of Object.entries(canonicalProperties)) {
        if (!Object.hasOwn(call.input, field)) continue
        changes.push({
          kind: 'set_property',
          slides: [slide],
          targetToken,
          property,
          value: call.input[field] as string | number | boolean,
        })
        added = true
      }
      if (!added) return { kind: 'bypass' }
      continue
    }
    if (call.name === 'set_element_fill' || call.name === 'set_element_stroke') {
      const color = call.input.color
      if (typeof color !== 'string' || color === 'none') return { kind: 'bypass' }
      changes.push({
        kind: 'set_property',
        slides: [slide],
        targetToken,
        property: call.name === 'set_element_fill' ? 'fill_color' : 'stroke_color',
        value: color,
      })
      continue
    }
    return { kind: 'bypass' }
  }
  if (!changes.length || !affected.size) return { kind: 'bypass' }
  const compiled = compileSlidesAcceptance(
    {
      taskId: input.taskId,
      affectedSlides: [...affected],
      changes,
      maxCorrectionPasses: 2,
    },
    input.authority,
  )
  if (compiled.status === 'compiled')
    return {
      ...compiled,
      plannedMutationTargets: [...new Set(compiled.plannedMutationTargets)],
      ...(input.calls.length > 1 || affected.size > 1
        ? { plan: ['presentation_apply_bounded_edits', 'presentation_verify_postconditions'] }
        : {}),
    }
  if (compiled.status === 'unchanged') return { kind: 'bypass' }
  return { kind: 'clarify', question: compiled.code }
}

export function createSlidesTaskController(deps: {
  enroll(
    calls: readonly AgentToolCall[],
    currentContract: PresentationAcceptanceContract | undefined,
    signal?: AbortSignal,
  ): Promise<SlidesTaskEnrollment | { kind: 'bypass' } | { kind: 'clarify'; question: string }>
  reviewAdapter: SlidesTaskReviewAdapter
  /** Test seam only: proves that reconciliation never replays the original batch. */
  executeOriginal?: () => unknown
  afterTaskReview?: (receipt: PresentationCompletionReceipt) => void | Promise<void>
  flags?: PresentationVerificationFlags
  telemetry?: (event: PresentationTelemetryEvent) => void
  onHostCorrection?: (pass: number) => void | Promise<void>
}): SlidesTaskController {
  type RunState = {
    generation: number
    enrollment: SlidesTaskEnrollment
    mutations: SlidesRecordedMutation[]
  }
  let generation = 0
  let activeTaskId: string | undefined
  const runs = new Map<string, RunState>()

  const reset = () => {
    generation += 1
    activeTaskId = undefined
  }

  const hooks: PresentationTaskHooks = {
    abandon: reset,
    prepare: () => ({ kind: 'bypass' }),
    enroll: async (calls, currentContract, signal) => {
      const runGeneration = ++generation
      const result = await deps.enroll(calls, currentContract, signal)
      if ('kind' in result) return result
      const state: RunState = { generation: runGeneration, enrollment: result, mutations: [] }
      runs.set(result.contract.taskId, state)
      while (runs.size > 8) {
        const oldest = runs.keys().next().value as string | undefined
        if (!oldest) break
        runs.delete(oldest)
      }
      if (runGeneration === generation) activeTaskId = result.contract.taskId
      emitPresentationTelemetry(deps.telemetry, {
        host: 'pc',
        phase: 'plan',
        outcome: 'success',
        code: 'ready',
        count: result.contract.affectedSlides.length,
        durationMs: 0,
      })
      return {
        kind: 'ready',
        contract: result.contract,
        ...(result.plan ? { plan: result.plan } : {}),
      }
    },
    complete: async ({ contract, mutated, cancelled, signal }) => {
      const startedAt = Date.now()
      const state = runs.get(contract.taskId)
      const applied = [...(state?.mutations ?? [])].filter(({ status }) => status === 'applied')
      let receipt: PresentationCompletionReceipt
      if (!state) {
        receipt = parsePresentationCompletionReceipt(
          {
            version: 1,
            taskId: contract.taskId,
            status: 'failed',
            mutationReceiptIds: [],
            passedCheckIds: [],
            failedCheckIds: contract.checks.map(({ id }) => id),
            unavailableCheckIds: [],
            correctionPasses: 0,
            affectedSlides: contract.affectedSlides,
            safeCode: 'mutation_failed',
          },
          contract,
        )
      } else if (applied.length === 0) {
        receipt = parsePresentationCompletionReceipt(
          {
            version: 1,
            taskId: contract.taskId,
            status: cancelled && !mutated ? 'unchanged' : 'failed',
            mutationReceiptIds: [],
            passedCheckIds: cancelled && !mutated ? contract.checks.map(({ id }) => id) : [],
            failedCheckIds: cancelled && !mutated ? [] : contract.checks.map(({ id }) => id),
            unavailableCheckIds: [],
            correctionPasses: 0,
            affectedSlides: contract.affectedSlides,
            safeCode:
              cancelled && !mutated
                ? 'cancelled'
                : mutated
                  ? 'office_state_uncertain'
                  : 'mutation_failed',
          },
          contract,
        )
      } else {
        const covered = new Set(applied.flatMap(({ mutatedTargetTokens }) => mutatedTargetTokens))
        if (
          covered.size !== state.enrollment.plannedMutationTargets.length ||
          state.enrollment.plannedMutationTargets.some((target) => !covered.has(target))
        ) {
          receipt = parsePresentationCompletionReceipt(
            {
              version: 1,
              taskId: contract.taskId,
              status: 'applied_unverified',
              mutationReceiptIds: applied.map(({ transactionId }) => transactionId),
              passedCheckIds: [],
              failedCheckIds: [],
              unavailableCheckIds: contract.checks.map(({ id }) => id),
              correctionPasses: 0,
              affectedSlides: contract.affectedSlides,
              safeCode: 'office_state_uncertain',
            },
            contract,
          )
        } else {
          const rollbackId = [...applied].reverse().find((item) => item.rollbackId)?.rollbackId
          receipt = await runSlidesTaskReview({
            contract,
            initialMutationReceiptIds: applied.map(({ transactionId }) => transactionId),
            plannedMutationTargets: state.enrollment.plannedMutationTargets,
            ...(rollbackId ? { rollbackId } : {}),
            adapter: deps.reviewAdapter,
            isCurrent: () => state.generation === generation && activeTaskId === contract.taskId,
            signal,
            flags: deps.flags,
            telemetry: deps.telemetry,
            onHostCorrection: deps.onHostCorrection,
          })
        }
      }
      if (state && state.generation === generation && activeTaskId === contract.taskId) {
        await deps.afterTaskReview?.(receipt)
        activeTaskId = undefined
      }
      runs.delete(contract.taskId)
      emitPresentationTelemetry(
        deps.telemetry,
        presentationCompletionTelemetry('pc', receipt, Date.now() - startedAt),
      )
      return { kind: 'receipt', receipt }
    },
  }

  return {
    hooks,
    recordMutation(receipt) {
      if (!activeTaskId) return
      const state = runs.get(activeTaskId)
      if (!state || state.generation !== generation) return
      if (
        receipt.status === 'applied' &&
        receipt.mutatedTargetTokens.some(
          (target) => !state.enrollment.plannedMutationTargets.includes(target),
        )
      ) {
        state.mutations.push({ ...receipt, status: 'uncertain' })
        return
      }
      state.mutations.push(receipt)
      emitPresentationTelemetry(deps.telemetry, {
        host: 'pc',
        phase: 'dispatch',
        outcome:
          receipt.status === 'applied'
            ? 'success'
            : receipt.status === 'unchanged'
              ? 'unchanged'
              : 'failed',
        code:
          receipt.status === 'applied'
            ? 'ready'
            : receipt.status === 'unchanged'
              ? 'unchanged'
              : 'failed',
        count: receipt.mutatedTargetTokens.length,
        durationMs: 0,
      })
    },
    reset,
  }
}
