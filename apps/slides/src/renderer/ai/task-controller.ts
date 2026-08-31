import type {
  PresentationAcceptanceContract,
  PresentationCompletionReceipt,
} from '@wiswork/presentation-verification'
import { parsePresentationCompletionReceipt } from '@wiswork/presentation-verification'
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
    if (!Number.isSafeInteger(slideIndex) || (slideIndex as number) < 0)
      return { kind: 'clarify', question: 'presentation_scope_required' }
    const slide = (slideIndex as number) + 1
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
    if (!targetToken) return { kind: 'clarify', question: 'presentation_target_required' }
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
  if (compiled.status === 'compiled') return compiled
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
}): SlidesTaskController {
  let enrollment: SlidesTaskEnrollment | undefined
  const mutations: SlidesRecordedMutation[] = []

  const reset = () => {
    enrollment = undefined
    mutations.length = 0
  }

  const hooks: PresentationTaskHooks = {
    prepare: () => ({ kind: 'bypass' }),
    enroll: async (calls, currentContract, signal) => {
      const result = await deps.enroll(calls, currentContract, signal)
      if ('kind' in result) return result
      enrollment = result
      mutations.length = 0
      return {
        kind: 'ready',
        contract: result.contract,
        ...(result.plan ? { plan: result.plan } : {}),
      }
    },
    complete: async ({ contract, mutated, cancelled, signal }) => {
      const applied = mutations.filter(({ status }) => status === 'applied')
      let receipt: PresentationCompletionReceipt
      if (!enrollment || enrollment.contract.taskId !== contract.taskId) {
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
        const rollbackId = [...applied].reverse().find((item) => item.rollbackId)?.rollbackId
        receipt = await runSlidesTaskReview({
          contract,
          initialMutationReceiptIds: applied.map(({ transactionId }) => transactionId),
          ...(rollbackId ? { rollbackId } : {}),
          adapter: deps.reviewAdapter,
          signal,
        })
      }
      await deps.afterTaskReview?.(receipt)
      reset()
      return { kind: 'receipt', receipt }
    },
  }

  return {
    hooks,
    recordMutation(receipt) {
      if (!enrollment) return
      if (
        receipt.status === 'applied' &&
        (receipt.mutatedTargetTokens.length !== enrollment.plannedMutationTargets.length ||
          receipt.mutatedTargetTokens.some(
            (target) => !enrollment!.plannedMutationTargets.includes(target),
          ))
      ) {
        mutations.push({ ...receipt, status: 'uncertain' })
        return
      }
      mutations.push(receipt)
    },
    reset,
  }
}
