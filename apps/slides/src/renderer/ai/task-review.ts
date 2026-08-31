import {
  parsePresentationAcceptanceContract,
  parsePresentationCompletionReceipt,
  parsePresentationRenderingFacts,
  parseVisualReviewResult,
  digestPresentationAcceptanceContract,
  type PresentationAcceptanceContract,
  type PresentationCompletionReceipt,
  type PresentationRenderingFacts,
  type SafeFixIntent,
  type VisualReviewResult,
  PRESENTATION_VERIFICATION_LIMITS,
  type PresentationVerificationFlags,
  type PresentationTelemetryEvent,
  emitPresentationTelemetry,
} from '@wiswork/presentation-verification'
import type { SlidesDeterministicResult } from './task-acceptance'
import { AgentLoop, type AgentImage, type AgentTransport } from '@wiswork/agent-core'

export type SlidesTaskReviewAuthority = {
  documentToken: string
  sessionToken: string
  revision: string
  leaseToken: string
  baseRevision?: string
  mutationReceiptIds?: string[]
  rollbackId?: string
  taskId?: string
}

export type SlidesTaskScreenshot = {
  slide: number
  role: 'affected' | 'reference'
  mediaToken: string
  bytes: number
  revision: string
  leaseToken: string
  sessionToken: string
}

export interface SlidesTaskReviewAdapter {
  /** Refreshes the renderer, then returns a new main-process authority lease. */
  refresh(
    input: {
      taskId: string
      baseRevision: string
      mutationReceiptIds: string[]
      isCurrent(): boolean
    },
    signal?: AbortSignal,
  ): Promise<SlidesTaskReviewAuthority>
  verifyDeterministic(
    contract: PresentationAcceptanceContract,
    authority: SlidesTaskReviewAuthority,
    plannedMutationTargets: string[],
    signal?: AbortSignal,
  ): Promise<SlidesDeterministicResult[]>
  capture(input: {
    slide: number
    role: 'affected' | 'reference'
    authority: SlidesTaskReviewAuthority
    signal?: AbortSignal
  }): Promise<SlidesTaskScreenshot | null>
  /** Receives frozen facts only; it has no host, IPC, or mutation capability. */
  review(facts: PresentationRenderingFacts, signal?: AbortSignal): Promise<VisualReviewResult>
  correct(
    intents: SafeFixIntent[],
    authority: SlidesTaskReviewAuthority,
    signal?: AbortSignal,
  ): Promise<{
    mutationReceiptId: string
    applied?: boolean
    rollbackId?: string
    correctedCheckIds: string[]
  }>
  isCurrent(authority: SlidesTaskReviewAuthority): boolean
  cleanup?(taskId: string): void
}

/** Strict, read-only reviewer. Images are ephemeral and are never copied into facts or receipts. */
export function reviewSlidesRendering(input: {
  facts: PresentationRenderingFacts
  images: AgentImage[]
  transport: AgentTransport
  signal?: AbortSignal
}): Promise<VisualReviewResult> {
  return new Promise((resolve, reject) => {
    let settled = false
    const finish = (result: VisualReviewResult | Error) => {
      if (settled) return
      settled = true
      input.signal?.removeEventListener('abort', abort)
      if (result instanceof Error) reject(result)
      else resolve(result)
    }
    const loop = new AgentLoop({
      transport: input.transport,
      maxTurns: 1,
      skill: {
        id: 'slides-task-review',
        systemPrompt:
          'You are a read-only presentation acceptance reviewer. Return ONLY strict JSON: {"status":"pass|needs_fix|cannot_verify","failedCheckIds":[],"observations":[],"fixIntents":[]}. Use only check ids and safe schema values supplied in the facts. Never invent targets or request tools.',
        tools: [],
        executeTool: () => ({ output: 'read_only', summary: 'read only', isError: true }),
      },
      events: {
        onDone: ({ text, cancelled }) => {
          if (cancelled) return finish(new Error('cancelled'))
          try {
            const json = text.slice(text.indexOf('{'), text.lastIndexOf('}') + 1)
            finish(parseVisualReviewResult(JSON.parse(json)))
          } catch (error) {
            finish(error instanceof Error ? error : new Error('review_invalid'))
          }
        },
        onError: (error) => finish(new Error(error)),
      },
    })
    const abort = () => loop.cancel()
    input.signal?.addEventListener('abort', abort, { once: true })
    loop.run(JSON.stringify(input.facts), input.images)
  })
}

const unavailable = (
  contract: PresentationAcceptanceContract,
  receipts: string[],
  correctionPasses: number,
  rollbackId?: string,
  code:
    | 'screenshot_unavailable'
    | 'review_unavailable'
    | 'verification_invalid'
    | 'stale_authority'
    | 'cancelled'
    | 'cancelled_after_apply'
    | 'visual_disabled' = 'screenshot_unavailable',
): PresentationCompletionReceipt =>
  parsePresentationCompletionReceipt(
    {
      version: 1,
      taskId: contract.taskId,
      status: receipts.length
        ? 'applied_unverified'
        : code === 'cancelled'
          ? 'unchanged'
          : 'failed',
      mutationReceiptIds: receipts,
      passedCheckIds:
        code === 'cancelled' && !receipts.length ? contract.checks.map(({ id }) => id) : [],
      failedCheckIds:
        receipts.length || code === 'cancelled' ? [] : contract.checks.map(({ id }) => id),
      unavailableCheckIds: receipts.length ? contract.checks.map(({ id }) => id) : [],
      correctionPasses,
      affectedSlides: contract.affectedSlides,
      ...(rollbackId && receipts.length ? { rollbackId } : {}),
      safeCode: code,
    },
    contract,
  )

const sameTarget = (a: SafeFixIntent['roleOrTarget'], b: SafeFixIntent['roleOrTarget']) =>
  a.kind === b.kind &&
  (a.kind === 'target'
    ? b.kind === 'target' && a.targetToken === b.targetToken
    : b.kind === 'role' && a.role === b.role)

function fixesAreFrozen(
  contract: PresentationAcceptanceContract,
  review: VisualReviewResult,
): boolean {
  const failed = new Set(review.failedCheckIds)
  return review.fixIntents.every((intent) => {
    if (!failed.has(intent.checkId)) return false
    const check = contract.checks.find(({ id }) => id === intent.checkId)
    return (
      check?.kind === 'element_property' &&
      check.property === intent.property &&
      sameTarget(check.roleOrTarget, intent.roleOrTarget) &&
      check.expected === intent.value
    )
  })
}

function accountedReceipt(input: {
  contract: PresentationAcceptanceContract
  status: 'verified' | 'needs_user'
  results: SlidesDeterministicResult[]
  visualFailed: string[]
  receipts: string[]
  correctionPasses: number
  rollbackId?: string
  safeCode?: 'confirmation_required' | 'unsupported_check'
}): PresentationCompletionReceipt {
  const failed = new Set([
    ...input.results.filter((result) => result.status === 'fail').map(({ checkId }) => checkId),
    ...input.visualFailed,
  ])
  const unavailableIds = new Set(
    input.results.filter((result) => result.status === 'unavailable').map(({ checkId }) => checkId),
  )
  const passed = input.contract.checks
    .map(({ id }) => id)
    .filter((id) => !failed.has(id) && !unavailableIds.has(id))
  return parsePresentationCompletionReceipt(
    {
      version: 1,
      taskId: input.contract.taskId,
      status: input.status,
      mutationReceiptIds: input.receipts,
      passedCheckIds: passed,
      failedCheckIds: [...failed].filter((id) => !unavailableIds.has(id)),
      unavailableCheckIds: [...unavailableIds],
      correctionPasses: input.correctionPasses,
      affectedSlides: input.contract.affectedSlides,
      ...(input.rollbackId ? { rollbackId: input.rollbackId } : {}),
      ...(input.safeCode ? { safeCode: input.safeCode } : {}),
    },
    input.contract,
  )
}

export async function runSlidesTaskReview(input: {
  contract: PresentationAcceptanceContract
  initialMutationReceiptIds: string[]
  plannedMutationTargets?: string[]
  isCurrent?: () => boolean
  rollbackId?: string
  adapter: SlidesTaskReviewAdapter
  signal?: AbortSignal
  flags?: Pick<PresentationVerificationFlags, 'visualReview' | 'autoCorrection'>
  telemetry?: (event: PresentationTelemetryEvent) => void
  onHostCorrection?: (pass: number) => void | Promise<void>
}): Promise<PresentationCompletionReceipt> {
  const contract = parsePresentationAcceptanceContract(input.contract)
  const contractDigest = await digestPresentationAcceptanceContract(contract)
  const receipts = [...input.initialMutationReceiptIds]
  let rollbackId = input.rollbackId
  let correctionPasses = 0
  try {
    for (;;) {
      input.signal?.throwIfAborted()
      if (input.isCurrent && !input.isCurrent())
        return unavailable(
          contract,
          receipts,
          correctionPasses,
          rollbackId,
          receipts.length ? 'cancelled_after_apply' : 'cancelled',
        )
      const authority = await input.adapter.refresh(
        {
          taskId: contract.taskId,
          baseRevision: contract.baseRevision,
          mutationReceiptIds: receipts,
          isCurrent: input.isCurrent ?? (() => true),
        },
        input.signal,
      )
      if (!rollbackId && authority.rollbackId) rollbackId = authority.rollbackId
      if (
        authority.documentToken !== contract.documentToken ||
        authority.sessionToken !== contract.sessionToken ||
        authority.baseRevision !== contract.baseRevision ||
        authority.taskId !== contract.taskId ||
        JSON.stringify(authority.mutationReceiptIds) !== JSON.stringify(receipts) ||
        !input.adapter.isCurrent(authority)
      )
        return unavailable(contract, receipts, correctionPasses, rollbackId, 'stale_authority')
      const deterministic = await input.adapter.verifyDeterministic(
        contract,
        authority,
        input.plannedMutationTargets ??
          contract.checks.flatMap((check) =>
            check.kind === 'element_property' && check.roleOrTarget.kind === 'target'
              ? [check.roleOrTarget.targetToken]
              : [],
          ),
        input.signal,
      )
      emitPresentationTelemetry(input.telemetry, {
        host: 'pc',
        phase: 'deterministic',
        outcome: 'success',
        code: 'ready',
        count: deterministic.length,
        durationMs: 0,
      })
      const expectedCheckIds = contract.checks.map(({ id }) => id).sort()
      const deterministicCheckIds = deterministic.map(({ checkId }) => checkId).sort()
      if (
        deterministicCheckIds.length !== expectedCheckIds.length ||
        new Set(deterministicCheckIds).size !== deterministicCheckIds.length ||
        deterministicCheckIds.some((checkId, index) => checkId !== expectedCheckIds[index])
      )
        return unavailable(contract, receipts, correctionPasses, rollbackId, 'verification_invalid')
      if (!input.adapter.isCurrent(authority))
        return unavailable(contract, receipts, correctionPasses, rollbackId)

      if (input.flags?.visualReview === false) {
        const failed = deterministic.filter(({ status }) => status !== 'pass')
        if (!failed.length && !receipts.length)
          return parsePresentationCompletionReceipt(
            {
              version: 1,
              taskId: contract.taskId,
              status: 'unchanged',
              mutationReceiptIds: [],
              passedCheckIds: contract.checks.map(({ id }) => id),
              failedCheckIds: [],
              unavailableCheckIds: [],
              correctionPasses: 0,
              affectedSlides: contract.affectedSlides,
            },
            contract,
          )
        if (!failed.length)
          return unavailable(contract, receipts, correctionPasses, rollbackId, 'visual_disabled')
        return accountedReceipt({
          contract,
          status: 'needs_user',
          results: deterministic,
          visualFailed: [],
          receipts,
          correctionPasses,
          ...(rollbackId ? { rollbackId } : {}),
          safeCode: 'unsupported_check',
        })
      }
      const pages = [
        ...contract.affectedSlides.map((slide) => ({ slide, role: 'affected' as const })),
        ...contract.referenceSlides.map((slide) => ({ slide, role: 'reference' as const })),
      ]
      if (pages.length > 8) return unavailable(contract, receipts, correctionPasses, rollbackId)
      const screenshots: SlidesTaskScreenshot[] = []
      for (const page of pages) {
        const shot = await input.adapter.capture({ ...page, authority, signal: input.signal })
        if (
          !shot ||
          shot.slide !== page.slide ||
          shot.role !== page.role ||
          shot.revision !== authority.revision ||
          shot.leaseToken !== authority.leaseToken ||
          shot.sessionToken !== authority.sessionToken ||
          !input.adapter.isCurrent(authority)
        )
          return unavailable(contract, receipts, correctionPasses, rollbackId)
        screenshots.push(shot)
      }
      let facts: PresentationRenderingFacts
      try {
        facts = parsePresentationRenderingFacts({
          contractDigest,
          revision: authority.revision,
          screenshots: screenshots.map(
            ({
              revision: _revision,
              leaseToken: _leaseToken,
              sessionToken: _sessionToken,
              ...shot
            }) => shot,
          ),
          deterministicResults: deterministic,
        })
      } catch {
        return unavailable(contract, receipts, correctionPasses, rollbackId)
      }
      let visual: VisualReviewResult
      try {
        const rawReview = await input.adapter.review(facts, input.signal)
        visual = parseVisualReviewResult(rawReview)
        emitPresentationTelemetry(input.telemetry, {
          host: 'pc',
          phase: 'visual',
          outcome:
            visual.status === 'pass'
              ? 'success'
              : visual.status === 'needs_fix'
                ? 'needs_user'
                : 'unverified',
          code:
            visual.status === 'pass'
              ? 'verified'
              : visual.status === 'needs_fix'
                ? 'needs_user'
                : 'review_unavailable',
          count: visual.failedCheckIds.length,
          durationMs: 0,
        })
      } catch {
        return unavailable(contract, receipts, correctionPasses, rollbackId, 'review_unavailable')
      }
      if (!input.adapter.isCurrent(authority))
        return unavailable(contract, receipts, correctionPasses, rollbackId)
      if (visual.status === 'cannot_verify')
        return unavailable(contract, receipts, correctionPasses, rollbackId, 'review_unavailable')
      if (visual.status === 'pass' && deterministic.every(({ status }) => status === 'pass'))
        if (receipts.length === 0)
          return parsePresentationCompletionReceipt(
            {
              version: 1,
              taskId: contract.taskId,
              status: 'unchanged',
              mutationReceiptIds: [],
              passedCheckIds: contract.checks.map(({ id }) => id),
              failedCheckIds: [],
              unavailableCheckIds: [],
              correctionPasses: 0,
              affectedSlides: contract.affectedSlides,
            },
            contract,
          )
        else
          return accountedReceipt({
            contract,
            status: 'verified',
            results: deterministic,
            visualFailed: [],
            receipts,
            correctionPasses,
            ...(rollbackId ? { rollbackId } : {}),
          })
      if (visual.status !== 'needs_fix' || !fixesAreFrozen(contract, visual))
        return accountedReceipt({
          contract,
          status: 'needs_user',
          results: deterministic,
          visualFailed: visual.failedCheckIds,
          receipts,
          correctionPasses,
          ...(rollbackId ? { rollbackId } : {}),
          safeCode: 'unsupported_check',
        })
      if (input.flags?.autoCorrection === false)
        return accountedReceipt({
          contract,
          status: 'needs_user',
          results: deterministic,
          visualFailed: visual.failedCheckIds,
          receipts,
          correctionPasses,
          ...(rollbackId ? { rollbackId } : {}),
          safeCode: 'confirmation_required',
        })
      if (correctionPasses >= contract.maxCorrectionPasses)
        return accountedReceipt({
          contract,
          status: 'needs_user',
          results: deterministic,
          visualFailed: visual.failedCheckIds,
          receipts,
          correctionPasses,
          ...(rollbackId ? { rollbackId } : {}),
          safeCode: 'confirmation_required',
        })
      if (receipts.length >= PRESENTATION_VERIFICATION_LIMITS.maxReceiptIds)
        return accountedReceipt({
          contract,
          status: 'needs_user',
          results: deterministic,
          visualFailed: visual.failedCheckIds,
          receipts,
          correctionPasses,
          ...(rollbackId ? { rollbackId } : {}),
          safeCode: 'confirmation_required',
        })
      await input.onHostCorrection?.(correctionPasses + 1)
      input.signal?.throwIfAborted()
      const correction = await input.adapter.correct(visual.fixIntents, authority, input.signal)
      if (correction.applied === false)
        return accountedReceipt({
          contract,
          status: 'needs_user',
          results: deterministic,
          visualFailed: visual.failedCheckIds,
          receipts,
          correctionPasses,
          ...(rollbackId ? { rollbackId } : {}),
          safeCode: 'confirmation_required',
        })
      receipts.push(correction.mutationReceiptId)
      emitPresentationTelemetry(input.telemetry, {
        host: 'pc',
        phase: 'correction',
        outcome: 'success',
        code: 'ready',
        count: correction.correctedCheckIds.length,
        durationMs: 0,
      })
      correctionPasses += 1
      if (!rollbackId && correction.rollbackId) rollbackId = correction.rollbackId
      const expectedCorrected = [...new Set(visual.fixIntents.map(({ checkId }) => checkId))].sort()
      const provedCorrected = [...new Set(correction.correctedCheckIds)].sort()
      if (
        provedCorrected.length !== correction.correctedCheckIds.length ||
        provedCorrected.length !== expectedCorrected.length ||
        provedCorrected.some((checkId, index) => checkId !== expectedCorrected[index])
      )
        return unavailable(contract, receipts, correctionPasses, rollbackId)
      // Never reuse pre-correction screenshots or authority. The next iteration refreshes both.
    }
  } catch {
    const cancelled = input.signal?.aborted === true
    return unavailable(
      contract,
      receipts,
      correctionPasses,
      rollbackId,
      cancelled ? (receipts.length ? 'cancelled_after_apply' : 'cancelled') : 'review_unavailable',
    )
  } finally {
    input.adapter.cleanup?.(contract.taskId)
  }
}
