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
} from '@wiswork/presentation-verification'
import type { SlidesDeterministicResult } from './task-acceptance'

export type SlidesTaskReviewAuthority = {
  documentToken: string
  sessionToken: string
  revision: string
  leaseToken: string
}

export type SlidesTaskScreenshot = {
  slide: number
  role: 'affected' | 'reference'
  mediaToken: string
  bytes: number
  /** Capture-side authority proof. Omission means the adapter binds it to its input lease. */
  revision?: string
}

export interface SlidesTaskReviewAdapter {
  /** Refreshes the renderer, then returns a new main-process authority lease. */
  refresh(signal?: AbortSignal): Promise<SlidesTaskReviewAuthority>
  verifyDeterministic(
    contract: PresentationAcceptanceContract,
    authority: SlidesTaskReviewAuthority,
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
  ): Promise<{ mutationReceiptId: string; rollbackId?: string; correctedCheckIds: string[] }>
  isCurrent(authority: SlidesTaskReviewAuthority): boolean
}

const unavailable = (
  contract: PresentationAcceptanceContract,
  receipts: string[],
  correctionPasses: number,
  rollbackId?: string,
): PresentationCompletionReceipt =>
  parsePresentationCompletionReceipt(
    {
      version: 1,
      taskId: contract.taskId,
      status: receipts.length ? 'applied_unverified' : 'failed',
      mutationReceiptIds: receipts,
      passedCheckIds: [],
      failedCheckIds: receipts.length ? [] : contract.checks.map(({ id }) => id),
      unavailableCheckIds: receipts.length ? contract.checks.map(({ id }) => id) : [],
      correctionPasses,
      affectedSlides: contract.affectedSlides,
      ...(rollbackId && receipts.length ? { rollbackId } : {}),
      safeCode: receipts.length ? 'screenshot_unavailable' : 'stale_authority',
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
  rollbackId?: string
  adapter: SlidesTaskReviewAdapter
  signal?: AbortSignal
}): Promise<PresentationCompletionReceipt> {
  const contract = parsePresentationAcceptanceContract(input.contract)
  const contractDigest = await digestPresentationAcceptanceContract(contract)
  const receipts = [...input.initialMutationReceiptIds]
  let rollbackId = input.rollbackId
  let correctionPasses = 0
  try {
    for (;;) {
      input.signal?.throwIfAborted()
      const authority = await input.adapter.refresh(input.signal)
      if (
        authority.documentToken !== contract.documentToken ||
        authority.sessionToken !== contract.sessionToken ||
        !input.adapter.isCurrent(authority)
      )
        return unavailable(contract, receipts, correctionPasses, rollbackId)
      const deterministic = await input.adapter.verifyDeterministic(
        contract,
        authority,
        input.signal,
      )
      const expectedCheckIds = contract.checks.map(({ id }) => id).sort()
      const deterministicCheckIds = deterministic.map(({ checkId }) => checkId).sort()
      if (
        deterministicCheckIds.length !== expectedCheckIds.length ||
        new Set(deterministicCheckIds).size !== deterministicCheckIds.length ||
        deterministicCheckIds.some((checkId, index) => checkId !== expectedCheckIds[index])
      )
        return unavailable(contract, receipts, correctionPasses, rollbackId)
      if (!input.adapter.isCurrent(authority))
        return unavailable(contract, receipts, correctionPasses, rollbackId)

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
          (shot.revision !== undefined && shot.revision !== authority.revision) ||
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
          screenshots: screenshots.map(({ revision: _revision, ...shot }) => shot),
          deterministicResults: deterministic,
        })
      } catch {
        return unavailable(contract, receipts, correctionPasses, rollbackId)
      }
      const visual = parseVisualReviewResult(await input.adapter.review(facts, input.signal))
      if (!input.adapter.isCurrent(authority))
        return unavailable(contract, receipts, correctionPasses, rollbackId)
      if (visual.status === 'cannot_verify')
        return unavailable(contract, receipts, correctionPasses, rollbackId)
      if (visual.status === 'pass' && deterministic.every(({ status }) => status === 'pass'))
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
      const correction = await input.adapter.correct(visual.fixIntents, authority, input.signal)
      receipts.push(correction.mutationReceiptId)
      correctionPasses += 1
      if (correction.rollbackId) rollbackId = correction.rollbackId
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
    return unavailable(contract, receipts, correctionPasses, rollbackId)
  }
}
