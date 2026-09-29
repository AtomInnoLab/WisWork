import { expect, vi } from 'vitest'
import type {
  ProposalController,
  StructuredProposalController,
} from '../src/agent/proposal-controller.js'

/** Model a user clicking after the approval UI has finished its read-only review. */
export async function confirmReviewed(
  controller: ProposalController | StructuredProposalController,
  id: string,
): Promise<void> {
  await vi.waitFor(() => {
    const pending = controller.pending()
    expect(pending && 'lockReview' in pending ? pending.lockReview?.state : undefined).not.toBe(
      'checking',
    )
  })
  return controller.confirm(id)
}
