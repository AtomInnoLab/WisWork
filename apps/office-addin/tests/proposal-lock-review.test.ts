import { expect, it, vi } from 'vitest'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'

const request = () => ({
  operation: 'edit',
  title: 'Edit',
  preview: {},
  impact: { host: 'powerpoint', targets: ['host'], count: 1 },
  fingerprint: 'v1',
  validate: async () => true,
  execute: vi.fn(async () => {}),
})

it('late reviews cannot replace a new proposal or revive a cancelled approval', async () => {
  const completions: ((value: { state: 'ready'; token: string; pages: [] }) => void)[] = []
  const signals: AbortSignal[] = []
  const controller = createStructuredProposalController(undefined, {
    review: (_proposal, signal) => {
      signals.push(signal)
      return new Promise((resolve) => {
        completions.push(resolve)
      })
    },
    beforeWrite: async () => {},
    afterWrite: () => {},
  })
  controller.propose(request())
  await Promise.resolve()
  const second = controller.propose(request())
  await Promise.resolve()
  expect(signals[0]!.aborted).toBe(true)
  completions[0]!({ state: 'ready', token: 'old', pages: [] })
  await Promise.resolve()
  expect(controller.pending()).toMatchObject({ id: second.id, lockReview: { state: 'checking' } })
  controller.newTurn()
  completions[1]!({ state: 'ready', token: 'late', pages: [] })
  await Promise.resolve()
  expect(controller.pending()).toBeUndefined()
})
it('holds confirmation until the exact pending proposal has completed its native lock review', async () => {
  let complete!: (value: unknown) => void
  const review = vi.fn(
    () =>
      new Promise((resolve) => {
        complete = resolve
      }),
  )
  const controller = createStructuredProposalController(undefined, {
    review,
    beforeWrite: async () => {},
    afterWrite: () => {},
  } as never)
  const first = controller.propose(request())
  expect(controller.pending()).toMatchObject({ lockReview: { state: 'checking' } })
  await expect(controller.confirm(first.id)).rejects.toThrow('presentation_lock_review_pending')
  const input = request()
  controller.propose(input)
  await Promise.resolve()
  complete({
    state: 'ready',
    pages: [{ projectId: 'p', pageId: 'page', title: 'Locked', slideIds: ['host'] }],
    token: 'exact',
  })
  await vi.waitFor(() =>
    expect(controller.pending()).toMatchObject({ lockReview: { state: 'ready' } }),
  )
  await controller.confirm(controller.pending()!.id)
  expect(input.execute).toHaveBeenCalledOnce()
})
it('a failed lock review leaves a rejectable proposal and cannot write or expose private errors', async () => {
  const input = request()
  const controller = createStructuredProposalController(undefined, {
    review: async () => {
      throw new Error('/private/path token')
    },
    beforeWrite: async () => {},
    afterWrite: () => {},
  } as never)
  const proposal = controller.propose(input)
  await vi.waitFor(() =>
    expect(controller.pending()).toMatchObject({ lockReview: { state: 'unavailable' } }),
  )
  await expect(controller.confirm(proposal.id)).rejects.toThrow(
    'presentation_lock_review_unavailable',
  )
  expect(input.execute).not.toHaveBeenCalled()
  const decision = controller.waitForDecision(proposal.id)
  controller.reject()
  await expect(decision).resolves.toEqual({ status: 'rejected' })
})
