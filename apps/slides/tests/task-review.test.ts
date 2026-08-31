import { describe, expect, it, vi } from 'vitest'
import type {
  PresentationAcceptanceContract,
  VisualReviewResult,
} from '@wiswork/presentation-verification'
import { runSlidesTaskReview, type SlidesTaskReviewAdapter } from '../src/renderer/ai/task-review'

const digest = `sha256:${'a'.repeat(64)}`
const revision = `sha256:${'b'.repeat(64)}`
const contract = (maxCorrectionPasses: 0 | 1 | 2 = 2): PresentationAcceptanceContract => ({
  version: 1,
  taskId: 'task-1',
  documentToken: 'doc-1',
  sessionToken: 'session-1',
  baseRevision: digest,
  affectedSlides: [2],
  referenceSlides: [1],
  checks: [
    {
      id: 'color',
      kind: 'element_property',
      slide: 2,
      roleOrTarget: { kind: 'target', targetToken: 'title-2' },
      property: 'color',
      expected: '#112233',
    },
  ],
  maxCorrectionPasses,
})

const review = (status: VisualReviewResult['status']): VisualReviewResult =>
  status === 'pass'
    ? { status, failedCheckIds: [], observations: [], fixIntents: [] }
    : status === 'cannot_verify'
      ? {
          status,
          failedCheckIds: ['color'],
          observations: [
            { code: 'review_unavailable', severity: 'error', checkId: 'color', slide: 2 },
          ],
          fixIntents: [],
        }
      : {
          status,
          failedCheckIds: ['color'],
          observations: [
            { code: 'reference_mismatch', severity: 'warning', checkId: 'color', slide: 2 },
          ],
          fixIntents: [
            {
              checkId: 'color',
              kind: 'set_property',
              roleOrTarget: { kind: 'target', targetToken: 'title-2' },
              property: 'color',
              value: '#112233',
            },
          ],
        }

function adapter(reviews: VisualReviewResult[] = [review('pass')]): SlidesTaskReviewAdapter {
  let pass = 0
  return {
    refresh: vi.fn(async (lineage) => ({
      documentToken: 'doc-1',
      sessionToken: 'session-1',
      revision,
      leaseToken: `lease-${pass}`,
      ...lineage,
    })),
    verifyDeterministic: vi.fn(async () => [{ checkId: 'color', status: 'pass' as const }]),
    capture: vi.fn(async ({ slide, role, authority }) => ({
      slide,
      role,
      mediaToken: `${authority.leaseToken}-${slide}`,
      bytes: 100,
      revision: authority.revision,
      leaseToken: authority.leaseToken,
      sessionToken: authority.sessionToken,
    })),
    review: vi.fn(async () => reviews[Math.min(pass++, reviews.length - 1)]!),
    correct: vi.fn(async () => ({
      mutationReceiptId: `fix-${pass}`,
      rollbackId: 'undo-1',
      correctedCheckIds: ['color'],
    })),
    isCurrent: vi.fn(() => true),
  }
}

describe('PC task-specific rendered verification', () => {
  it('returns verified only after revision-bound deterministic and visual pass', async () => {
    const a = adapter()
    const result = await runSlidesTaskReview({
      contract: contract(),
      initialMutationReceiptIds: ['edit-1'],
      rollbackId: 'undo-1',
      adapter: a,
    })
    expect(result).toMatchObject({
      status: 'verified',
      passedCheckIds: ['color'],
      correctionPasses: 0,
      mutationReceiptIds: ['edit-1'],
    })
    expect(a.capture).toHaveBeenCalledTimes(2)
    expect(a.refresh).toHaveBeenCalledWith(
      { baseRevision: digest, mutationReceiptIds: ['edit-1'] },
      undefined,
    )
    expect(a.verifyDeterministic).toHaveBeenCalledWith(
      expect.anything(),
      expect.anything(),
      ['title-2'],
      undefined,
    )
  })

  it('always clears ephemeral screenshot state when review fails', async () => {
    const a = adapter()
    const cleanup = vi.fn()
    a.cleanup = cleanup
    vi.mocked(a.review).mockRejectedValueOnce(new Error('offline'))
    await runSlidesTaskReview({
      contract: contract(),
      initialMutationReceiptIds: ['edit-1'],
      adapter: a,
    })
    expect(cleanup).toHaveBeenCalledOnce()
  })

  it('fails closed on screenshot revision mismatch', async () => {
    const a = adapter()
    vi.mocked(a.capture).mockResolvedValueOnce({
      slide: 2,
      role: 'affected',
      mediaToken: 'shot',
      bytes: 100,
      revision: digest,
      leaseToken: 'lease-0',
      sessionToken: 'session-1',
    })
    const result = await runSlidesTaskReview({
      contract: contract(),
      initialMutationReceiptIds: ['edit-1'],
      adapter: a,
    })
    expect(result).toMatchObject({
      status: 'applied_unverified',
      safeCode: 'screenshot_unavailable',
      unavailableCheckIds: ['color'],
    })
    expect(a.review).not.toHaveBeenCalled()
  })

  it('never treats missing deterministic evidence as a pass', async () => {
    const a = adapter()
    vi.mocked(a.verifyDeterministic).mockResolvedValue([])
    const result = await runSlidesTaskReview({
      contract: contract(),
      initialMutationReceiptIds: ['edit-1'],
      adapter: a,
    })
    expect(result).toMatchObject({ status: 'applied_unverified', unavailableCheckIds: ['color'] })
    expect(a.capture).not.toHaveBeenCalled()
  })

  it('applies one canonical fix and re-verifies', async () => {
    const a = adapter([review('needs_fix'), review('pass')])
    const result = await runSlidesTaskReview({
      contract: contract(),
      initialMutationReceiptIds: ['edit-1'],
      adapter: a,
    })
    expect(result).toMatchObject({
      status: 'verified',
      correctionPasses: 1,
      mutationReceiptIds: ['edit-1', 'fix-1'],
    })
    expect(a.correct).toHaveBeenCalledTimes(1)
    expect(a.refresh).toHaveBeenCalledTimes(2)
  })

  it('stops after the contract maximum of two corrections', async () => {
    const a = adapter([review('needs_fix')])
    const result = await runSlidesTaskReview({
      contract: contract(2),
      initialMutationReceiptIds: ['edit-1'],
      adapter: a,
    })
    expect(result).toMatchObject({
      status: 'needs_user',
      correctionPasses: 2,
      failedCheckIds: ['color'],
      safeCode: 'confirmation_required',
    })
    expect(a.correct).toHaveBeenCalledTimes(2)
  })

  it('does not reexecute after screenshot rejection following an applied fix', async () => {
    const a = adapter([review('needs_fix')])
    vi.mocked(a.capture).mockImplementation(async ({ slide, role, authority }) =>
      authority.leaseToken === 'lease-0'
        ? {
            slide,
            role,
            mediaToken: `shot-${slide}`,
            bytes: 100,
            revision: authority.revision,
            leaseToken: authority.leaseToken,
            sessionToken: authority.sessionToken,
          }
        : null,
    )
    const result = await runSlidesTaskReview({
      contract: contract(),
      initialMutationReceiptIds: ['edit-1'],
      adapter: a,
    })
    expect(result).toMatchObject({
      status: 'applied_unverified',
      correctionPasses: 1,
      mutationReceiptIds: ['edit-1', 'fix-1'],
    })
    expect(a.correct).toHaveBeenCalledTimes(1)
  })

  it('stops as applied-unverified when correction evidence expands beyond failed checks', async () => {
    const a = adapter([review('needs_fix')])
    vi.mocked(a.correct).mockResolvedValue({
      mutationReceiptId: 'fix-unsafe',
      correctedCheckIds: ['color', 'other'],
    })
    const result = await runSlidesTaskReview({
      contract: contract(),
      initialMutationReceiptIds: ['edit-1'],
      adapter: a,
    })
    expect(result).toMatchObject({
      status: 'applied_unverified',
      correctionPasses: 1,
      mutationReceiptIds: ['edit-1', 'fix-unsafe'],
    })
    expect(a.correct).toHaveBeenCalledTimes(1)
    expect(a.refresh).toHaveBeenCalledTimes(1)
  })

  it('rejects unsafe or out-of-scope visual fixes without mutation', async () => {
    const a = adapter([
      {
        ...review('needs_fix'),
        fixIntents: [
          {
            checkId: 'color',
            kind: 'set_property',
            roleOrTarget: { kind: 'target', targetToken: 'other' },
            property: 'color',
            value: '#112233',
          },
        ],
      },
    ])
    const result = await runSlidesTaskReview({
      contract: contract(),
      initialMutationReceiptIds: ['edit-1'],
      adapter: a,
    })
    expect(result).toMatchObject({
      status: 'needs_user',
      correctionPasses: 0,
      safeCode: 'unsupported_check',
    })
    expect(a.correct).not.toHaveBeenCalled()
  })

  it('reconciles cancellation and session switches without erasing mutation truth', async () => {
    const a = adapter()
    vi.mocked(a.isCurrent).mockReturnValue(false)
    const result = await runSlidesTaskReview({
      contract: contract(),
      initialMutationReceiptIds: ['edit-1'],
      adapter: a,
    })
    expect(result).toMatchObject({
      status: 'applied_unverified',
      safeCode: 'stale_authority',
    })
    expect(a.capture).not.toHaveBeenCalled()
  })

  it('preserves rollback/history truth from applied corrections', async () => {
    const a = adapter([review('needs_fix'), review('pass')])
    const result = await runSlidesTaskReview({
      contract: contract(),
      initialMutationReceiptIds: ['edit-1'],
      rollbackId: 'undo-original',
      adapter: a,
    })
    expect(result.rollbackId).toBe('undo-original')
  })

  it('enforces the eight screenshot and 2 MiB visual bounds', async () => {
    const many = { ...contract(), affectedSlides: [1, 2, 3, 4, 5, 6, 7, 8], referenceSlides: [9] }
    const a = adapter()
    const bounded = await runSlidesTaskReview({
      contract: many,
      initialMutationReceiptIds: ['edit-1'],
      adapter: a,
    })
    expect(bounded).toMatchObject({ status: 'applied_unverified' })
    expect(a.review).not.toHaveBeenCalled()

    const b = adapter()
    vi.mocked(b.capture).mockImplementation(async ({ slide, role, authority }) => ({
      slide,
      role,
      mediaToken: `shot-${slide}`,
      bytes: 2 * 1024 * 1024,
      revision: authority.revision,
      leaseToken: authority.leaseToken,
      sessionToken: authority.sessionToken,
    }))
    const oversized = await runSlidesTaskReview({
      contract: contract(),
      initialMutationReceiptIds: ['edit-1'],
      adapter: b,
    })
    expect(oversized).toMatchObject({ status: 'applied_unverified' })
    expect(b.review).not.toHaveBeenCalled()
  })

  it('returns unchanged rather than verified when no mutation occurred', async () => {
    const a = adapter()
    const result = await runSlidesTaskReview({
      contract: contract(),
      initialMutationReceiptIds: [],
      adapter: a,
    })
    expect(result).toMatchObject({ status: 'unchanged', mutationReceiptIds: [] })
  })

  it('requires exact lease and session proof on every screenshot', async () => {
    const a = adapter()
    vi.mocked(a.capture).mockImplementation(async ({ slide, role, authority }) => ({
      slide,
      role,
      mediaToken: `shot-${slide}`,
      bytes: 100,
      revision: authority.revision,
      leaseToken: 'wrong-lease',
      sessionToken: authority.sessionToken,
    }))
    const result = await runSlidesTaskReview({
      contract: contract(),
      initialMutationReceiptIds: ['edit-1'],
      adapter: a,
    })
    expect(result).toMatchObject({
      status: 'applied_unverified',
      safeCode: 'screenshot_unavailable',
    })
    expect(a.review).not.toHaveBeenCalled()
  })

  it('distinguishes cancellation before and after dispatch', async () => {
    const before = adapter()
    const beforeSignal = new AbortController()
    beforeSignal.abort()
    await expect(
      runSlidesTaskReview({
        contract: contract(),
        initialMutationReceiptIds: [],
        adapter: before,
        signal: beforeSignal.signal,
      }),
    ).resolves.toMatchObject({ status: 'unchanged', safeCode: 'cancelled' })

    const after = adapter()
    const afterSignal = new AbortController()
    afterSignal.abort()
    await expect(
      runSlidesTaskReview({
        contract: contract(),
        initialMutationReceiptIds: ['edit-1'],
        adapter: after,
        signal: afterSignal.signal,
      }),
    ).resolves.toMatchObject({ status: 'applied_unverified', safeCode: 'cancelled_after_apply' })
  })
})
