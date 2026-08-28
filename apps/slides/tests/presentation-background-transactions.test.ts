import { describe, expect, it, vi } from 'vitest'
import { executePreparedBackgroundFamilyTransaction } from '../src/renderer/ai/presentation-background-transactions'

const fp = (char: string) => `sha256:${char.repeat(64)}`

describe('background family transaction preparation', () => {
  it('prepares every slide and executes one fingerprint-bound atomic transaction', async () => {
    const api = {
      preparePresentationTarget: vi
        .fn()
        .mockResolvedValueOnce({
          status: 'prepared',
          expectedDeckRevision: fp('a'),
          target: { slideId: 'ppt/slides/slide1.xml', expectedFingerprint: fp('b') },
        })
        .mockResolvedValueOnce({
          status: 'prepared',
          expectedDeckRevision: fp('a'),
          target: { slideId: 'ppt/slides/slide2.xml', expectedFingerprint: fp('c') },
        }),
      executePresentationTransaction: vi.fn(async (transaction) => ({
        status: 'applied' as const,
        transactionId: transaction.transactionId,
        resultingDeckRevision: fp('d'),
        operationCount: transaction.operations.length,
      })),
      cancelPresentationTransaction: vi.fn(async () => true),
    }
    const result = await executePreparedBackgroundFamilyTransaction(api as any, {
      transactionId: 'slides-background-test',
      backgrounds: [
        { slideIndex: 0, color: '#112233' },
        { slideIndex: 1, color: '#AABBCC' },
      ],
    })
    expect(api.executePresentationTransaction).toHaveBeenCalledWith({
      transactionId: 'slides-background-test',
      expectedDeckRevision: fp('a'),
      mode: 'atomic',
      operations: [
        {
          kind: 'set_slide_background',
          clientId: 'background-1',
          target: { slideId: 'ppt/slides/slide1.xml', expectedFingerprint: fp('b') },
          color: '#112233',
        },
        {
          kind: 'set_slide_background',
          clientId: 'background-2',
          target: { slideId: 'ppt/slides/slide2.xml', expectedFingerprint: fp('c') },
          color: '#AABBCC',
        },
      ],
    })
    expect(result.receipt).toMatchObject({ status: 'applied', operationCount: 2 })
  })

  it('aborts before apply when slide preparations observe different deck revisions', async () => {
    const api = {
      preparePresentationTarget: vi
        .fn()
        .mockResolvedValueOnce({
          status: 'prepared',
          expectedDeckRevision: fp('a'),
          target: { slideId: 'slide-1', expectedFingerprint: fp('b') },
        })
        .mockResolvedValueOnce({
          status: 'prepared',
          expectedDeckRevision: fp('c'),
          target: { slideId: 'slide-2', expectedFingerprint: fp('d') },
        }),
      executePresentationTransaction: vi.fn(),
      cancelPresentationTransaction: vi.fn(async () => true),
    }
    const result = await executePreparedBackgroundFamilyTransaction(api as any, {
      transactionId: 'slides-background-conflict',
      backgrounds: [
        { slideIndex: 0, color: '#112233' },
        { slideIndex: 1, color: '#112233' },
      ],
    })
    expect(result.receipt).toMatchObject({ status: 'conflict', code: 'target_stale' })
    expect(api.executePresentationTransaction).not.toHaveBeenCalled()
    expect(api.cancelPresentationTransaction).toHaveBeenCalledOnce()
  })

  it('rejects an over-limit all-slide request before preparing any target', async () => {
    const api = {
      preparePresentationTarget: vi.fn(),
      executePresentationTransaction: vi.fn(),
      cancelPresentationTransaction: vi.fn(),
    }
    const result = await executePreparedBackgroundFamilyTransaction(api as any, {
      transactionId: 'slides-background-over-limit',
      backgrounds: Array.from({ length: 51 }, (_, slideIndex) => ({
        slideIndex,
        color: '#112233',
      })),
    })
    expect(result.receipt).toMatchObject({ status: 'unchanged', code: 'write_not_applied' })
    expect(api.preparePresentationTarget).not.toHaveBeenCalled()
    expect(api.executePresentationTransaction).not.toHaveBeenCalled()
  })
})
