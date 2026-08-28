/**
 * Post-transaction quality review helpers:
 *  - generatedPageRange / mergeQcPages: which pages a landing marks for QC (incl. insert_at shifting)
 *  - createSlideFixSkill: read-only visual reviewer with no executor tools
 */
import { describe, it, expect, vi } from 'vitest'
import {
  generatedPageRange,
  mergeQcPages,
  createSlideFixSkill,
  isQcEnabled,
  qcSlidePage,
  captureCurrentQcShot,
  toVisualQualityReceipt,
  publishAppliedDeterministicQuality,
} from '../src/renderer/ai/slide-qc'
import type { DeckAccess } from '../src/renderer/ai/slides-skill'

const access: DeckAccess = {
  getSlides: () => [],
  getCurrent: () => 0,
  getSelectedIds: () => [],
  applySlide: () => {},
  applyDeck: () => {},
  fitWidthPx: 1280,
}

describe('visual quality receipts', () => {
  it('publishes one deterministic receipt for one applied AiPanel transaction', async () => {
    const published = vi.fn()
    const completedKeys = new Set<string>()
    const one: DeckAccess = {
      ...access,
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] } as never],
    }
    const run = () =>
      publishAppliedDeterministicQuality({
        transactionId: 'tx-1',
        receiptStatus: 'applied',
        sessionId: 'session-1',
        pageIndexes: [0, 0],
        completedKeys,
        access: one,
        prepareSlide: async () => ({ status: 'prepared', slideId: 'ppt/slides/slide1.xml' }),
        publish: published,
        isCurrent: () => true,
      })
    await expect(run()).resolves.toEqual([0])
    await expect(run()).resolves.toEqual([])
    expect(published).toHaveBeenCalledOnce()
    expect(published.mock.calls[0]![0]).toMatchObject({
      transactionId: 'tx-1',
      slideId: 'ppt/slides/slide1.xml',
      source: 'deterministic',
      status: 'available',
    })
  })

  it('publishes nothing when the AiPanel session becomes stale during durable target resolution', async () => {
    let resolvePrepare!: (value: { status: string; slideId?: string }) => void
    let current = true
    const published = vi.fn()
    const one: DeckAccess = {
      ...access,
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] } as never],
    }
    const pending = publishAppliedDeterministicQuality({
      transactionId: 'tx-stale',
      receiptStatus: 'applied',
      sessionId: 'session-1',
      pageIndexes: [0],
      completedKeys: new Set(),
      access: one,
      prepareSlide: () => new Promise((resolve) => (resolvePrepare = resolve)),
      publish: published,
      isCurrent: () => current,
    })
    current = false
    resolvePrepare({ status: 'prepared', slideId: 'ppt/slides/slide1.xml' })
    await expect(pending).resolves.toEqual([])
    expect(published).not.toHaveBeenCalled()
  })

  it('never copies model text into the shared receipt', () => {
    const receipt = toVisualQualityReceipt('qc-1', 'tx-1', 'slide-1', {
      ok: true,
      edited: false,
      reply: 'PRIVATE quoted content',
      preIssues: 0,
      postIssues: 0,
    })
    expect(receipt).toMatchObject({ status: 'available', findings: [{ code: 'visual_quality' }] })
    expect(JSON.stringify(receipt)).not.toContain('PRIVATE')
  })

  it('keeps transport failure independent as quality unavailable', () => {
    expect(
      toVisualQualityReceipt('qc-1', 'tx-1', 'slide-1', {
        ok: true,
        edited: false,
        reply: '',
        preIssues: 0,
        postIssues: 0,
        error: 'offline details',
      }),
    ).toEqual({
      qualityRunId: 'qc-1',
      transactionId: 'tx-1',
      slideId: 'slide-1',
      source: 'visual',
      status: 'unavailable',
      code: 'transport_unavailable',
    })
  })
})

describe('generatedPageRange', () => {
  it('replace covers the whole deck', () => {
    expect(generatedPageRange('replace', { pages: 3 })).toEqual([0, 1, 2])
  })

  it('append covers only the new tail', () => {
    expect(generatedPageRange('append', { pages: 5, appendedFrom: 3 })).toEqual([3, 4])
  })

  it('replace_at / insert_at cover the single touched page', () => {
    expect(generatedPageRange('replace_at', { pages: 5, insertedIndex: 2 })).toEqual([2])
    expect(generatedPageRange('insert_at', { pages: 5, insertedIndex: 0 })).toEqual([0])
  })

  it('missing insertedIndex yields nothing', () => {
    expect(generatedPageRange('insert_at', { pages: 5 })).toEqual([])
  })
})

describe('mergeQcPages', () => {
  it('replace discards earlier pendings', () => {
    expect(mergeQcPages([7, 8], 'replace', { pages: 2 })).toEqual([0, 1])
  })

  it('append unions and dedupes', () => {
    expect(mergeQcPages([1, 3], 'append', { pages: 5, appendedFrom: 3 })).toEqual([1, 3, 4])
  })

  it('insert_at shifts pendings at/after the insertion point', () => {
    expect(mergeQcPages([1, 3], 'insert_at', { pages: 5, insertedIndex: 2 })).toEqual([1, 2, 4])
  })

  it('replace_at adds the page without shifting', () => {
    expect(mergeQcPages([1], 'replace_at', { pages: 5, insertedIndex: 3 })).toEqual([1, 3])
  })
})

describe('createSlideFixSkill', () => {
  it('is a read-only visual reviewer with no mutation tools', () => {
    const skill = createSlideFixSkill(access)
    expect(skill.tools).toEqual([])
  })

  it('rejects tool execution at the quality boundary', async () => {
    const one: DeckAccess = {
      ...access,
      getSlides: () => [
        {
          widthPx: 1280,
          heightPx: 720,
          nodes: [],
        } as never,
      ],
    }
    const skill = createSlideFixSkill(one)
    const r = await skill.executeTool({ id: 't1', name: 'read_slide', input: { slideIndex: 0 } })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('quality_read_only')
  })
})

describe('isQcEnabled', () => {
  it("localStorage 'ai-slides-qc'='0' is the kill switch", () => {
    localStorage.removeItem('ai-slides-qc')
    expect(isQcEnabled()).toBe(true)
    localStorage.setItem('ai-slides-qc', '0')
    expect(isQcEnabled()).toBe(false)
    localStorage.removeItem('ai-slides-qc')
  })
})

describe('qcSlidePage cancellation', () => {
  it('maps an oversized screenshot to quality unavailable without transport or mutation', async () => {
    const stream = vi.fn()
    const one: DeckAccess = {
      ...access,
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] } as never],
    }
    await expect(
      qcSlidePage({
        access: one,
        transport: { stream },
        pageIndex: 0,
        screenshot: { mime: 'image/png', base64: 'A'.repeat(2_700_000) },
      }),
    ).resolves.toMatchObject({ ok: false, edited: false, error: 'screenshot_unavailable' })
    expect(stream).not.toHaveBeenCalled()
  })

  it('maps transport failure to quality failure without changing the applied state', async () => {
    const one: DeckAccess = {
      ...access,
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] } as never],
    }
    const transport = {
      stream: (_request: unknown, callbacks: { onError: (error: string) => void }) => {
        queueMicrotask(() => callbacks.onError('offline'))
        return { cancel: vi.fn() }
      },
    }
    await expect(
      qcSlidePage({ access: one, transport, pageIndex: 0, screenshot: null }),
    ).resolves.toMatchObject({ ok: true, edited: false, error: 'offline' })
  })

  it('discards a late visual result after the session switches', async () => {
    let callbacks!: { onDelta: (text: string) => void; onDone: () => void }
    let current = true
    const one: DeckAccess = {
      ...access,
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] } as never],
    }
    const pending = qcSlidePage({
      access: one,
      transport: {
        stream: (_request: unknown, next: typeof callbacks) => {
          callbacks = next
          return { cancel: vi.fn() }
        },
      },
      pageIndex: 0,
      screenshot: null,
      isCurrent: () => current,
    })
    await vi.waitFor(() => expect(callbacks).toBeTruthy())
    current = false
    callbacks.onDelta('late warning')
    callbacks.onDone()
    await expect(pending).resolves.toMatchObject({ ok: false, error: 'stale_session', reply: '' })
  })

  it('rejects an already-aborted run before starting transport work', async () => {
    const stream = vi.fn()
    const controller = new AbortController()
    controller.abort()
    const one: DeckAccess = {
      ...access,
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] } as never],
    }

    expect(() =>
      qcSlidePage({
        access: one,
        transport: { stream },
        pageIndex: 0,
        screenshot: null,
        signal: controller.signal,
      }),
    ).toThrowError(expect.objectContaining({ name: 'AbortError' }))
    expect(stream).not.toHaveBeenCalled()
  })

  it('handles abort racing with loop setup without starting transport work', async () => {
    const stream = vi.fn()
    const controller = new AbortController()
    const one: DeckAccess = {
      ...access,
      getSlides: () => {
        controller.abort()
        return [{ widthPx: 1280, heightPx: 720, nodes: [] } as never]
      },
    }

    await expect(
      qcSlidePage({
        access: one,
        transport: { stream },
        pageIndex: 0,
        screenshot: null,
        signal: controller.signal,
      }),
    ).rejects.toMatchObject({ name: 'AbortError' })
    expect(stream).not.toHaveBeenCalled()
  })

  it('does not continue after capture resolves into an aborted run', async () => {
    let resolveCapture!: (value: string) => void
    const controller = new AbortController()
    const pending = captureCurrentQcShot({
      capture: () => new Promise<string>((resolve) => (resolveCapture = resolve)),
      signal: controller.signal,
      isCurrent: () => true,
    })
    controller.abort()
    resolveCapture('late-shot')
    await expect(pending).rejects.toMatchObject({ name: 'AbortError' })
  })
})
