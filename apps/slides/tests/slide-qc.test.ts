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
  parseQcReview,
  applyQcGeometryFixes,
} from '../src/renderer/ai/slide-qc'
import type { DeckAccess } from '../src/renderer/ai/slides-skill'
import { createElectronTransport } from '../src/renderer/ai/transport'

const access: DeckAccess = {
  getSlides: () => [],
  getCurrent: () => 0,
  getSelectedIds: () => [],
  applySlide: () => {},
  applyDeck: () => {},
  fitWidthPx: 1280,
}

describe('visual quality receipts', () => {
  it('accepts only bounded same-slide geometry fixes from visual review', () => {
    const slide = {
      widthPx: 1280,
      heightPx: 720,
      nodes: [{ sourceId: 'title', decoration: false }],
    } as never
    expect(
      parseQcReview(
        JSON.stringify({
          status: 'needs_fix',
          summary: 'Title is clipped',
          fixes: [{ sourceId: 'title', x: 72, y: 48, width: 900, height: 90 }],
        }),
        slide,
      ),
    ).toEqual({
      status: 'needs_fix',
      summary: 'Title is clipped',
      fixes: [{ sourceId: 'title', x: 72, y: 48, width: 900, height: 90 }],
    })
    expect(() =>
      parseQcReview(
        JSON.stringify({
          status: 'needs_fix',
          summary: 'bad target',
          fixes: [{ sourceId: 'other', x: 0, y: 0, width: 100, height: 100 }],
        }),
        slide,
      ),
    ).toThrow()
  })

  it('applies one atomic geometry-only correction to the reviewed page', async () => {
    const executePresentationOperation = vi.fn(async () => ({
      receipt: {
        status: 'applied' as const,
        transactionId: 'qc-fix',
        resultingDeckRevision: `sha256:${'a'.repeat(64)}`,
        operationCount: 1,
      },
      authoritativeState: 'fresh' as const,
    }))
    const one: DeckAccess = {
      ...access,
      getSlides: () => [
        {
          widthPx: 1280,
          heightPx: 720,
          scale: 1,
          nodes: [{ sourceId: 'title', decoration: false }],
        } as never,
      ],
      executePresentationOperation,
    }
    await expect(
      applyQcGeometryFixes(one, 0, [{ sourceId: 'title', x: 72, y: 48, width: 900, height: 90 }]),
    ).resolves.toBe(true)
    expect(executePresentationOperation).toHaveBeenCalledWith(
      expect.objectContaining({
        slideIndex: 0,
        operations: [
          expect.objectContaining({
            kind: 'set_geometry',
            sourceId: 'title',
            geometry: { x: 54, y: 36, width: 675, height: 67.5 },
          }),
        ],
      }),
      undefined,
    )
  })

  it('rejects out-of-canvas corrections at the mutation boundary', async () => {
    const executePresentationOperation = vi.fn()
    const one: DeckAccess = {
      ...access,
      getSlides: () => [
        {
          widthPx: 1280,
          heightPx: 720,
          scale: 1,
          nodes: [{ sourceId: 'title', decoration: false }],
        } as never,
      ],
      executePresentationOperation,
    }
    await expect(
      applyQcGeometryFixes(one, 0, [{ sourceId: 'title', x: 1200, y: 48, width: 900, height: 90 }]),
    ).resolves.toBe(false)
    expect(executePresentationOperation).not.toHaveBeenCalled()
  })
  it('measures the exact IPC envelope including settings before slidesApi.aiStream', async () => {
    const previous = window.slidesApi
    const aiStream = vi.fn()
    const onError = vi.fn()
    Object.assign(window, {
      slidesApi: {
        aiStream,
        aiStreamCancel: vi.fn(),
        onAiStream: vi.fn(() => () => {}),
      },
    })
    try {
      createElectronTransport(
        () => ({ hugeEnvelopeSetting: 'X'.repeat(2 * 1024 * 1024) }) as never,
        { maxSerializedRequestBytes: 2 * 1024 * 1024 },
      ).stream(
        { system: 'small', messages: [], tools: [] },
        { onDelta: vi.fn(), onToolCall: vi.fn(), onDone: vi.fn(), onError },
      )
      await vi.waitFor(() => expect(onError).toHaveBeenCalledWith('quality_request_too_large'))
      expect(aiStream).not.toHaveBeenCalled()
    } finally {
      Object.assign(window, { slidesApi: previous })
    }
  })

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

  it('publishes distinct durable creation IDs for two overflowing production elements', async () => {
    const published = vi.fn()
    const overflowing = (sourceId: string) => ({
      id: sourceId,
      sourceId,
      type: 'shape',
      box: {
        x: -20,
        y: 0,
        w: 10,
        h: 10,
        rotationDeg: 0,
        flipH: false,
        flipV: false,
        centerX: -15,
        centerY: 5,
      },
      fill: { kind: 'none' },
    })
    const one: DeckAccess = {
      ...access,
      getSlides: () => [
        {
          widthPx: 1280,
          heightPx: 720,
          nodes: [overflowing('runtime-a'), overflowing('runtime-b')],
        } as never,
      ],
    }
    await publishAppliedDeterministicQuality({
      transactionId: 'tx-two',
      receiptStatus: 'applied',
      sessionId: 's',
      pageIndexes: [0],
      completedKeys: new Set(),
      access: one,
      prepareSlide: async () => ({
        status: 'prepared',
        slideId: 'ppt/slides/slide1.xml',
        elementIds: { 'runtime-a': 'creation-a', 'runtime-b': 'creation-b' },
      }),
      publish: published,
      isCurrent: () => true,
    })
    const receipt = published.mock.calls[0]![0]
    expect(receipt.findings.map((finding: { elementId?: string }) => finding.elementId)).toEqual([
      'creation-a',
      'creation-b',
    ])
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

  it('publishes deterministic receipts for targets 21 through 50 independently of visual cap', async () => {
    const slides = Array.from({ length: 50 }, () => ({ widthPx: 1280, heightPx: 720, nodes: [] }))
    const published = vi.fn()
    const pages = Array.from({ length: 50 }, (_, index) => index)
    await expect(
      publishAppliedDeterministicQuality({
        transactionId: 'tx-50',
        receiptStatus: 'applied',
        sessionId: 's',
        pageIndexes: pages,
        completedKeys: new Set(),
        access: { ...access, getSlides: () => slides as never },
        prepareSlide: async (pageIndex) => ({
          status: 'prepared',
          slideId: `ppt/slides/slide${pageIndex + 1}.xml`,
        }),
        publish: published,
        isCurrent: () => true,
      }),
    ).resolves.toHaveLength(50)
    expect(published).toHaveBeenCalledTimes(50)
  })

  it('emits an explicit visual capacity receipt instead of silently skipping', () => {
    expect(
      toVisualQualityReceipt('qc-1', 'tx-1', 'ppt/slides/slide21.xml', {
        ok: false,
        edited: false,
        reply: '',
        preIssues: 0,
        postIssues: 0,
        error: 'visual_capacity_exceeded',
      }),
    ).toMatchObject({ status: 'unavailable', code: 'visual_capacity_exceeded' })
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

  it('rejects a full serialized request over 2 MiB before delegating transport', async () => {
    const delegate = vi.fn()
    const one: DeckAccess = {
      ...access,
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] } as never],
    }
    await expect(
      qcSlidePage({
        access: one,
        transport: { stream: delegate },
        pageIndex: 0,
        screenshot: null,
        systemSuffix: () => 'X'.repeat(2 * 1024 * 1024),
      }),
    ).resolves.toMatchObject({ edited: false, error: 'quality_request_too_large' })
    expect(delegate).not.toHaveBeenCalled()
  })

  it('keeps a critical deterministic failure even when visual review says OK', async () => {
    const delegate = vi.fn((_request, callbacks) => {
      queueMicrotask(() => {
        callbacks.onDelta('OK')
        callbacks.onDone()
      })
      return { cancel: vi.fn() }
    })
    const one: DeckAccess = {
      ...access,
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] } as never],
    }
    await expect(
      qcSlidePage({ access: one, transport: { stream: delegate }, pageIndex: 0, screenshot: null }),
    ).resolves.toMatchObject({ edited: false, reply: 'empty_slide', postIssues: 1 })
    expect(delegate).toHaveBeenCalledOnce()
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
