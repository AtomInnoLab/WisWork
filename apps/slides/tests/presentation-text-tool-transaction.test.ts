import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { RenderSlide, ShapeRenderNode } from '@wiswork/pptx-render'
import type { PresentationReceipt } from '@wiswork/presentation-ops'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'
import { executePreparedTextFamilyTransaction } from '../src/renderer/ai/presentation-text-transactions'
import { textToolTransactionId } from '../src/renderer/ai/presentation-text-transactions'
import { executePreparedGeometryFamilyTransaction } from '../src/renderer/ai/presentation-geometry-transactions'

const fp = (char: string) => `sha256:${char.repeat(64)}`
const slide: RenderSlide = {
  widthPx: 1280,
  heightPx: 720,
  scale: 1,
  background: { kind: 'solid', color: '#FFFFFF' },
  nodes: [
    {
      id: 'r_2',
      sourceId: '2',
      type: 'text',
      box: {
        x: 0,
        y: 0,
        w: 100,
        h: 40,
        rotationDeg: 0,
        flipH: false,
        flipV: false,
        centerX: 50,
        centerY: 20,
      },
      fill: { kind: 'none' },
      text: {
        lines: [
          {
            runs: [
              {
                text: 'Before',
                x: 0,
                baselineY: 20,
                fontFamily: 'Arial',
                fontSizePx: 16,
                color: '#000000',
                bold: false,
                italic: false,
                underline: false,
                widthPx: 50,
              },
            ],
            top: 0,
            height: 20,
          },
        ],
        insets: { l: 0, t: 0, r: 0, b: 0 },
        anchor: 'top',
        fontScale: 1,
        wrap: true,
        contentHeight: 20,
      },
    } as ShapeRenderNode,
  ],
}

describe('Slides canonical text-family transactions', () => {
  let access: DeckAccess
  let executePresentationOperation: ReturnType<
    typeof vi.fn<NonNullable<DeckAccess['executePresentationOperation']>>
  >

  beforeEach(() => {
    executePresentationOperation = vi.fn(async () => ({
      receipt: {
        status: 'applied' as const,
        transactionId: 'tx',
        resultingDeckRevision: fp('b'),
        operationCount: 1,
      },
      authoritativeState: 'fresh' as const,
    }))
    access = {
      getSlides: () => [slide],
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: vi.fn(),
      applyDeck: vi.fn(),
      executePresentationOperation,
      fitWidthPx: 1280,
    }
  })

  it('routes rich set_element_text through one canonical set_text transaction', async () => {
    const result = await createSlidesSkill(access).executeTool({
      id: 'tool-call-1',
      name: 'set_element_text',
      input: {
        slideIndex: 0,
        sourceId: '2',
        paragraphs: [{ text: 'New', bold: true, color: '#123456' }],
      },
    })
    expect(executePresentationOperation).toHaveBeenCalledWith(
      expect.objectContaining({
        transactionId: expect.stringMatching(/^slides-text-/),
        slideIndex: 0,
        sourceId: '2',
        operation: expect.objectContaining({
          kind: 'set_text',
          paragraphs: [
            expect.objectContaining({
              runs: [expect.objectContaining({ text: 'New', bold: true })],
            }),
          ],
        }),
      }),
      undefined,
    )
    expect(result).toMatchObject({ mutated: true })
    expect(result.isError).not.toBe(true)
  })

  it.each([
    [
      'unchanged',
      { status: 'unchanged', transactionId: 'tx', code: 'operation_noop', operationCount: 1 },
      false,
    ],
    ['conflict', { status: 'conflict', transactionId: 'tx', code: 'target_stale' }, false],
    [
      'uncertain',
      { status: 'uncertain', transactionId: 'tx', code: 'write_state_uncertain' },
      true,
    ],
  ] as const)(
    'maps %s without falling back to legacy edit IPC',
    async (_label, receipt, mutated) => {
      executePresentationOperation.mockResolvedValue({
        receipt,
        authoritativeState: 'fresh',
      })
      const result = await createSlidesSkill(access).executeTool({
        id: `tool-${_label}`,
        name: 'set_element_text',
        input: { slideIndex: 0, sourceId: '2', paragraphs: [{ text: 'New' }] },
      })
      expect(result.isError).toBe(receipt.status === 'unchanged' ? undefined : true)
      expect(result.mutated).toBe(mutated)
    },
  )

  it('routes speaker notes through canonical set_speaker_notes', async () => {
    const result = await createSlidesSkill(access).executeTool({
      id: 'notes-1',
      name: 'set_speaker_notes',
      input: { slideIndex: 0, text: 'Presenter note' },
    })
    expect(executePresentationOperation).toHaveBeenCalledWith(
      expect.objectContaining({
        operation: { kind: 'set_speaker_notes', notes: 'Presenter note' },
      }),
      undefined,
    )
    expect(result).toMatchObject({ mutated: true })
  })

  it('routes set_element_style through canonical set_text without legacy edit IPC', async () => {
    const result = await createSlidesSkill(access).executeTool({
      id: 'style-1',
      name: 'set_element_style',
      input: { slideIndex: 0, sourceId: '2', bold: true, color: '#123456' },
    })
    expect(executePresentationOperation).toHaveBeenCalledWith(
      expect.objectContaining({
        sourceId: '2',
        operation: expect.objectContaining({ kind: 'set_text', paragraphs: expect.any(Array) }),
      }),
      undefined,
    )
    expect(result).toMatchObject({ mutated: true })
  })

  it('preserves a newer notes draft that appears while the transaction is running', async () => {
    let version = 1
    let visible = 'draft before agent'
    access.prepareSpeakerNotesWrite = vi.fn(async () => ({
      ready: true,
      expectedDraftVersion: version,
    }))
    access.applySpeakerNotes = vi.fn((_slideIndex, text, expectedVersion) => {
      if (version === expectedVersion) visible = text
    })
    executePresentationOperation.mockImplementation(async () => {
      version = 2
      visible = 'user typed during agent write'
      return {
        receipt: {
          status: 'applied',
          transactionId: 'notes-race',
          resultingDeckRevision: fp('c'),
          operationCount: 1,
        },
        authoritativeState: 'fresh',
      }
    })
    const result = await createSlidesSkill(access).executeTool({
      id: 'notes-race',
      name: 'set_speaker_notes',
      input: { slideIndex: 0, text: 'agent notes' },
    })
    expect(result).toMatchObject({ mutated: true })
    expect(visible).toBe('user typed during agent write')
    expect(access.applySpeakerNotes).toHaveBeenCalledWith(0, 'agent notes', 1)
  })

  it('does not prepare or execute a notes transaction when draft persistence is unsafe', async () => {
    access.prepareSpeakerNotesWrite = vi.fn(async () => ({
      ready: false,
      expectedDraftVersion: 3,
    }))
    const result = await createSlidesSkill(access).executeTool({
      id: 'notes-not-ready',
      name: 'set_speaker_notes',
      input: { slideIndex: 0, text: 'agent notes' },
    })
    expect(result).toMatchObject({ mutated: false, isError: true })
    expect(executePresentationOperation).not.toHaveBeenCalled()
  })

  it('stops the tool batch after an applied change whose authoritative refresh failed', async () => {
    executePresentationOperation.mockResolvedValue({
      receipt: {
        status: 'applied',
        transactionId: 'refresh-failed',
        resultingDeckRevision: fp('d'),
        operationCount: 1,
      },
      authoritativeState: 'reload_required',
    })
    const result = await createSlidesSkill(access).executeTool({
      id: 'refresh-failed',
      name: 'set_element_text',
      input: { slideIndex: 0, sourceId: '2', paragraphs: [{ text: 'landed' }] },
    })
    expect(result).toMatchObject({ mutated: true, stopToolBatch: true })
    expect(result.isError).not.toBe(true)
    expect(result.output).toContain('was applied')
  })
})

describe('Slides canonical geometry-family transactions', () => {
  it('routes set_element_transform through one canonical geometry transaction in points', async () => {
    const editTransform = vi.fn()
    ;(globalThis as any).window = { slidesApi: { editTransform } }
    const executePresentationOperation = vi.fn(async (request) => ({
      receipt: {
        status: 'applied' as const,
        transactionId: request.transactionId,
        resultingDeckRevision: fp('b'),
        operationCount: 1,
      },
      authoritativeState: 'fresh' as const,
    }))
    const result = await createSlidesSkill({
      getSlides: () => [{ ...slide, scale: 1 }],
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: vi.fn(),
      applyDeck: vi.fn(),
      executePresentationOperation,
      fitWidthPx: 1280,
    }).executeTool({
      id: 'geometry-1',
      name: 'set_element_transform',
      input: { slideIndex: 0, sourceId: '2', x: 96, y: 48, w: 192, h: 96, rotationDeg: 375 },
    })

    expect(executePresentationOperation).toHaveBeenCalledWith(
      expect.objectContaining({
        transactionId: expect.stringMatching(/^slides-geometry-/),
        slideIndex: 0,
        operations: [
          {
            sourceId: '2',
            geometry: { x: 72, y: 36, width: 144, height: 72, rotation: 15 },
          },
        ],
      }),
      undefined,
    )
    expect(result).toMatchObject({ mutated: true })
    expect(editTransform).not.toHaveBeenCalled()
  })

  it('allows moving an existing subpixel element without forcing a one-pixel resize', async () => {
    const executePresentationOperation = vi.fn(async (request) => ({
      receipt: {
        status: 'applied' as const,
        transactionId: request.transactionId,
        resultingDeckRevision: fp('b'),
        operationCount: 1,
      },
      authoritativeState: 'fresh' as const,
    }))
    const tiny = structuredClone(slide)
    tiny.nodes[0]!.box.w = 0.5
    tiny.nodes[0]!.box.h = 0.25
    const result = await createSlidesSkill({
      getSlides: () => [tiny],
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: vi.fn(),
      applyDeck: vi.fn(),
      executePresentationOperation,
      fitWidthPx: 1280,
    }).executeTool({
      id: 'geometry-subpixel',
      name: 'set_element_transform',
      input: { slideIndex: 0, sourceId: '2', x: 1.5 },
    })
    expect(result).toMatchObject({ mutated: true })
    expect(executePresentationOperation).toHaveBeenCalledWith(
      expect.objectContaining({
        operations: [
          expect.objectContaining({
            geometry: expect.objectContaining({ width: 0.375, height: 0.1875 }),
          }),
        ],
      }),
      undefined,
    )
  })

  it.each([
    [
      {
        status: 'unchanged' as const,
        transactionId: 'tx',
        code: 'operation_noop' as const,
        operationCount: 1,
      },
      false,
      false,
    ],
    [
      { status: 'conflict' as const, transactionId: 'tx', code: 'target_stale' as const },
      false,
      true,
    ],
    [
      { status: 'uncertain' as const, transactionId: 'tx', code: 'write_state_uncertain' as const },
      true,
      true,
    ],
  ])('maps geometry receipt state without legacy fallback', async (receipt, mutated, isError) => {
    const executePresentationOperation = vi.fn(async () => ({
      receipt,
      authoritativeState: 'fresh' as const,
    }))
    const result = await createSlidesSkill({
      getSlides: () => [slide],
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: vi.fn(),
      applyDeck: vi.fn(),
      executePresentationOperation,
      fitWidthPx: 1280,
    }).executeTool({
      id: 'geometry-state',
      name: 'set_element_transform',
      input: { slideIndex: 0, sourceId: '2', x: 10 },
    })
    expect(result.mutated).toBe(mutated)
    expect(result.isError === true).toBe(isError)
  })
})

describe('Slides canonical fill/stroke-family transactions', () => {
  it.each([
    [
      'set_element_fill',
      { fill: '#12abef' },
      { kind: 'set_fill', fill: { kind: 'solid', color: '#12ABEF' } },
    ],
    ['set_element_fill', { fill: 'none' }, { kind: 'set_fill', fill: { kind: 'none' } }],
    [
      'set_element_stroke',
      { color: '#12abef', widthPt: 2 },
      { kind: 'set_stroke', stroke: { color: '#12ABEF', width: 2 } },
    ],
    [
      'set_element_stroke',
      { color: '#12abef', widthPt: 0 },
      { kind: 'set_stroke', stroke: { color: '#12ABEF', width: 0 } },
    ],
    ['set_element_stroke', { remove: true }, { kind: 'set_stroke', stroke: null }],
  ] as const)(
    'routes %s through one canonical style transaction',
    async (name, patch, operation) => {
      const executePresentationOperation = vi.fn(async (request) => ({
        receipt: {
          status: 'applied' as const,
          transactionId: request.transactionId,
          resultingDeckRevision: fp('b'),
          operationCount: 1,
        },
        authoritativeState: 'fresh' as const,
      }))
      const legacy = { editFill: vi.fn(), editStroke: vi.fn() }
      ;(globalThis as any).window = { slidesApi: legacy }
      const result = await createSlidesSkill({
        getSlides: () => [slide],
        getCurrent: () => 0,
        getSelectedIds: () => [],
        applySlide: vi.fn(),
        applyDeck: vi.fn(),
        executePresentationOperation,
        fitWidthPx: 1280,
      }).executeTool({
        id: `paint-${name}`,
        name,
        input: { slideIndex: 0, sourceId: '2', ...patch },
      })
      expect(executePresentationOperation).toHaveBeenCalledWith(
        expect.objectContaining({ slideIndex: 0, operations: [{ sourceId: '2', ...operation }] }),
        undefined,
      )
      expect(legacy.editFill).not.toHaveBeenCalled()
      expect(legacy.editStroke).not.toHaveBeenCalled()
      expect(result).toMatchObject({ mutated: true })
    },
  )
})

describe('text tool invocation identity', () => {
  it('reuses one nonce for transport retry but not for a later identical invocation', async () => {
    const base = { id: 'same', name: 'set_element_text', input: { text: 'same' } }
    const first = await textToolTransactionId({ ...base, invocationId: 'run-1-call-1' })
    const retry = await textToolTransactionId({ ...base, invocationId: 'run-1-call-1' })
    const later = await textToolTransactionId({ ...base, invocationId: 'run-2-call-1' })
    expect(retry).toBe(first)
    expect(later).not.toBe(first)
  })
})

describe('renderer to preload canonical transaction contract', () => {
  const target = {
    slideId: 'ppt/slides/slide1.xml',
    elementId: '{01234567-89AB-CDEF-0123-456789ABCDEF}',
    expectedType: 'text' as const,
    expectedFingerprint: fp('a'),
  }

  it('uses the authoritative preparation verbatim in the canonical transaction', async () => {
    const api = {
      preparePresentationTarget: vi.fn(async () => ({
        status: 'prepared' as const,
        expectedDeckRevision: fp('0'),
        target,
      })),
      executePresentationTransaction: vi.fn(async (transaction) => ({
        status: 'applied' as const,
        transactionId: transaction.transactionId,
        resultingDeckRevision: fp('1'),
        operationCount: 1,
      })),
      cancelPresentationTransaction: vi.fn(async () => true),
    }
    await executePreparedTextFamilyTransaction(api, {
      transactionId: 'slides-text-contract',
      slideIndex: 0,
      sourceId: '2',
      operation: { kind: 'set_text', paragraphs: [{ runs: [{ text: 'after' }] }] },
    })
    expect(api.executePresentationTransaction).toHaveBeenCalledWith({
      transactionId: 'slides-text-contract',
      expectedDeckRevision: fp('0'),
      mode: 'atomic',
      operations: [
        {
          kind: 'set_text',
          clientId: 'text-1',
          target,
          paragraphs: [{ runs: [{ text: 'after' }] }],
        },
      ],
    })
  })

  it('forwards Stop to the main transaction cancellation channel', async () => {
    let finish!: (value: never) => void
    const api = {
      preparePresentationTarget: vi.fn(async () => ({
        status: 'prepared' as const,
        expectedDeckRevision: fp('0'),
        target,
      })),
      executePresentationTransaction: vi.fn(
        () => new Promise<never>((resolve) => (finish = resolve)),
      ),
      cancelPresentationTransaction: vi.fn(async () => true),
    }
    const controller = new AbortController()
    const pending = executePreparedTextFamilyTransaction(
      api,
      {
        transactionId: 'slides-text-stop',
        slideIndex: 0,
        sourceId: '2',
        operation: { kind: 'set_text', paragraphs: [{ runs: [{ text: 'after' }] }] },
      },
      controller.signal,
    )
    await vi.waitFor(() => expect(api.executePresentationTransaction).toHaveBeenCalled())
    const stopped = expect(pending).rejects.toMatchObject({ name: 'AbortError' })
    controller.abort()
    expect(api.cancelPresentationTransaction).toHaveBeenCalledWith('slides-text-stop')
    finish(undefined as never)
    await stopped
  })

  it('fails closed without invoking the transaction while the session is busy', async () => {
    const api = {
      preparePresentationTarget: vi.fn(async () => ({ status: 'busy' as const })),
      executePresentationTransaction: vi.fn(),
      cancelPresentationTransaction: vi.fn(async () => false),
    }
    await expect(
      executePreparedTextFamilyTransaction(api, {
        transactionId: 'slides-text-busy',
        slideIndex: 0,
        sourceId: '2',
        operation: { kind: 'set_text', paragraphs: [{ runs: [{ text: 'after' }] }] },
      }),
    ).resolves.toMatchObject({
      receipt: { status: 'unchanged', code: 'write_not_applied' },
    })
    expect(api.executePresentationTransaction).not.toHaveBeenCalled()
  })

  it('keeps an applied receipt authoritative when renderer refresh fails', async () => {
    const api = {
      preparePresentationTarget: vi.fn(async () => ({
        status: 'prepared' as const,
        expectedDeckRevision: fp('0'),
        target,
      })),
      executePresentationTransaction: vi.fn(async () => ({
        status: 'applied' as const,
        transactionId: 'slides-text-refresh',
        resultingDeckRevision: fp('1'),
        operationCount: 1,
      })),
      cancelPresentationTransaction: vi.fn(async () => false),
    }
    await expect(
      executePreparedTextFamilyTransaction(
        api,
        {
          transactionId: 'slides-text-refresh',
          slideIndex: 0,
          sourceId: '2',
          operation: { kind: 'set_text', paragraphs: [{ runs: [{ text: 'after' }] }] },
        },
        undefined,
        async () => false,
      ),
    ).resolves.toMatchObject({
      receipt: { status: 'applied' },
      authoritativeState: 'reload_required',
    })

    await expect(
      executePreparedTextFamilyTransaction(
        api,
        {
          transactionId: 'slides-text-refresh-rejected',
          slideIndex: 0,
          sourceId: '2',
          operation: { kind: 'set_text', paragraphs: [{ runs: [{ text: 'after' }] }] },
        },
        undefined,
        async () => {
          throw new Error('renderer reload failed')
        },
      ),
    ).resolves.toMatchObject({
      receipt: { status: 'applied' },
      authoritativeState: 'reload_required',
    })
  })
})

describe('renderer geometry transaction contract', () => {
  const target = (id: string) => ({
    slideId: 'ppt/slides/slide1.xml',
    elementId: id,
    expectedType: 'shape' as const,
    expectedFingerprint: fp(id === 'a' ? 'a' : 'b'),
  })

  it('prepares every durable target and emits one ordered atomic transaction', async () => {
    const api = {
      preparePresentationTarget: vi.fn(async ({ sourceId }: { sourceId?: string }) => ({
        status: 'prepared' as const,
        expectedDeckRevision: fp('0'),
        target: target(sourceId!),
      })),
      executePresentationTransaction: vi.fn(async (transaction) => ({
        status: 'applied' as const,
        transactionId: transaction.transactionId,
        resultingDeckRevision: fp('1'),
        operationCount: transaction.operations.length,
      })),
      cancelPresentationTransaction: vi.fn(async () => true),
    }
    await executePreparedGeometryFamilyTransaction(api, {
      transactionId: 'geometry-many',
      slideIndex: 0,
      operations: [
        { sourceId: 'a', geometry: { x: 1, y: 2, width: 3, height: 4, rotation: 5 } },
        { sourceId: 'b', geometry: { x: 6, y: 7, width: 8, height: 9 } },
      ],
    })
    expect(api.preparePresentationTarget).toHaveBeenCalledTimes(2)
    expect(api.executePresentationTransaction).toHaveBeenCalledWith(
      expect.objectContaining({
        expectedDeckRevision: fp('0'),
        mode: 'atomic',
        operations: [
          expect.objectContaining({
            kind: 'set_geometry',
            clientId: 'geometry-1',
            target: target('a'),
          }),
          expect.objectContaining({
            kind: 'set_geometry',
            clientId: 'geometry-2',
            target: target('b'),
          }),
        ],
      }),
    )
  })

  it('emits ordered fill/stroke operations using authoritative prepared targets', async () => {
    const api = {
      preparePresentationTarget: vi.fn(async ({ sourceId }: { sourceId?: string }) => ({
        status: 'prepared' as const,
        expectedDeckRevision: fp('0'),
        target: target(sourceId!),
      })),
      executePresentationTransaction: vi.fn(async (transaction) => ({
        status: 'applied' as const,
        transactionId: transaction.transactionId,
        resultingDeckRevision: fp('1'),
        operationCount: transaction.operations.length,
      })),
      cancelPresentationTransaction: vi.fn(async () => true),
    }
    await executePreparedGeometryFamilyTransaction(api, {
      transactionId: 'paint-many',
      slideIndex: 0,
      operations: [
        {
          sourceId: 'a',
          kind: 'set_fill',
          fill: { kind: 'solid', color: '#112233', transparency: 0.25 },
        },
        { sourceId: 'b', kind: 'set_stroke', stroke: null },
        {
          sourceId: 'a',
          kind: 'set_stroke',
          stroke: { color: '#445566', width: 2, dash: 'dash_dot' },
        },
      ],
    })
    expect(api.preparePresentationTarget).toHaveBeenCalledTimes(2)
    expect(api.executePresentationTransaction).toHaveBeenCalledWith({
      transactionId: 'paint-many',
      expectedDeckRevision: fp('0'),
      mode: 'atomic',
      operations: [
        {
          kind: 'set_fill',
          clientId: 'fill-1',
          target: target('a'),
          fill: { kind: 'solid', color: '#112233', transparency: 0.25 },
        },
        { kind: 'set_stroke', clientId: 'stroke-2', target: target('b'), stroke: null },
        {
          kind: 'set_stroke',
          clientId: 'stroke-3',
          target: target('a'),
          stroke: { color: '#445566', width: 2, dash: 'dash_dot' },
        },
      ],
    })
  })

  it('rejects invalid canonical paint/text payloads before target preparation', async () => {
    const api = {
      preparePresentationTarget: vi.fn(),
      executePresentationTransaction: vi.fn(),
      cancelPresentationTransaction: vi.fn(async () => true),
    }
    await expect(
      executePreparedGeometryFamilyTransaction(api, {
        transactionId: 'invalid-paint',
        slideIndex: 0,
        operations: [
          { sourceId: 'a', kind: 'set_stroke', stroke: { color: '#112233', width: 1001 } },
        ],
      }),
    ).resolves.toMatchObject({ receipt: { status: 'unchanged', code: 'write_not_applied' } })
    await expect(
      executePreparedGeometryFamilyTransaction(api, {
        transactionId: 'invalid-rich-text',
        slideIndex: 0,
        operations: [
          {
            sourceId: 'a',
            kind: 'set_text',
            paragraphs: [{ runs: [{ text: 'x', strike: true } as any] }],
          },
        ],
      }),
    ).resolves.toMatchObject({ receipt: { status: 'unchanged', code: 'write_not_applied' } })
    expect(api.preparePresentationTarget).not.toHaveBeenCalled()
    expect(api.executePresentationTransaction).not.toHaveBeenCalled()
  })

  it('cancels earlier preparations when a later target cannot prepare, allowing retry', async () => {
    const api = {
      preparePresentationTarget: vi
        .fn()
        .mockResolvedValueOnce({
          status: 'prepared' as const,
          expectedDeckRevision: fp('0'),
          target: target('a'),
        })
        .mockResolvedValueOnce({ status: 'busy' as const }),
      executePresentationTransaction: vi.fn(),
      cancelPresentationTransaction: vi.fn(async () => true),
    }
    await expect(
      executePreparedGeometryFamilyTransaction(api, {
        transactionId: 'prepare-cleanup',
        slideIndex: 0,
        operations: [
          { sourceId: 'a', kind: 'set_fill', fill: { kind: 'none' } },
          { sourceId: 'b', kind: 'set_fill', fill: { kind: 'none' } },
        ],
      }),
    ).resolves.toMatchObject({ receipt: { status: 'unchanged' } })
    expect(api.cancelPresentationTransaction).toHaveBeenCalledWith('prepare-cleanup')
    expect(api.executePresentationTransaction).not.toHaveBeenCalled()
  })

  it('refreshes after an uncertain write and requires reload when authority cannot be restored', async () => {
    const api = {
      preparePresentationTarget: vi.fn(async ({ sourceId }: { sourceId?: string }) => ({
        status: 'prepared' as const,
        expectedDeckRevision: fp('0'),
        target: target(sourceId!),
      })),
      executePresentationTransaction: vi.fn(async (transaction) => ({
        status: 'uncertain' as const,
        transactionId: transaction.transactionId,
        code: 'write_state_uncertain' as const,
      })),
      cancelPresentationTransaction: vi.fn(async () => true),
    }
    const request = {
      transactionId: 'uncertain-paint',
      slideIndex: 0,
      operations: [{ sourceId: 'a', kind: 'set_fill' as const, fill: { kind: 'none' as const } }],
    }
    await expect(
      executePreparedGeometryFamilyTransaction(api, request, undefined, async () => true),
    ).resolves.toMatchObject({ receipt: { status: 'uncertain' }, authoritativeState: 'fresh' })
    await expect(
      executePreparedGeometryFamilyTransaction(
        api,
        { ...request, transactionId: 'uncertain-paint-reload' },
        undefined,
        async () => false,
      ),
    ).resolves.toMatchObject({
      receipt: { status: 'uncertain' },
      authoritativeState: 'reload_required',
    })
    await expect(
      executePreparedGeometryFamilyTransaction(api, {
        ...request,
        transactionId: 'uncertain-paint-no-refresh',
      }),
    ).resolves.toMatchObject({
      receipt: { status: 'uncertain' },
      authoritativeState: 'reload_required',
    })
  })

  it('fails closed when target preparations observe different deck revisions', async () => {
    let count = 0
    const api = {
      preparePresentationTarget: vi.fn(async ({ sourceId }: { sourceId?: string }) => ({
        status: 'prepared' as const,
        expectedDeckRevision: fp(count++ === 0 ? '0' : '1'),
        target: target(sourceId!),
      })),
      executePresentationTransaction: vi.fn(),
      cancelPresentationTransaction: vi.fn(async () => true),
    }
    await expect(
      executePreparedGeometryFamilyTransaction(api, {
        transactionId: 'geometry-stale',
        slideIndex: 0,
        operations: [
          { sourceId: 'a', geometry: { x: 1, y: 2, width: 3, height: 4 } },
          { sourceId: 'b', geometry: { x: 5, y: 6, width: 7, height: 8 } },
        ],
      }),
    ).resolves.toMatchObject({ receipt: { status: 'conflict', code: 'target_stale' } })
    expect(api.executePresentationTransaction).not.toHaveBeenCalled()
  })

  it('rejects non-finite and out-of-bounds geometry before target preparation', async () => {
    const api = {
      preparePresentationTarget: vi.fn(),
      executePresentationTransaction: vi.fn(),
      cancelPresentationTransaction: vi.fn(async () => true),
    }
    for (const width of [Number.NaN, 0, 1_000_001]) {
      await expect(
        executePreparedGeometryFamilyTransaction(api, {
          transactionId: `geometry-invalid-${String(width)}`.replace(/[^A-Za-z0-9._:-]/g, '-'),
          slideIndex: 0,
          operations: [{ sourceId: 'a', geometry: { x: 1, y: 2, width, height: 4 } }],
        }),
      ).resolves.toMatchObject({ receipt: { status: 'unchanged', code: 'write_not_applied' } })
    }
    expect(api.preparePresentationTarget).not.toHaveBeenCalled()
    expect(api.executePresentationTransaction).not.toHaveBeenCalled()
  })

  it('classifies an IPC rejection after dispatch as uncertain instead of clean failure', async () => {
    const api = {
      preparePresentationTarget: vi.fn(async () => ({
        status: 'prepared' as const,
        expectedDeckRevision: fp('0'),
        target: target('a'),
      })),
      executePresentationTransaction: vi.fn(async () => {
        throw new Error('reply channel closed')
      }),
      cancelPresentationTransaction: vi.fn(async () => true),
    }
    await expect(
      executePreparedGeometryFamilyTransaction(api, {
        transactionId: 'geometry-reply-lost',
        slideIndex: 0,
        operations: [{ sourceId: 'a', geometry: { x: 1, y: 2, width: 3, height: 4 } }],
      }),
    ).resolves.toMatchObject({
      receipt: { status: 'uncertain', code: 'write_state_uncertain' },
    })
  })

  it('forwards abort and marks applied-but-unrefreshable state reload_required', async () => {
    let finish!: (value: any) => void
    const api = {
      preparePresentationTarget: vi.fn(async () => ({
        status: 'prepared' as const,
        expectedDeckRevision: fp('0'),
        target: target('a'),
      })),
      executePresentationTransaction: vi.fn(
        () => new Promise<PresentationReceipt>((resolve) => (finish = resolve)),
      ),
      cancelPresentationTransaction: vi.fn(async () => true),
    }
    const controller = new AbortController()
    const pending = executePreparedGeometryFamilyTransaction(
      api,
      {
        transactionId: 'geometry-abort',
        slideIndex: 0,
        operations: [{ sourceId: 'a', geometry: { x: 1, y: 2, width: 3, height: 4 } }],
      },
      controller.signal,
      async () => false,
    )
    await vi.waitFor(() => expect(api.executePresentationTransaction).toHaveBeenCalled())
    controller.abort()
    expect(api.cancelPresentationTransaction).toHaveBeenCalledWith('geometry-abort')
    finish({
      status: 'applied',
      transactionId: 'geometry-abort',
      resultingDeckRevision: fp('1'),
      operationCount: 1,
    })
    await expect(pending).resolves.toMatchObject({
      receipt: { status: 'applied' },
      authoritativeState: 'reload_required',
    })
  })

  it('cancels prepared targets and never dispatches when Stop wins after preparation', async () => {
    const controller = new AbortController()
    const api = {
      preparePresentationTarget: vi.fn(async () => {
        controller.abort()
        return {
          status: 'prepared' as const,
          expectedDeckRevision: fp('0'),
          target: target('a'),
        }
      }),
      executePresentationTransaction: vi.fn(),
      cancelPresentationTransaction: vi.fn(async () => true),
    }
    await expect(
      executePreparedGeometryFamilyTransaction(
        api,
        {
          transactionId: 'geometry-prepared-stop',
          slideIndex: 0,
          operations: [{ sourceId: 'a', geometry: { x: 1, y: 2, width: 3, height: 4 } }],
        },
        controller.signal,
      ),
    ).rejects.toMatchObject({ name: 'AbortError' })
    expect(api.cancelPresentationTransaction).toHaveBeenCalledWith('geometry-prepared-stop')
    expect(api.executePresentationTransaction).not.toHaveBeenCalled()
  })
})
