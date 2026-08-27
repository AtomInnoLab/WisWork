import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { RenderSlide, ShapeRenderNode } from '@wiswork/pptx-render'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'
import { executePreparedTextFamilyTransaction } from '../src/renderer/ai/presentation-text-transactions'
import { textToolTransactionId } from '../src/renderer/ai/presentation-text-transactions'

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
