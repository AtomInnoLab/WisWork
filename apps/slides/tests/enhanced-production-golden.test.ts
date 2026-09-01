import { describe, expect, it, vi } from 'vitest'
import type { RenderSlide, ShapeRenderNode } from '@wiswork/pptx-render'
import { runEnhancedGolden } from '../../../packages/agent-runtime/src/production-golden'
import { createHostGoldenBridge } from '../../../packages/agent-runtime/tests/host-golden-bridge'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'
import { executePreparedTextFamilyTransaction } from '../src/renderer/ai/presentation-text-transactions'

const revision = (character: string) => `sha256:${character.repeat(64)}`

describe('Slides production Enhanced golden', () => {
  it('runs the production slide transaction and proves receipt/readback/rollback', async () => {
    let text = 'Before'
    const slide = (): RenderSlide => ({
      widthPx: 1280,
      heightPx: 720,
      scale: 1,
      background: { kind: 'solid', color: '#FFFFFF' },
      nodes: [
        {
          id: 'render-2',
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
                    text,
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
    })
    const history: string[] = []
    const hostApi = {
      preparePresentationTarget: vi.fn(async () => ({
        status: 'prepared' as const,
        expectedDeckRevision: revision('a'),
        target: {
          slideId: 'ppt/slides/slide1.xml',
          elementId: '{01234567-89AB-CDEF-0123-456789ABCDEF}',
          expectedType: 'text' as const,
          expectedFingerprint: revision('c'),
        },
      })),
      executePresentationTransaction: vi.fn(async (transaction) => {
        history.push(text)
        text = transaction.operations[0]?.kind === 'set_text' ? 'After' : text
        return {
          status: 'applied' as const,
          transactionId: transaction.transactionId,
          resultingDeckRevision: revision('b'),
          operationCount: transaction.operations.length,
        }
      }),
      cancelPresentationTransaction: vi.fn(async () => true),
      undoPresentationTransaction: vi.fn(async () => {
        const previous = history.pop()
        if (previous === undefined) return false
        text = previous
        return true
      }),
    }
    const executePresentationOperation: NonNullable<DeckAccess['executePresentationOperation']> = (
      request,
      signal,
    ) => executePreparedTextFamilyTransaction(hostApi, request, signal)
    const skill = createSlidesSkill({
      getSlides: () => [slide()],
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: vi.fn(),
      applyDeck: vi.fn(),
      executePresentationOperation,
      fitWidthPx: 1280,
    })
    const call = {
      id: 'slides-golden-call',
      name: 'set_element_text',
      input: { slideIndex: 0, sourceId: '2', paragraphs: [{ text: 'After' }] },
    }
    const result = await runEnhancedGolden('slides', {
      documentId: 'slides-golden-document',
      generation: 1,
      instruction: 'Change slide title',
      bridge: createHostGoldenBridge({ documentId: 'slides-golden-document', generation: 1, call }),
      skill,
      confirm: async () => ({ mutationReceiptId: 'slides-golden-receipt' }),
      readback: async () => ({ status: text === 'After' ? 'verified' : 'failed' }),
      rollback: async () => {
        const restored = await hostApi.undoPresentationTransaction()
        return { status: restored ? ('restored' as const) : ('failed' as const) }
      },
    })
    expect(hostApi.executePresentationTransaction).toHaveBeenCalledOnce()
    expect(result).toMatchObject({
      verification: { status: 'verified' },
      rollback: { status: 'restored' },
    })
    expect(text).toBe('Before')
    console.log(
      'ENHANCED_GOLDEN_REPORT',
      JSON.stringify({
        host: 'slides',
        verification: result.verification.status,
        rollback: result.rollback.status,
      }),
    )
  })
})
