import { describe, expect, it, vi } from 'vitest'
import type { RenderSlide } from '@wiswork/pptx-render'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'

const slides: RenderSlide[] = [0, 1].map(() => ({
  widthPx: 1280,
  heightPx: 720,
  scale: 1,
  background: { kind: 'solid', color: '#FFFFFF' },
  nodes: [],
}))

describe('Slides canonical background tool transaction', () => {
  it('compiles slideIndex=-1 into one all-slide atomic request without legacy IPC', async () => {
    const editBackground = vi.fn()
    Object.defineProperty(window, 'slidesApi', { configurable: true, value: { editBackground } })
    const executePresentationOperation = vi.fn(async () => ({
      receipt: {
        status: 'applied' as const,
        transactionId: 'tx',
        resultingDeckRevision: `sha256:${'a'.repeat(64)}`,
        operationCount: 2,
      },
      authoritativeState: 'fresh' as const,
    }))
    const access: DeckAccess = {
      getSlides: () => slides,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: vi.fn(),
      applyDeck: vi.fn(),
      executePresentationOperation,
      fitWidthPx: 1280,
    }
    const result = await createSlidesSkill(access).executeTool({
      id: 'background-call',
      invocationId: 'background-invocation',
      name: 'set_slide_background',
      input: { slideIndex: -1, color: '#1a2b3c' },
    })
    expect(executePresentationOperation).toHaveBeenCalledOnce()
    expect(executePresentationOperation).toHaveBeenCalledWith(
      expect.objectContaining({
        transactionId: expect.stringMatching(/^slides-background-/),
        backgrounds: [
          { slideIndex: 0, color: '#1A2B3C' },
          { slideIndex: 1, color: '#1A2B3C' },
        ],
      }),
      undefined,
    )
    expect(result).toMatchObject({ mutated: true })
    expect(editBackground).not.toHaveBeenCalled()
  })
})
