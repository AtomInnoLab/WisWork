import { describe, expect, it, vi } from 'vitest'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'

const slide = {
  id: 'slide-1',
  widthPx: 1280,
  heightPx: 720,
  nodes: [],
} as any

const access = (): DeckAccess => ({
  getSlides: () => [slide],
  getCurrent: () => 0,
  getSelectedIds: () => [],
  applySlide: () => undefined,
  applyDeck: () => undefined,
  fitWidthPx: 1280,
  captureSlideScreenshot: vi.fn(async () => ({ base64: 'aGVsbG8=', mime: 'image/png' })),
})

describe('Slides main-agent screenshot tool', () => {
  it('advertises a read-only screenshot and returns it as model content', async () => {
    const deck = access()
    const skill = createSlidesSkill(deck)
    expect(skill.tools.map((tool) => tool.name)).toContain('screenshot_slide')

    const result = await skill.executeTool({
      id: 'shot-1',
      name: 'screenshot_slide',
      input: { slideIndex: 0 },
    })

    expect(result).toMatchObject({
      mutated: false,
      summary: 'Captured slide 1',
      modelContent: [{ type: 'image', image: { base64: 'aGVsbG8=', mime: 'image/png' } }],
    })
    expect(deck.captureSlideScreenshot).toHaveBeenCalledWith(0)
  })

  it('rejects an out-of-range slide without capturing', async () => {
    const deck = access()
    const result = await createSlidesSkill(deck).executeTool({
      id: 'shot-2',
      name: 'screenshot_slide',
      input: { slideIndex: 7 },
    })
    expect(result.isError).toBe(true)
    expect(deck.captureSlideScreenshot).not.toHaveBeenCalled()
  })
})
