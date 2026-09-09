import { describe, expect, it, vi } from 'vitest'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'
import { createPcHostRegistration } from '@wiswork/agent-runtime'

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
  it('survives the Enhanced registration allowlist as a read', () => {
    const registration = createPcHostRegistration({
      host: 'slides',
      documentId: 'deck',
      generation: 1,
      skill: createSlidesSkill(access()),
    })
    expect(registration.tools.some((tool) => tool.name === 'screenshot_slide')).toBe(true)
    expect(registration.mutatingTools).not.toContain('screenshot_slide')
  })
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

  it('bounds an unavailable capture and reports a retryable error', async () => {
    vi.useFakeTimers()
    const deck = access()
    deck.captureSlideScreenshot = vi.fn(async () => await new Promise<never>(() => undefined))
    try {
      const pending = createSlidesSkill(deck).executeTool({
        id: 'shot-timeout',
        name: 'screenshot_slide',
        input: { slideIndex: 0 },
      })
      await vi.advanceTimersByTimeAsync(15_000)
      await expect(pending).resolves.toMatchObject({
        isError: true,
        mutated: false,
      })
      await expect(pending).resolves.toHaveProperty(
        'output',
        expect.stringContaining('visual_capture_unavailable'),
      )
      await expect(pending).resolves.toHaveProperty('output', expect.stringContaining('retry'))
    } finally {
      vi.useRealTimers()
    }
  })

  it('normalizes a rejected capture as retryable instead of throwing', async () => {
    const deck = access()
    deck.captureSlideScreenshot = vi.fn(async () => {
      throw new Error('renderer detached')
    })
    await expect(
      createSlidesSkill(deck).executeTool({
        id: 'shot-rejected',
        name: 'screenshot_slide',
        input: { slideIndex: 0 },
      }),
    ).resolves.toMatchObject({
      isError: true,
      mutated: false,
      output: expect.stringContaining('visual_capture_unavailable'),
    })
  })
})
