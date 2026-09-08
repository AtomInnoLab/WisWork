import { describe, expect, it, vi } from 'vitest'
import { createSlidesSkill } from '../src/renderer/ai/slides-skill'

describe('presentation design plan', () => {
  it('persists and returns the editable design contract before production', async () => {
    const saveSidecar = vi.fn(async () => undefined)
    const skill = createSlidesSkill({
      getSlides: () => [{ widthPx: 1280, heightPx: 720, nodes: [] }] as never,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => undefined,
      applyDeck: () => undefined,
      saveSidecar,
      fitWidthPx: 1280,
    })

    const result = await skill.executeTool({
      id: 'plan',
      name: 'plan_deck',
      input: {
        core_hook: 'One coherent story',
        style: 'Background: #0A0A0A\nTitle: 32pt',
        pages: [
          {
            title: 'A conclusion-led cover',
            brief: 'Open with the central claim',
            layout: 'cover',
            purpose: 'Open the story',
            visual: 'One typographic hero',
            acceptance: ['Title is dominant'],
            density: 'low',
          },
        ],
        prototype_pages: [0],
      },
    })

    expect(result.output).toContain('# DESIGN.md')
    expect(saveSidecar).toHaveBeenCalledWith(
      expect.objectContaining({
        topic: 'One coherent story',
        styleSkill: 'Background: #0A0A0A\nTitle: 32pt',
      }),
    )
  })
})
