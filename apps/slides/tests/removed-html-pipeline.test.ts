import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

const read = (path: string) => readFileSync(new URL(path, import.meta.url), 'utf8')

describe('removed HTML-to-PPTX cloud pipeline', () => {
  it('has no renderer or preload API and no main-process handler', () => {
    const skill = read('../src/renderer/ai/slides-skill.ts')
    const panel = read('../src/renderer/ai/AiPanel.tsx')
    const preload = read('../src/preload/index.ts')
    const shared = read('../src/shared/ipc.ts')
    const main = read('../src/main/slides-main.ts')

    expect(skill).not.toContain('generateFromHtml')
    expect(panel).not.toContain('generateFromHtml')
    expect(panel).not.toContain('htmlToPptx')
    expect(preload).not.toContain("'slides:html-to-pptx'")
    expect(shared).not.toContain('htmlToPptx')
    expect(main).not.toContain("'slides:html-to-pptx'")
    expect(main).not.toContain('CLOUD_PAGE_PREFIX')
    expect(main).not.toContain('issuedCloudPages')
  })

  it('keeps supported local slide creation tools', async () => {
    const { createSlidesSkill } = await import('../src/renderer/ai/slides-skill')
    const skill = createSlidesSkill({
      getSlides: () => [],
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => {},
      applyDeck: () => {},
      fitWidthPx: 1280,
    })
    expect(skill.tools.map((tool) => tool.name)).toEqual(
      expect.arrayContaining(['add_slide', 'add_text_box', 'add_shape', 'add_smartart']),
    )
  })
})
