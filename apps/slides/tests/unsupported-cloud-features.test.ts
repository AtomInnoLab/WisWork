import { describe, expect, it } from 'vitest'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'
import type { AgentToolCall } from '../src/shared/ipc'
import type { RenderSlide } from '@wiswork/pptx-render'

const disabledTools = [
  'generate_image',
  'analyze_media',
  'regenerate_slide',
  'generate_deck',
] as const

function makeAccess(slides: RenderSlide[] = []): DeckAccess {
  return {
    getSlides: () => slides,
    getCurrent: () => 0,
    getSelectedIds: () => [],
    applySlide: () => {},
    applyDeck: () => {},
    executePresentationOperation: async (request) => ({
      receipt: {
        status: 'applied',
        transactionId: request.transactionId,
        resultingDeckRevision: `sha256:${'a'.repeat(64)}`,
        operationCount: 'operations' in request ? request.operations.length : 1,
        createdIds: ['{11111111-2222-3333-4444-555555555555}'],
      },
      authoritativeState: 'fresh',
    }),
    fitWidthPx: 1280,
  }
}

describe('unsupported cloud features', () => {
  it('removes production presentation hooks when verified completion is rolled back', () => {
    const access = { ...makeAccess(), taskReviewAdapter: {} as DeckAccess['taskReviewAdapter'] }
    const skill = createSlidesSkill(access, {
      planning: true,
      verifiedCompletion: false,
      visualReview: true,
      autoCorrection: false,
    })
    expect(skill.presentation).toBeUndefined()
  })

  it('does not advertise or prompt for disabled cloud tools', () => {
    const skill = createSlidesSkill(makeAccess())
    const names = skill.tools.map((tool) => tool.name)
    for (const name of disabledTools) {
      expect(names).not.toContain(name)
      expect(skill.systemPrompt).not.toContain(name)
      expect(JSON.stringify(skill.tools)).not.toContain(name)
    }
  })

  it('planning never injects synthetic incomplete progress into later turns', async () => {
    const blank = { widthPx: 1280, heightPx: 720, nodes: [] } as unknown as RenderSlide
    Object.assign(window, {
      slidesApi: { addElement: async () => ({ slide: blank, sourceId: 'local-1' }) },
    })
    const skill = createSlidesSkill(makeAccess([blank]))
    const result = await skill.executeTool({
      id: 'plan-1',
      name: 'plan_deck',
      input: {
        core_hook: 'Hook',
        style: 'Style',
        pages: [{ title: 'One', brief: 'Brief', layout: 'title' }],
      },
    } as AgentToolCall)
    expect(result.isError).toBeUndefined()
    for (const text of ['first', 'second']) {
      const addition = await skill.executeTool({
        id: `add-${text}`,
        name: 'add_text_box',
        input: {
          slideIndex: 0,
          x: 20,
          y: 20,
          w: 200,
          h: 60,
          paragraphs: [{ text }],
        },
      } as AgentToolCall)
      expect(addition.isError).toBeUndefined()
    }
    const contexts = [skill.buildContext?.() ?? '', skill.buildContext?.() ?? '']
    for (const context of contexts) {
      expect(context).not.toContain('<generation-progress>')
      expect(context).not.toMatch(/Incomplete|0 generated|still missing/)
    }
    const text = result.output + '\n' + contexts.join('\n')
    expect(text).not.toMatch(/do not stop|do not claim completion/i)
    for (const name of disabledTools) expect(text).not.toContain(name)
  })

  it.each(disabledTools)('rejects a defensive %s invocation with the stable code', async (name) => {
    const skill = createSlidesSkill(makeAccess())
    const result = await skill.executeTool(
      { id: 'call-1', name, input: {} } as AgentToolCall,
      new AbortController().signal,
    )

    expect(result.mutated).toBe(false)
    expect(result.output).toContain('unsupported_feature')
  })
})
