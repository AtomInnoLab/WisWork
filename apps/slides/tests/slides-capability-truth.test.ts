import { describe, expect, it } from 'vitest'
import {
  createSlidesSkill,
  reviewSlidesFinalResponse,
  type DeckAccess,
} from '../src/renderer/ai/slides-skill'
import type { RenderSlide } from '@wiswork/pptx-render'

const CORRECTION =
  '[System correction] This requested edit is supported. Use the available Slides editing tools to apply it, verify the result, and only then report completion. Do not stop at inspection or advice.'

function access(): DeckAccess {
  return {
    getSlides: () => [] as RenderSlide[],
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
        createdIds: [],
      },
      authoritativeState: 'fresh',
    }),
    fitWidthPx: 1280,
  }
}

describe('Slides capability truth', () => {
  it.each([
    '但当前可用的幻灯片编辑接口无法直接修改现有文字的字体颜色，只能改文字内容或位置。',
    '当前工具不能调整标题栏的位置，所以我暂未修改。',
    'The available editor does not support changing the font color of existing text.',
    'I cannot move or resize the title with the available tools.',
  ])('corrects a tool-free false denial: %s', (text) => {
    expect(reviewSlidesFinalResponse({ text, mutated: false })).toBe(CORRECTION)
  })

  it('wires the pure policy into the Slides skill', () => {
    expect(
      createSlidesSkill(access()).reviewFinalResponse?.({
        text: 'The tools cannot change an element position.',
        mutated: false,
      }),
    ).toBe(CORRECTION)
  })

  it.each([
    'Would you like me to change the title color and position?',
    'Can I change the font color?',
    'Embedded font packaging is not supported.',
    'Morph transitions are unavailable in this editor.',
    '已尝试提交统一调整，但当前 PowerPoint 读取失败，未能应用修改。请重新打开或刷新 PPT 后再试。',
    'set_element_style failed closed, so I did not claim the edit succeeded.',
    'set_element_style returned unsupported for this selected object, so I cannot change its font color.',
    'This title is locked and read-only, so I cannot move it.',
    'set_element_transform 返回 fail-closed，因此不能移动这个标题。',
    'I cannot guarantee the exact font color without first inspecting the slide.',
  ])('does not reject safe or honest terminal text: %s', (text) => {
    expect(reviewSlidesFinalResponse({ text, mutated: false })).toBeUndefined()
  })

  it('does not reject a denial after a successful mutation', () => {
    expect(
      reviewSlidesFinalResponse({
        text: 'I cannot move another title because it is locked.',
        mutated: true,
      }),
    ).toBeUndefined()
  })

  it('publishes an executable capability map and multi-page edit contract', () => {
    const prompt = createSlidesSkill(access()).systemPrompt
    expect(prompt).toContain('set_element_style')
    expect(prompt).toContain('setStyle')
    expect(prompt).toContain('set_element_transform')
    expect(prompt).toContain('setBox')
    expect(prompt).toContain('each target page')
    expect(prompt).toContain('must not finish with inspection or advice only')
    expect(prompt).toContain('unsupported or fail-closed')
  })

  it('uses a fixed correction that cannot echo document or response content', () => {
    const secret = 'CONFIDENTIAL_SLIDE_VALUE_91'
    const result = reviewSlidesFinalResponse({
      text: `The tools cannot change text color. ${secret}`,
      mutated: false,
    })
    expect(result).toBe(CORRECTION)
    expect(result).not.toContain(secret)
  })
})
