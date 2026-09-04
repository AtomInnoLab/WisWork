import { describe, expect, it, vi } from 'vitest'
import type { RenderSlide } from '@wiswork/pptx-render'
import { createSlidesSkill } from '../src/renderer/ai/slides-skill'
import { executePreparedGeometryFamilyTransaction } from '../src/renderer/ai/presentation-geometry-transactions'

const blank = (): RenderSlide => ({
  widthPx: 1280,
  heightPx: 720,
  scale: 1,
  background: { kind: 'solid', color: '#FFFFFF' },
  nodes: [],
})

describe('build_deck', () => {
  it('creates every page and sends title and body through canonical transactions', async () => {
    let slides = [blank()]
    const executePresentationTransaction = vi.fn(async (transaction) => ({
      status: 'applied' as const,
      transactionId: transaction.transactionId,
      resultingDeckRevision: `sha256:${'a'.repeat(64)}`,
      operationCount: transaction.operations.length,
    }))
    const executePresentationOperation = vi.fn((request, signal?: AbortSignal) =>
      executePreparedGeometryFamilyTransaction(
        {
          preparePresentationTarget: async ({ slideIndex }) => ({
            status: 'prepared' as const,
            expectedDeckRevision: `sha256:${'0'.repeat(64)}`,
            target: {
              slideId: `ppt/slides/slide${slideIndex + 1}.xml`,
              expectedFingerprint: `sha256:${'1'.repeat(64)}`,
            },
          }),
          cancelPresentationTransaction: async () => true,
          executePresentationTransaction,
        },
        request,
        signal,
        async () => true,
      ),
    )
    ;(globalThis as unknown as { window: Record<string, unknown> }).window = {
      slidesApi: {
        addSlide: vi.fn(async ({ sourceIndex }: { sourceIndex: number }) => {
          const next = slides.slice()
          next.splice(sourceIndex + 1, 0, blank())
          return { slides: next, index: sourceIndex + 1 }
        }),
      },
    }
    const skill = createSlidesSkill({
      getSlides: () => slides,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => undefined,
      applyDeck: (next) => {
        slides = next
      },
      executePresentationOperation,
      fitWidthPx: 1280,
    })

    const result = await skill.executeTool({
      id: 'build',
      name: 'build_deck',
      input: {
        pages: [
          { title: '认识 LLM', body: ['从预测下一个词开始', '形成通用语言能力'] },
          { title: '如何工作', body: ['训练', '上下文', '生成'] },
          { title: '如何使用', body: ['明确目标', '验证结果'] },
        ],
      },
    })

    expect(result.mutated).toBe(true)
    expect(result.isError).not.toBe(true)
    expect(slides).toHaveLength(3)
    expect(executePresentationOperation).toHaveBeenCalledTimes(3)
    expect(executePresentationTransaction).toHaveBeenCalledTimes(3)
    const requests = executePresentationOperation.mock.calls.map(([request]) => request)
    expect(requests.every((request) => request.operations.length === 4)).toBe(true)
    expect(JSON.stringify(requests)).toContain('认识 LLM')
    expect(JSON.stringify(requests)).toContain('验证结果')
  })

  it('does not overwrite an existing presentation', async () => {
    const existing = blank()
    existing.nodes.push({ id: 'existing' } as never)
    const skill = createSlidesSkill({
      getSlides: () => [existing],
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => undefined,
      applyDeck: () => undefined,
      executePresentationOperation: vi.fn(),
      fitWidthPx: 1280,
    })
    const result = await skill.executeTool({
      id: 'build',
      invocationId: 'build-invocation',
      name: 'build_deck',
      input: {
        pages: [
          { title: 'A', body: ['B'] },
          { title: 'C', body: ['D'] },
        ],
      },
    })
    expect(result).toMatchObject({ isError: true, mutated: false })
  })
})
