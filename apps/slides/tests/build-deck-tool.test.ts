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

async function plannedImageDeck(
  firstImage:
    'missing' | 'throws' | 'cancel' | 'success' | 'write_then_missing' | 'write_then_throw',
  refreshFails = false,
) {
  let slides = [blank()]
  let writtenSlide: RenderSlide | undefined
  const controller = new AbortController()
  const applySlide = vi.fn((index: number, slide: RenderSlide) => {
    slides[index] = slide
  })
  const insertImageUrl = vi.fn(async ({ slideIndex }: { slideIndex: number }) => {
    if (slideIndex === 0) {
      if (firstImage.startsWith('write_then_')) {
        writtenSlide = {
          ...blank(),
          nodes: [
            {
              id: 'native-image',
              sourceId: 'native-image',
              type: 'picture',
              box: {
                x: 730,
                y: 0,
                w: 550,
                h: 720,
                rotationDeg: 0,
                flipH: false,
                flipV: false,
                centerX: 1005,
                centerY: 360,
              },
            },
          ],
        }
        if (firstImage === 'write_then_throw') throw new Error('rebuild_failed')
        return null
      }
      if (firstImage === 'missing') return null
      if (firstImage === 'throws') throw new Error('slides_session_busy')
      if (firstImage === 'cancel') {
        controller.abort()
        throw new DOMException('Cancelled', 'AbortError')
      }
    }
    return {
      sourceId: `image-${slideIndex}`,
      slide: { ...blank(), background: { kind: 'solid' as const, color: '#123456' } },
    }
  })
  const executePresentationOperation = vi.fn(async (request) => ({
    receipt: {
      status: 'applied' as const,
      transactionId: request.transactionId,
      resultingDeckRevision: `sha256:${'a'.repeat(64)}`,
      operationCount: request.operations.length,
    },
    authoritativeState: 'fresh' as const,
  }))
  const deleteSlide = vi.fn(async (index: number) => {
    slides = slides.filter((_, i) => i !== index)
    return slides
  })
  ;(globalThis as unknown as { window: Record<string, unknown> }).window = {
    slidesApi: {
      imageSearch: vi.fn(async () => ({
        images: [{ imageUrl: 'https://images.example/hero.jpg', title: 'Hero' }],
        method: 'serper',
      })),
      addSlide: vi.fn(async () => ({ slides: [...slides, blank()], index: slides.length })),
      insertImageUrl,
      deleteSlide,
    },
  }
  const skill = createSlidesSkill({
    getSlides: () => slides,
    getCurrent: () => 0,
    getSelectedIds: () => [],
    applySlide,
    applyDeck: (next) => {
      slides = next
    },
    refreshAuthoritativeState: async () => {
      if (refreshFails) return false
      if (writtenSlide) slides[0] = writtenSlide
      return true
    },
    executePresentationOperation,
    fitWidthPx: 1280,
  })
  const pages = [
    {
      layout: 'cover',
      title: 'Start',
      body: ['Introduction'],
      imageUrl: 'https://images.example/hero.jpg',
      imageAlt: 'Hero',
    },
    {
      layout: 'statement',
      title: 'End',
      body: ['Conclusion'],
      imageUrl: 'https://images.example/hero.jpg',
      imageAlt: 'Hero',
    },
  ]
  await skill.executeTool({
    id: 'plan',
    name: 'plan_deck',
    input: { core_hook: 'Hook', style: 'Dark', pages },
  })
  await skill.executeTool({ id: 'search', name: 'image_search', input: { query: 'team' } })
  return {
    skill,
    pages,
    controller,
    insertImageUrl,
    applySlide,
    executePresentationOperation,
    deleteSlide,
  }
}

describe('build_deck', () => {
  it.each(['write_then_missing', 'write_then_throw'] as const)(
    'refreshes native image state before recommending repair after %s',
    async (mode) => {
      const test = await plannedImageDeck(mode)
      const result = await test.skill.executeTool({
        id: 'build',
        invocationId: 'build',
        name: 'build_deck',
        input: { theme: { mode: 'dark' }, pages: test.pages },
      })
      expect(result).toMatchObject({ isError: true, mutated: true })
      const read = await test.skill.executeTool({
        id: 'read',
        name: 'read_slide',
        input: { slideIndex: 0 },
      })
      expect(read.output).toContain('native-image')
      expect(test.insertImageUrl).toHaveBeenCalledTimes(2)
    },
  )

  it('does not expose stale reads or permit blind repair when native refresh fails', async () => {
    const test = await plannedImageDeck('write_then_missing', true)
    const result = await test.skill.executeTool({
      id: 'build',
      invocationId: 'build',
      name: 'build_deck',
      input: { theme: { mode: 'dark' }, pages: test.pages },
    })
    expect(result).toMatchObject({ isError: true, mutated: true, stopToolBatch: true })
    expect(result.output).toContain('authoritative_reload_required')
    const read = await test.skill.executeTool({
      id: 'read',
      name: 'read_slide',
      input: { slideIndex: 0 },
    })
    expect(read.isError).toBe(true)
    expect(read.output).toContain('authoritative_reload_required')
    const repair = await test.skill.executeTool({
      id: 'repair',
      name: 'delete_slide',
      input: { slideIndex: 1 },
    })
    expect(repair.isError).toBe(true)
    expect(test.deleteSlide).not.toHaveBeenCalled()
  })

  it.each(['missing', 'throws'] as const)(
    'permits same-run repair and applies later images after an image %s',
    async (failure) => {
      const test = await plannedImageDeck(failure)
      const result = await test.skill.executeTool({
        id: 'build',
        invocationId: 'build',
        name: 'build_deck',
        input: { theme: { mode: 'dark' }, pages: test.pages },
      })
      expect(result).toMatchObject({ isError: true, mutated: true, stopToolBatch: true })
      expect(result.output).toContain('pages 1')
      expect(result.output).toMatch(/read|inspect/i)
      expect(test.insertImageUrl).toHaveBeenCalledTimes(2)
      expect(test.applySlide).toHaveBeenCalledWith(
        1,
        expect.objectContaining({ background: { kind: 'solid', color: '#123456' } }),
      )
      const repair = await test.skill.executeTool({
        id: 'repair',
        name: 'delete_slide',
        input: { slideIndex: 1 },
      })
      expect(repair.isError).not.toBe(true)
      expect(test.deleteSlide).toHaveBeenCalledOnce()
    },
  )

  it('stops image insertion on cancellation but leaves the partial deck available for a later repair', async () => {
    const test = await plannedImageDeck('cancel')
    await expect(
      test.skill.executeTool(
        {
          id: 'build',
          invocationId: 'build',
          name: 'build_deck',
          input: { theme: { mode: 'dark' }, pages: test.pages },
        },
        test.controller.signal,
      ),
    ).rejects.toMatchObject({ name: 'AbortError' })
    expect(test.insertImageUrl).toHaveBeenCalledOnce()
    const repair = await test.skill.executeTool({
      id: 'repair',
      name: 'delete_slide',
      input: { slideIndex: 1 },
    })
    expect(repair.isError).not.toBe(true)
    expect(test.deleteSlide).toHaveBeenCalledOnce()
  })

  it('keeps cover and final statement page numbers outside the topmost image panel', async () => {
    const test = await plannedImageDeck('success')
    const result = await test.skill.executeTool({
      id: 'build',
      invocationId: 'build',
      name: 'build_deck',
      input: { theme: { mode: 'dark' }, pages: test.pages },
    })
    expect(result.isError).not.toBe(true)
    for (const [index, [request]] of test.executePresentationOperation.mock.calls.entries()) {
      const footer = request.operations.find(
        (op: { clientId?: string; kind: string }) =>
          op.kind === 'add_text_box' && op.clientId === `deck-page-${index}`,
      )
      const title = request.operations.find(
        (op: { clientId?: string; kind: string }) =>
          op.kind === 'add_text_box' && op.clientId === `deck-title-${index}`,
      )
      expect(test.insertImageUrl).toHaveBeenNthCalledWith(
        index + 1,
        expect.objectContaining({ slideIndex: index, xPx: 730, yPx: 0, wPx: 550, hPx: 720 }),
      )
      // Canonical coordinates are points; the mocked deck scale is one.
      expect((footer.geometry.x + footer.geometry.width) / 0.75).toBeLessThanOrEqual(730)
      expect((title.geometry.x + title.geometry.width) / 0.75).toBeLessThanOrEqual(730)
    }
  })

  it('uses the returned deck while React state is still stale', async () => {
    const staleSlides = [blank()]
    let authoritativeSlides = staleSlides
    const addSlide = vi.fn(async ({ sourceIndex }: { sourceIndex: number }) => {
      if (authoritativeSlides.length >= 2) return null
      const next = authoritativeSlides.slice()
      next.splice(sourceIndex + 1, 0, blank())
      authoritativeSlides = next
      return { slides: next, index: sourceIndex + 1 }
    })
    ;(globalThis as unknown as { window: Record<string, unknown> }).window = {
      slidesApi: { addSlide },
    }
    const executePresentationOperation = vi.fn(async (request) => ({
      receipt: {
        status: 'applied' as const,
        transactionId: request.transactionId,
        resultingDeckRevision: `sha256:${'a'.repeat(64)}`,
        operationCount: request.operations.length,
      },
      authoritativeState: 'fresh' as const,
    }))
    const skill = createSlidesSkill({
      // Mirrors production: applyDeck schedules React state, while the render-owned
      // getter can remain stale until this async tool yields back to React.
      getSlides: () => staleSlides,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => undefined,
      applyDeck: () => undefined,
      executePresentationOperation,
      fitWidthPx: 1280,
    })

    const result = await skill.executeTool({
      id: 'stale-react-state',
      name: 'build_deck',
      input: {
        pages: [
          { title: 'A', body: ['B'] },
          { title: 'C', body: ['D'] },
        ],
      },
    })

    expect(result).toMatchObject({ mutated: true })
    expect(result.isError).not.toBe(true)
    expect(addSlide).toHaveBeenCalledOnce()
    expect(executePresentationOperation).toHaveBeenCalledTimes(2)
  })

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

  it('builds a varied themed deck and places searched imagery', async () => {
    let slides = [blank()]
    const transactions: Array<{ operations: Array<Record<string, unknown>> }> = []
    const insertImageUrl = vi.fn(async ({ slideIndex }: { slideIndex: number }) => ({
      sourceId: `image-${slideIndex}`,
      slide: slides[slideIndex],
    }))
    ;(globalThis as unknown as { window: Record<string, unknown> }).window = {
      slidesApi: {
        imageSearch: vi.fn(async () => ({
          images: [
            { imageUrl: 'https://images.example/hero.jpg', title: 'Hero' },
            { imageUrl: 'https://images.example/model.jpg', title: 'Model' },
          ],
          method: 'serper',
        })),
        addSlide: vi.fn(async ({ sourceIndex }: { sourceIndex: number }) => {
          const next = slides.slice()
          next.splice(sourceIndex + 1, 0, blank())
          return { slides: next, index: sourceIndex + 1 }
        }),
        insertImageUrl,
      },
    }
    const skill = createSlidesSkill({
      getSlides: () => slides,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: (index, slide) => {
        slides[index] = slide
      },
      applyDeck: (next) => {
        slides = next
      },
      executePresentationOperation: vi.fn(async (request) => {
        transactions.push(request as never)
        return {
          receipt: {
            status: 'applied' as const,
            transactionId: request.transactionId,
            resultingDeckRevision: `sha256:${'a'.repeat(64)}`,
            operationCount: request.operations.length,
          },
          authoritativeState: 'fresh' as const,
        }
      }),
      fitWidthPx: 1280,
    })

    await skill.executeTool({
      id: 'search-images',
      name: 'image_search',
      input: { query: 'artificial intelligence' },
    })

    const result = await skill.executeTool({
      id: 'designed-build',
      invocationId: 'designed-build-invocation',
      name: 'build_deck',
      input: {
        theme: { mode: 'dark', primary: '#0B1020', accent: '#66E3FF' },
        pages: [
          {
            layout: 'statement',
            kicker: 'LLM · 2026',
            title: '语言模型，正在变成新界面',
            body: ['从回答问题，到完成工作'],
            imageUrl: 'https://images.example/hero.jpg',
            imageAlt: '抽象的人工智能网络',
          },
          {
            layout: 'split_image',
            title: '它如何工作',
            body: ['海量训练形成模式', '上下文决定当下任务', '逐步生成并调用工具'],
            imageUrl: 'https://images.example/model.jpg',
            imageAlt: '神经网络可视化',
          },
          {
            layout: 'cards',
            title: '八类核心能力',
            body: ['理解', '生成', '行动', '检索', '分析', '规划', '校验', '协作'],
          },
        ],
      },
    })

    expect(result).toMatchObject({ mutated: true })
    expect(insertImageUrl).toHaveBeenCalledTimes(2)
    expect(insertImageUrl).toHaveBeenNthCalledWith(
      1,
      expect.objectContaining({ slideIndex: 0, url: 'https://images.example/hero.jpg' }),
    )
    const serialized = JSON.stringify(transactions)
    expect(serialized).toContain('#0B1020')
    expect(serialized).toContain('#66E3FF')
    expect(serialized).toContain('LLM · 2026')
    expect(serialized).toContain('deck-card-2-7')
    expect(transactions.map((item) => item.operations.length)).not.toEqual([4, 4, 4])
    expect(new Set(transactions.map((item) => JSON.stringify(item.operations))).size).toBe(3)
  })

  it('treats empty optional image fields as no image', async () => {
    let slides = [blank()]
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
      executePresentationOperation: vi.fn(async (request) => ({
        receipt: {
          status: 'applied' as const,
          transactionId: request.transactionId,
          resultingDeckRevision: `sha256:${'a'.repeat(64)}`,
          operationCount: request.operations.length,
        },
        authoritativeState: 'fresh' as const,
      })),
      fitWidthPx: 1280,
    })

    const result = await skill.executeTool({
      id: 'empty-images',
      name: 'build_deck',
      input: {
        theme: { mode: 'dark' },
        pages: [
          { layout: 'cover', title: 'A', body: ['B'], imageUrl: '', imageAlt: '' },
          { layout: 'cards', title: 'C', body: ['D'], imageUrl: '', imageAlt: '' },
        ],
      },
    })

    expect(result).toMatchObject({ mutated: true })
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

  it('rejects unsearched images and designed pages without an explicit layout', async () => {
    const executePresentationOperation = vi.fn()
    const skill = createSlidesSkill({
      getSlides: () => [blank()],
      getCurrent: () => 0,
      getSelectedIds: () => [],
      applySlide: () => undefined,
      applyDeck: () => undefined,
      executePresentationOperation,
      fitWidthPx: 1280,
    })
    const unsearched = await skill.executeTool({
      id: 'unsafe-image',
      name: 'build_deck',
      input: {
        theme: { mode: 'dark' },
        pages: [
          {
            layout: 'cover',
            title: 'A',
            body: ['B'],
            imageUrl: 'https://invented.example/image.jpg',
            imageAlt: 'Invented',
          },
          { layout: 'cards', title: 'C', body: ['1', '2', '3'] },
        ],
      },
    })
    const missingLayout = await skill.executeTool({
      id: 'missing-layout',
      name: 'build_deck',
      input: {
        theme: { mode: 'dark' },
        pages: [
          { layout: 'cover', title: 'A', body: ['B'] },
          { title: 'C', body: ['1', '2', '3', '4', '5', '6'] },
        ],
      },
    })
    expect(unsearched).toMatchObject({ isError: true, mutated: false })
    expect(missingLayout).toMatchObject({ isError: true, mutated: false })
    expect(executePresentationOperation).not.toHaveBeenCalled()
  })

  it('reports partial mutation truth when a late image insertion fails', async () => {
    let slides = [blank()]
    ;(globalThis as unknown as { window: Record<string, unknown> }).window = {
      slidesApi: {
        imageSearch: vi.fn(async () => ({
          images: [{ imageUrl: 'https://images.example/hero.jpg', title: 'Hero' }],
          method: 'serper',
        })),
        addSlide: vi.fn(async () => {
          slides = [...slides, blank()]
          return { slides, index: 1 }
        }),
        insertImageUrl: vi.fn(async () => null),
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
      executePresentationOperation: vi.fn(async (request) => ({
        receipt: {
          status: 'applied' as const,
          transactionId: request.transactionId,
          resultingDeckRevision: `sha256:${'a'.repeat(64)}`,
          operationCount: request.operations.length,
        },
        authoritativeState: 'fresh' as const,
      })),
      fitWidthPx: 1280,
    })
    await skill.executeTool({ id: 'search', name: 'image_search', input: { query: 'team' } })
    const result = await skill.executeTool({
      id: 'partial',
      invocationId: 'partial-invocation',
      name: 'build_deck',
      input: {
        theme: { mode: 'dark' },
        pages: [
          {
            layout: 'cover',
            title: 'A',
            body: ['B'],
            imageUrl: 'https://images.example/hero.jpg',
            imageAlt: 'Hero',
          },
          {
            layout: 'split_image',
            title: 'C',
            body: ['D'],
            imageUrl: 'https://images.example/hero.jpg',
            imageAlt: 'Hero',
          },
        ],
      },
    })
    expect(result).toMatchObject({ isError: true, mutated: true, stopToolBatch: true })
    expect(window.slidesApi.insertImageUrl).toHaveBeenCalledTimes(2)
  })

  it('unlocks repair tools after a partially applied build fails', async () => {
    let slides = [blank()]
    const deleteSlide = vi.fn(async (slideIndex: number) => {
      slides = slides.filter((_, index) => index !== slideIndex)
      return slides
    })
    ;(globalThis as unknown as { window: Record<string, unknown> }).window = {
      slidesApi: {
        addSlide: vi.fn(async () => {
          slides = [...slides, blank()]
          return { slides, index: 1 }
        }),
        deleteSlide,
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
      executePresentationOperation: vi.fn(async (request) => ({
        receipt: {
          status: 'unchanged' as const,
          transactionId: request.transactionId,
          code: 'write_not_applied' as const,
          operationCount: request.operations.length,
        },
        authoritativeState: 'fresh' as const,
      })),
      fitWidthPx: 1280,
    })

    await skill.executeTool({
      id: 'plan',
      name: 'plan_deck',
      input: {
        core_hook: 'A hook',
        style: 'A style',
        pages: [
          { title: 'A', brief: 'B', layout: 'cover' },
          { title: 'C', brief: 'D', layout: 'cards' },
        ],
      },
    })
    const failed = await skill.executeTool({
      id: 'failed-build',
      name: 'build_deck',
      input: {
        pages: [
          { title: 'A', body: ['B'] },
          { title: 'C', body: ['D'] },
        ],
      },
    })
    const repair = await skill.executeTool({
      id: 'repair',
      name: 'delete_slide',
      input: { slideIndex: 1 },
    })

    expect(failed).toMatchObject({ isError: true, mutated: true })
    expect(repair).toMatchObject({ mutated: true })
    expect(repair.output).not.toContain('Call build_deck once')
    expect(deleteSlide).toHaveBeenCalledOnce()
  })
})
