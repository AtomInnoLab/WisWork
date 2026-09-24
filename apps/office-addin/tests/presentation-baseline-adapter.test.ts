import { afterEach, describe, expect, it, vi } from 'vitest'
import { BrowserPresentationBaselineAdapter } from '../src/skills/powerpoint/browser-presentation-baseline-adapter.js'

function host() {
  const load = vi.fn()
  const shape = {
    id: 'native-shape',
    name: 'Title',
    type: 'TextBox',
    left: 1,
    top: 2,
    width: 30,
    height: 40,
    load,
    textFrame: {
      textRange: {
        text: 'Existing document',
        load,
        font: { name: null, size: 18, color: null, load },
      },
    },
  }
  const shapes = { items: [shape], load }
  const slide = {
    id: 'native-slide',
    load,
    shapes,
    slideMaster: { id: 'master', load },
    layout: { id: 'layout', load },
  }
  const slides = { items: [slide], load, getItem: vi.fn(() => slide) }
  const selectedSlides = { items: [slide], load }
  const selectedShapes = { items: [shape], load }
  const context = {
    sync: vi.fn(async () => {}),
    presentation: {
      slides,
      getSelectedSlides: () => selectedSlides,
      getSelectedShapes: () => selectedShapes,
      pageSetup: { slideWidth: 960, slideHeight: 540, load },
    },
  }
  const supports = vi.fn(() => true)
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', requirements: { isSetSupported: supports } },
  })
  const run = vi.fn(async (callback: (value: unknown) => unknown) => callback(context))
  vi.stubGlobal('PowerPoint', { run })
  return { shape, shapes, slide, slides, selectedSlides, selectedShapes, context, supports, run }
}
afterEach(() => vi.unstubAllGlobals())
describe('existing presentation baseline adapter', () => {
  it('reads native order, selection, dimensions and mixed fonts without generated artifacts', async () => {
    host()
    const adapter = new BrowserPresentationBaselineAdapter()
    expect(await adapter.readContext()).toEqual({
      slideIds: ['native-slide'],
      selectedSlideIds: ['native-slide'],
      selectedShapeIds: ['native-shape'],
      slideWidth: 960,
      slideHeight: 540,
    })
    expect(await adapter.readPage('native-slide')).toEqual({
      slideId: 'native-slide',
      masterId: 'master',
      layoutId: 'layout',
      shapes: [
        {
          id: 'native-shape',
          name: 'Title',
          type: 'TextBox',
          left: 1,
          top: 2,
          width: 30,
          height: 40,
          text: 'Existing document',
          font: { name: null, size: 18, color: null },
        },
      ],
    })
  })

  it('preserves selected order and leaves empty selection empty without a first-page fallback', async () => {
    const h = host()
    const other = { ...h.slide, id: 'other' }
    h.slides.items.push(other)
    h.selectedSlides.items = [other, h.slide]
    expect((await new BrowserPresentationBaselineAdapter().readContext()).selectedSlideIds).toEqual(
      ['other', 'native-slide'],
    )
    h.selectedSlides.items = []
    h.selectedShapes.items = []
    expect((await new BrowserPresentationBaselineAdapter().readContext()).selectedSlideIds).toEqual(
      [],
    )
  })
  it('omits unsupported dimensions and refuses unavailable selection APIs', async () => {
    const h = host()
    h.supports.mockImplementation((_name?: string, version?: string) => version !== '1.10')
    expect(await new BrowserPresentationBaselineAdapter().readContext()).not.toHaveProperty(
      'slideWidth',
    )
    h.supports.mockReturnValue(false)
    await expect(new BrowserPresentationBaselineAdapter().readContext()).rejects.toThrow(
      'office_api_unsupported',
    )
  })
  it('keeps complex shape metadata without touching text APIs', async () => {
    const h = host()
    for (const type of ['Group', 'Table', 'Chart', 'Picture', 'Unsupported']) {
      h.shape.type = type
      Object.defineProperty(h.shape, 'textFrame', {
        configurable: true,
        get: () => {
          throw new Error('complex text accessed')
        },
      })
      const page = await new BrowserPresentationBaselineAdapter().readPage('native-slide')
      expect(page.shapes[0]).toMatchObject({ id: 'native-shape', type })
      expect(page.shapes[0]).not.toHaveProperty('text')
      expect(page.shapes[0]).not.toHaveProperty('font')
    }
  })
  it.each(['slide count', 'shape count', 'single text', 'total text', 'bytes'])(
    'rejects %s overflow without truncating',
    async (reason) => {
      const h = host()
      const adapter = new BrowserPresentationBaselineAdapter()
      if (reason === 'slide count') {
        h.slides.items = Array.from({ length: 501 }, (_, i) => ({ ...h.slide, id: String(i) }))
        await expect(adapter.readContext()).rejects.toThrow('presentation_baseline_limit_exceeded')
        return
      }
      if (reason === 'shape count')
        h.shapes.items = Array.from({ length: 101 }, (_, i) => ({ ...h.shape, id: String(i) }))
      if (reason === 'single text') h.shape.textFrame.textRange.text = 'x'.repeat(12_001)
      if (reason === 'total text' || reason === 'bytes') {
        h.shape.textFrame.textRange.text = (reason === 'bytes' ? '汉' : 'x').repeat(12_000)
        h.shapes.items = Array.from({ length: reason === 'bytes' ? 8 : 11 }, (_, i) => ({
          ...h.shape,
          id: String(i),
        }))
      }
      await expect(adapter.readPage('native-slide')).rejects.toThrow(
        'presentation_baseline_limit_exceeded',
      )
    },
  )
  it.each([
    'duplicate slide',
    'duplicate shape',
    'unknown selected slide',
    'missing page',
    'invalid geometry',
  ])('refuses %s host data', async (reason) => {
    const h = host()
    const adapter = new BrowserPresentationBaselineAdapter()
    if (reason === 'duplicate slide') h.slides.items.push(h.slide)
    if (reason === 'unknown selected slide')
      h.selectedSlides.items = [{ ...h.slide, id: 'missing' }]
    if (reason === 'duplicate slide' || reason === 'unknown selected slide') {
      await expect(adapter.readContext()).rejects.toThrow('office_read_failed')
      return
    }
    if (reason === 'duplicate shape') h.shapes.items.push(h.shape)
    if (reason === 'missing page') h.slide.id = 'other'
    if (reason === 'invalid geometry') h.shape.left = Number.NaN
    await expect(adapter.readPage('native-slide')).rejects.toThrow('office_read_failed')
  })
  it('does not return a result when canceled before or during host synchronization', async () => {
    const h = host()
    const controller = new AbortController()
    controller.abort()
    await expect(
      new BrowserPresentationBaselineAdapter().readContext(controller.signal),
    ).rejects.toThrow('cancelled')
    expect(h.run).not.toHaveBeenCalled()
    const during = new AbortController()
    h.context.sync.mockImplementation(async () => {
      during.abort()
    })
    await expect(
      new BrowserPresentationBaselineAdapter().readPage('native-slide', during.signal),
    ).rejects.toThrow('cancelled')
  })
  it('reads title and body placeholders and observes later text changes', async () => {
    const h = host()
    h.shape.type = 'Placeholder'
    Object.assign(h.shape, {
      getTextFrameOrNullObject: () => ({ ...h.shape.textFrame, isNullObject: false }),
    })
    const adapter = new BrowserPresentationBaselineAdapter()
    expect((await adapter.readPage('native-slide')).shapes[0]).toMatchObject({
      text: 'Existing document',
      font: { name: null, size: 18, color: null },
    })
    h.shape.textFrame.textRange.text = 'Changed body'
    expect((await adapter.readPage('native-slide')).shapes[0]?.text).toBe('Changed body')
  })
  it('does not touch text ranges of picture placeholders', async () => {
    const h = host()
    h.shape.type = 'Placeholder'
    Object.assign(h.shape, {
      getTextFrameOrNullObject: () => ({
        isNullObject: true,
        get textRange() {
          throw new Error('picture text accessed')
        },
      }),
    })
    expect(
      (await new BrowserPresentationBaselineAdapter().readPage('native-slide')).shapes[0],
    ).not.toHaveProperty('text')
  })
  it('refuses unreadable placeholders on older hosts instead of claiming a complete baseline', async () => {
    const h = host()
    h.shape.type = 'Placeholder'
    h.supports.mockImplementation((_name?: string, version?: string) => version !== '1.10')
    await expect(new BrowserPresentationBaselineAdapter().readPage('native-slide')).rejects.toThrow(
      'office_api_unsupported',
    )
  })
})
