import { PNG } from 'pngjs'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { BrowserPowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'
const png =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII='
afterEach(() => vi.unstubAllGlobals())
function setup(items: Record<string, unknown>[] = []) {
  const shapes = { load: vi.fn(), items }
  const slide = {
    id: 'host-page-25',
    load: vi.fn(),
    shapes,
    getImageAsBase64: vi.fn(() => ({ value: png })),
  }
  const slides = {
    getItem: vi.fn(() => slide),
    getItemAt: vi.fn(() => {
      throw new Error('must address by host ID')
    }),
    load: vi.fn(),
  }
  const pageSetup = { load: vi.fn(), slideWidth: 960, slideHeight: 540 }
  const context = { presentation: { slides, pageSetup }, sync: vi.fn(async () => {}) }
  const supports = vi.fn(() => true)
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', requirements: { isSetSupported: supports } },
  })
  vi.stubGlobal('PowerPoint', {
    run: vi.fn(async (cb: (c: typeof context) => Promise<unknown>) => cb(context)),
  })
  return {
    slide,
    slides,
    shapes,
    pageSetup,
    context,
    supports,
    adapter: new BrowserPowerPointAdapter(),
  }
}
const shape = (id: string, left = 0, top = 0, width = 10, height = 10) => ({
  id,
  name: id,
  type: 'GeometricShape',
  left,
  top,
  width,
  height,
})
describe('exact imported PowerPoint page inspection', () => {
  it('reads an exact host page beyond slide twenty with real geometry and screenshot', async () => {
    const { adapter, slides, shapes, slide, supports } = setup([
      shape('a', -1, 0, 10, 20),
      shape('b', 2, 2, 10, 10),
      shape('c', 955, 535, 10, 10),
    ])
    const result = await adapter.inspectPresentationPage('host-page-25')
    expect(slides.getItem).toHaveBeenCalledWith('host-page-25')
    expect(slides.load).not.toHaveBeenCalled()
    expect(slides.getItemAt).not.toHaveBeenCalled()
    expect(supports).toHaveBeenCalledWith('PowerPointApi', '1.10')
    expect(shapes.load).toHaveBeenCalledWith(expect.objectContaining({ $top: 101 }))
    expect(slide.getImageAsBase64).toHaveBeenCalledWith({ width: 960 })
    expect(result).toMatchObject({
      slideId: 'host-page-25',
      slideWidth: 960,
      slideHeight: 540,
      shapesTruncated: false,
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: png },
    })
    expect(result.overflows).toEqual([
      { shapeId: 'a', edge: 'left', overflowBy: 1 },
      { shapeId: 'c', edge: 'right', overflowBy: 5 },
      { shapeId: 'c', edge: 'bottom', overflowBy: 5 },
    ])
    expect(result.overlaps).toEqual([{ shapeAId: 'a', shapeBId: 'b', overlapX: 7, overlapY: 10 }])
  })
  it('reduces dense screenshots to the model image budget and rejects unbounded output', async () => {
    const { adapter, slide } = setup()
    const dense = new PNG({ width: 400, height: 200 })
    let seed = 7
    for (let i = 0; i < dense.data.length; i++) {
      seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0
      dense.data[i] = seed >>> 24
    }
    const large = PNG.sync.write(dense).toString('base64')
    expect(Buffer.from(large, 'base64').length).toBeGreaterThan(64 * 1024)
    slide.getImageAsBase64.mockReturnValueOnce({ value: large }).mockReturnValue({ value: png })
    expect((await adapter.inspectPresentationPage('host-page-25')).screenshot.base64).toBe(png)
    expect(slide.getImageAsBase64.mock.calls).toEqual([[{ width: 960 }], [{ width: 640 }]])
    slide.getImageAsBase64.mockClear().mockReturnValue({ value: large })
    await expect(adapter.inspectPresentationPage('host-page-25')).rejects.toThrow(
      'office_image_too_large',
    )
    expect(slide.getImageAsBase64.mock.calls).toEqual(
      [960, 640, 480, 320, 240].map((width) => [{ width }]),
    )
  })
  it('explicitly marks shape and overlap truncation instead of returning a clean result', async () => {
    const { adapter } = setup(Array.from({ length: 101 }, (_, i) => shape(String(i))))
    const result = await adapter.inspectPresentationPage('host-page-25')
    expect(result.shapes).toHaveLength(100)
    expect(result.shapesTruncated).toBe(true)
    expect(result.overlaps).toHaveLength(1000)
    expect(result.overlapsTruncated).toBe(true)
  })
  it('fails closed for missing or mismatched host page IDs', async () => {
    const { adapter, slides, slide } = setup()
    slide.id = 'other'
    await expect(adapter.inspectPresentationPage('host-page-25')).rejects.toThrow(
      'office_read_failed',
    )
    slides.getItem.mockImplementation(() => {
      throw new Error('ItemNotFound')
    })
    await expect(adapter.inspectPresentationPage('host-page-25')).rejects.toThrow()
    await expect(adapter.inspectPresentationPage('x'.repeat(257))).rejects.toThrow(
      'invalid_tool_input',
    )
  })
  it('rejects unsupported API and cancellation before or after the host sync', async () => {
    const { adapter, supports, context } = setup()
    supports.mockReturnValue(false)
    await expect(adapter.inspectPresentationPage('host-page-25')).rejects.toThrow(
      'office_api_unsupported',
    )
    supports.mockReturnValue(true)
    const ac = new AbortController()
    ac.abort()
    await expect(adapter.inspectPresentationPage('host-page-25', ac.signal)).rejects.toThrow(
      'cancelled',
    )
    const during = new AbortController()
    context.sync.mockImplementation(async () => {
      during.abort()
    })
    await expect(adapter.inspectPresentationPage('host-page-25', during.signal)).rejects.toThrow(
      'cancelled',
    )
  })
  it('rejects invalid host geometry and screenshots', async () => {
    const { adapter, pageSetup, shapes, slide } = setup([shape('a')])
    pageSetup.slideWidth = NaN
    await expect(adapter.inspectPresentationPage('host-page-25')).rejects.toThrow(
      'office_read_failed',
    )
    pageSetup.slideWidth = 960
    shapes.items[0]!.width = Infinity
    await expect(adapter.inspectPresentationPage('host-page-25')).rejects.toThrow(
      'office_read_failed',
    )
    shapes.items[0]!.width = 10
    for (const value of ['abc', 'not a PNG', png + '\n', 'A'.repeat(3 * 1024 * 1024)]) {
      slide.getImageAsBase64.mockReturnValue({ value })
      await expect(adapter.inspectPresentationPage('host-page-25')).rejects.toThrow(
        'office_read_failed',
      )
    }
  })
  it('rejects missing screenshot support and out-of-range screenshot dimensions', async () => {
    const { adapter, slide } = setup()
    const image = Buffer.from(png, 'base64')
    image.writeUInt32BE(8193, 16)
    slide.getImageAsBase64.mockReturnValue({ value: image.toString('base64') })
    await expect(adapter.inspectPresentationPage('host-page-25')).rejects.toThrow(
      'office_read_failed',
    )
    Object.assign(slide, { getImageAsBase64: undefined })
    await expect(adapter.inspectPresentationPage('host-page-25')).rejects.toThrow(
      'office_api_unsupported',
    )
  })
  it('rejects ambiguous or negative shape geometry while preserving zero-height lines', async () => {
    const { adapter, shapes } = setup([shape('same'), shape('same')])
    await expect(adapter.inspectPresentationPage('host-page-25')).rejects.toThrow(
      'office_read_failed',
    )
    shapes.items = [shape('a', 0, 0, -1, 1)]
    await expect(adapter.inspectPresentationPage('host-page-25')).rejects.toThrow(
      'office_read_failed',
    )
    shapes.items = [shape('line', 0, 0, 10, 0)]
    expect((await adapter.inspectPresentationPage('host-page-25')).shapes[0]!.height).toBe(0)
  })
})
