import { afterEach, describe, expect, it, vi } from 'vitest'
import { BrowserPowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'
afterEach(() => vi.unstubAllGlobals())
function setup() {
  let text = 'original',
    writes = 0
  const font = {
    load: vi.fn(),
    name: 'Arial',
    size: 20,
    color: '#000000',
    bold: true,
    italic: false,
    underline: 'none',
  }
  const range = {
    load: vi.fn(),
    getSubstring: vi.fn((start: number, length: number) => ({
      load: vi.fn(),
      font,
      get text() {
        return text.slice(start, start + length)
      },
      set text(value: string) {
        writes++
        text = text.slice(0, start) + value + text.slice(start + length)
      },
    })),
    get text() {
      return text
    },
    set text(value: string) {
      writes++
      text = value
    },
  }
  const shape = {
    id: 'shape-id',
    name: 'Title',
    type: 'TextBox',
    left: 0,
    top: 0,
    width: 200,
    height: 30,
    load: vi.fn(),
    textFrame: { textRange: range },
  }
  const shapes = { getItem: vi.fn(() => shape), load: vi.fn(), items: [shape] }
  const slide = { id: 'host-27', load: vi.fn(), shapes }
  const slides = {
    getItem: vi.fn(() => slide),
    getItemAt: vi.fn(() => {
      throw new Error('index forbidden')
    }),
  }
  const context = { presentation: { slides }, sync: vi.fn(async () => {}) }
  const supports = vi.fn(() => true)
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', requirements: { isSetSupported: supports } },
  })
  vi.stubGlobal('PowerPoint', {
    run: vi.fn(async (callback: (value: typeof context) => Promise<unknown>) => callback(context)),
  })
  return {
    adapter: new BrowserPowerPointAdapter(),
    shape,
    shapes,
    slide,
    slides,
    context,
    range,
    font,
    supports,
    writes: () => writes,
    setText: (value: string) => {
      text = value
    },
  }
}
describe('stable host-ID text access', () => {
  it('edits one equal-length native text subrange and checks surrounding text and font', async () => {
    const f = setup()
    const before = await f.adapter.readPresentationPageTextRange('host-27', 'shape-id', 0, 4)
    expect(before).toMatchObject({
      fullText: 'original',
      text: 'orig',
      font: { name: 'Arial', bold: true },
    })
    await f.adapter.editPresentationPageTextRange(before, 'edit')
    expect(f.range.text).toBe('editinal')
    expect(f.writes()).toBe(1)
    expect(f.range.getSubstring).toHaveBeenCalledWith(0, 4)
  })
  it('rejects stale range font, length changes and ambiguous Unicode offsets before writing', async () => {
    const f = setup()
    const before = await f.adapter.readPresentationPageTextRange('host-27', 'shape-id', 0, 4)
    await expect(f.adapter.editPresentationPageTextRange(before, 'longer')).rejects.toThrow(
      'invalid_tool_input',
    )
    f.font.bold = false
    await expect(f.adapter.editPresentationPageTextRange(before, 'edit')).rejects.toThrow(
      'office_concurrent_change',
    )
    f.setText('😀text')
    await expect(
      f.adapter.readPresentationPageTextRange('host-27', 'shape-id', 0, 2),
    ).rejects.toThrow('office_api_unsupported')
    expect(f.writes()).toBe(0)
  })
  it('reads and edits the same host page directly without consulting slide order', async () => {
    const { adapter, slides, shapes, range, supports } = setup()
    expect(await adapter.listPresentationPageShapes('host-27')).toMatchObject({
      slideId: 'host-27',
      shapes: [{ id: 'shape-id' }],
      shapesTruncated: false,
    })
    expect(shapes.load).toHaveBeenCalledWith(expect.objectContaining({ $top: 101 }))
    expect(await adapter.readPresentationPageText('host-27', 'shape-id')).toEqual({
      slideId: 'host-27',
      shapeId: 'shape-id',
      text: 'original',
      paragraphs: ['original'],
    })
    await adapter.editPresentationPageText('host-27', 'shape-id', 'edited', 'original')
    expect(range.text).toBe('edited')
    expect(slides.getItem).toHaveBeenCalledWith('host-27')
    expect(slides.getItemAt).not.toHaveBeenCalled()
    expect(supports).toHaveBeenCalledWith('PowerPointApi', '1.10')
  })
  it('reports a bounded shape list with explicit truncation', async () => {
    const { adapter, shapes, shape } = setup()
    shapes.items = Array.from({ length: 101 }, (_, i) => ({ ...shape, id: `shape-${i}` }))
    const result = await adapter.listPresentationPageShapes('host-27')
    expect(result.shapes).toHaveLength(100)
    expect(result.shapesTruncated).toBe(true)
  })
  it('refuses stale expected text and invalid or deleted targets without writing', async () => {
    const { adapter, shape, slide, slides, writes } = setup()
    await expect(
      adapter.editPresentationPageText('host-27', 'shape-id', 'edit', 'stale'),
    ).rejects.toThrow('office_concurrent_change')
    shape.id = 'other'
    await expect(adapter.readPresentationPageText('host-27', 'shape-id')).rejects.toThrow(
      'office_read_failed',
    )
    slide.id = 'other'
    await expect(adapter.listPresentationPageShapes('host-27')).rejects.toThrow(
      'office_read_failed',
    )
    slides.getItem.mockImplementation(() => {
      throw new Error('ItemNotFound')
    })
    await expect(
      adapter.editPresentationPageText('host-27', 'shape-id', 'edit', 'original'),
    ).rejects.toThrow()
    expect(writes()).toBe(0)
  })
  it('rejects unsupported API, oversized inputs, and oversized existing text', async () => {
    const { adapter, supports, setText, writes } = setup()
    supports.mockReturnValue(false)
    await expect(adapter.listPresentationPageShapes('host-27')).rejects.toThrow(
      'office_api_unsupported',
    )
    supports.mockReturnValue(true)
    await expect(adapter.readPresentationPageText('', 'shape-id')).rejects.toThrow(
      'invalid_tool_input',
    )
    await expect(adapter.readPresentationPageText('host-27', 'x'.repeat(257))).rejects.toThrow(
      'invalid_tool_input',
    )
    await expect(
      adapter.editPresentationPageText('host-27', 'shape-id', 'x'.repeat(12001), 'original'),
    ).rejects.toThrow('invalid_tool_input')
    await expect(
      adapter.editPresentationPageText('host-27', 'shape-id', 'edit', 'x'.repeat(12001)),
    ).rejects.toThrow('invalid_tool_input')
    setText('x'.repeat(12001))
    await expect(adapter.readPresentationPageText('host-27', 'shape-id')).rejects.toThrow(
      'office_read_failed',
    )
    await expect(
      adapter.editPresentationPageText('host-27', 'shape-id', 'edit', 'x'.repeat(12000)),
    ).rejects.toThrow('office_read_failed')
    expect(writes()).toBe(0)
  })
  it('cancels before dispatch but reconciles cancellation after the assignment committed', async () => {
    const { adapter, context, writes, range } = setup()
    const before = new AbortController()
    before.abort()
    await expect(
      adapter.editPresentationPageText('host-27', 'shape-id', 'edited', 'original', before.signal),
    ).rejects.toThrow('cancelled')
    expect(writes()).toBe(0)
    const during = new AbortController()
    context.sync.mockImplementation(async () => {
      if (writes()) during.abort()
    })
    await adapter.editPresentationPageText(
      'host-27',
      'shape-id',
      'edited',
      'original',
      during.signal,
    )
    expect(range.text).toBe('edited')
    expect(writes()).toBe(1)
  })
  it('does not assign when cancellation occurs during the immediate pre-write read', async () => {
    const { adapter, context, writes } = setup()
    const controller = new AbortController()
    let syncs = 0
    context.sync.mockImplementation(async () => {
      if (++syncs === 3) controller.abort()
    })
    await expect(
      adapter.editPresentationPageText(
        'host-27',
        'shape-id',
        'edited',
        'original',
        controller.signal,
      ),
    ).rejects.toThrow('cancelled')
    expect(writes()).toBe(0)
  })
  it('reconciles rejected write sync without overwriting concurrent user text', async () => {
    const { adapter, context, writes, setText, range } = setup()
    let failed = false
    context.sync.mockImplementation(async () => {
      if (writes() && !failed) {
        failed = true
        setText('user text')
        throw new Error('sync rejected')
      }
    })
    await expect(
      adapter.editPresentationPageText('host-27', 'shape-id', 'edited', 'original'),
    ).rejects.toThrow('office_concurrent_change')
    expect(range.text).toBe('user text')
    expect(writes()).toBe(1)
  })
})
