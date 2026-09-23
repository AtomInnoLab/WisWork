import { afterEach, describe, expect, it, vi } from 'vitest'
import { BrowserPowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'
afterEach(() => vi.unstubAllGlobals())
const before = { left: 10, top: 20, width: 100, height: 50 },
  target = { left: 30, top: 40, width: 120, height: 60 }
function setup() {
  const shape = { id: 'shape-1', ...before, load: vi.fn() }
  const shapes = { getItem: vi.fn(() => shape) }
  const slide = { id: 'host-29', shapes, load: vi.fn() }
  const slides = {
    getItem: vi.fn(() => slide),
    getItemAt: vi.fn(() => {
      throw new Error('wrong-index')
    }),
  }
  const context = { presentation: { slides }, sync: vi.fn(async () => {}) }
  const supports = vi.fn(() => true)
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', requirements: { isSetSupported: supports } },
  })
  vi.stubGlobal('PowerPoint', {
    run: vi.fn(async (cb: (ctx: typeof context) => Promise<unknown>) => cb(context)),
  })
  return {
    adapter: new BrowserPowerPointAdapter(),
    shape,
    slide,
    slides,
    shapes,
    context,
    supports,
  }
}
describe('host-ID page geometry adapter', () => {
  it('reads and changes four point values on the exact host ID', async () => {
    const { adapter, slides, shape, supports } = setup()
    expect(await adapter.readPresentationPageGeometry('host-29', 'shape-1')).toEqual({
      slideId: 'host-29',
      shapeId: 'shape-1',
      geometry: before,
    })
    await adapter.editPresentationPageGeometry('host-29', 'shape-1', target, before)
    expect(shape).toMatchObject(target)
    expect(slides.getItem).toHaveBeenCalledWith('host-29')
    expect(slides.getItemAt).not.toHaveBeenCalled()
    expect(supports).toHaveBeenCalledWith('PowerPointApi', '1.10')
  })
  it('rejects stale baselines, invalid geometry, wrong IDs and unsupported hosts before mutation', async () => {
    const { adapter, shape, slide, supports } = setup()
    await expect(
      adapter.editPresentationPageGeometry('host-29', 'shape-1', target, {
        ...before,
        left: 10.001,
      }),
    ).rejects.toThrow('office_concurrent_change')
    for (const geometry of [
      { ...target, width: -1 },
      { ...target, left: Infinity },
      { ...target, top: 100001 },
      { ...target, extra: 1 },
    ])
      await expect(
        adapter.editPresentationPageGeometry('host-29', 'shape-1', geometry, before),
      ).rejects.toThrow('invalid_tool_input')
    expect(shape).toMatchObject(before)
    slide.id = 'other'
    await expect(adapter.readPresentationPageGeometry('host-29', 'shape-1')).rejects.toThrow(
      'office_read_failed',
    )
    slide.id = 'host-29'
    shape.id = 'other'
    await expect(adapter.readPresentationPageGeometry('host-29', 'shape-1')).rejects.toThrow(
      'office_read_failed',
    )
    shape.id = 'shape-1'
    supports.mockReturnValue(false)
    await expect(adapter.readPresentationPageGeometry('host-29', 'shape-1')).rejects.toThrow(
      'office_api_unsupported',
    )
  })
  it('accepts host rounding within 0.01 points and zero-size line dimensions', async () => {
    const { adapter, shape, context } = setup()
    context.sync.mockImplementation(async () => {
      if (shape.left === target.left) shape.left += 0.005
    })
    await adapter.editPresentationPageGeometry('host-29', 'shape-1', target, before)
    await adapter.editPresentationPageGeometry(
      'host-29',
      'shape-1',
      { left: -10, top: 0, width: 0, height: 0 },
      { left: shape.left, top: shape.top, width: shape.width, height: shape.height },
    )
    expect(shape).toMatchObject({ width: 0, height: 0, left: -10 })
  })
  it('reconciles a rejected sync or cancellation only as success when all values applied', async () => {
    const { adapter, shape, context } = setup()
    const controller = new AbortController()
    let rejected = false
    context.sync.mockImplementation(async () => {
      if (shape.left === target.left && !rejected) {
        rejected = true
        controller.abort()
        throw new Error('sync rejected')
      }
    })
    await adapter.editPresentationPageGeometry(
      'host-29',
      'shape-1',
      target,
      before,
      controller.signal,
    )
    expect(shape).toMatchObject(target)
  })
  it.each(['before', 'partial', 'third', 'read-failed'] as const)(
    'classifies %s readback without rollback',
    async (state) => {
      const { adapter, shape, context } = setup()
      let dispatched = false
      context.sync.mockImplementation(async () => {
        if (shape.left === target.left && !dispatched) {
          dispatched = true
          if (state === 'before') Object.assign(shape, before)
          if (state === 'partial') shape.height = before.height
          if (state === 'third') shape.height = 999
          throw new Error('sync rejected')
        }
        if (dispatched && state === 'read-failed') throw new Error('read failed')
      })
      const code = {
        before: 'office_write_failed',
        partial: 'office_state_uncertain',
        third: 'office_concurrent_change',
        'read-failed': 'office_state_uncertain',
      }[state]
      await expect(
        adapter.editPresentationPageGeometry('host-29', 'shape-1', target, before),
      ).rejects.toThrow(code)
      if (state === 'partial') expect(shape).toMatchObject({ ...target, height: before.height })
      if (state === 'third') expect(shape.height).toBe(999)
    },
  )
  it('reports cancellation when reconciliation proves no values changed', async () => {
    const { adapter, shape, context } = setup()
    const controller = new AbortController()
    context.sync.mockImplementation(async () => {
      if (shape.left === target.left) {
        Object.assign(shape, before)
        controller.abort()
        throw new Error('not applied')
      }
    })
    await expect(
      adapter.editPresentationPageGeometry('host-29', 'shape-1', target, before, controller.signal),
    ).rejects.toThrow('cancelled')
    expect(shape).toMatchObject(before)
  })
  it('does not queue geometry writes if cancellation arrives during the pre-write read', async () => {
    const { adapter, shape, context } = setup()
    const controller = new AbortController()
    let calls = 0
    context.sync.mockImplementation(async () => {
      if (++calls === 2) controller.abort()
    })
    await expect(
      adapter.editPresentationPageGeometry('host-29', 'shape-1', target, before, controller.signal),
    ).rejects.toThrow('cancelled')
    expect(shape).toMatchObject(before)
  })
  it('reconciles partial property setter failure rather than leaking a false clean failure', async () => {
    const { adapter, shape } = setup()
    Object.defineProperty(shape, 'top', {
      get: () => before.top,
      set: () => {
        throw new Error('setter failed')
      },
      enumerable: true,
    })
    await expect(
      adapter.editPresentationPageGeometry('host-29', 'shape-1', target, before),
    ).rejects.toThrow('office_state_uncertain')
    expect(shape.left).toBe(target.left)
    expect(shape.width).toBe(before.width)
  })
  it('rejects invalid host geometry and cancellation before writing', async () => {
    const { adapter, shape } = setup()
    shape.width = NaN
    await expect(adapter.readPresentationPageGeometry('host-29', 'shape-1')).rejects.toThrow(
      'office_read_failed',
    )
    shape.width = before.width
    const controller = new AbortController()
    controller.abort()
    await expect(
      adapter.editPresentationPageGeometry('host-29', 'shape-1', target, before, controller.signal),
    ).rejects.toThrow('cancelled')
    expect(shape).toMatchObject(before)
  })
})
