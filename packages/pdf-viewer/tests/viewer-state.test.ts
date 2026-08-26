import { describe, expect, it, vi } from 'vitest'
import {
  acceptDocumentLoad,
  beginDocumentLoad,
  createViewerResourceScope,
  createViewerState,
  setViewerPage,
  setViewerScale,
  assertDocumentPageBudget,
  computeCanvasBudget,
  createPageResourceScope,
  getPageWindow,
  getDocumentPages,
  nearestPageToViewport,
  pdfPointToViewportCss,
  viewportClientToPdfPoint,
  resetViewerPageForSource,
  runPageRenderSafely,
  clampViewerPageToDocument,
} from '../src/index.js'

describe('shared read-only PDF viewer state', () => {
  it('ignores stale document loads and clamps page and zoom', () => {
    let state = createViewerState()
    const first = beginDocumentLoad(state, 'first')
    const second = beginDocumentLoad(first.state, 'second')
    state = acceptDocumentLoad(second.state, first.token, { pageCount: 20 })
    expect(state.document).toBeNull()
    state = acceptDocumentLoad(state, second.token, { pageCount: 20 })
    state = setViewerPage(state, 99)
    state = setViewerScale(state, 20)
    expect(state.page).toBe(20)
    expect(state.scale).toBe(4)
  })

  it('cancels render tasks, destroys documents, and revokes object URLs exactly once', async () => {
    const cancel = vi.fn()
    const destroy = vi.fn(async () => undefined)
    const revoke = vi.fn()
    const scope = createViewerResourceScope({ revokeObjectUrl: revoke })
    scope.trackRenderTask({ cancel })
    scope.trackDocument({ destroy })
    scope.trackObjectUrl('blob:pdf')
    await scope.dispose()
    await scope.dispose()
    expect(cancel).toHaveBeenCalledOnce()
    expect(destroy).toHaveBeenCalledOnce()
    expect(revoke).toHaveBeenCalledOnce()
  })

  it('rejects excessive page counts, windows page DOM, and caps canvas allocation', () => {
    expect(() => assertDocumentPageBudget(2_001)).toThrow(/page limit/i)
    const budget = computeCanvasBudget(8_000, 8_000, 2)
    expect(budget.width).toBeLessThanOrEqual(8_192)
    expect(budget.height).toBeLessThanOrEqual(8_192)
    expect(budget.width * budget.height).toBeLessThanOrEqual(8_000_000)
    expect(() => computeCanvasBudget(Number.POSITIVE_INFINITY, 100, 2)).toThrow(/viewport/i)
    expect(() => computeCanvasBudget(1e12, 1e12, 2)).toThrow(/viewport/i)
    expect(getPageWindow(500, 1_000)).toEqual([499, 500, 501])
  })

  it('keeps every page scroll-addressable while choosing the nearest visible page', () => {
    expect(getDocumentPages(6)).toEqual([1, 2, 3, 4, 5, 6])
    expect(
      nearestPageToViewport([
        { page: 1, top: -700, bottom: 80 },
        { page: 2, top: 100, bottom: 880 },
        { page: 3, top: 900, bottom: 1_680 },
      ]),
    ).toBe(2)
  })

  it('resets page on source change and clamps it when a smaller document is accepted', () => {
    expect(resetViewerPageForSource(1_000)).toBe(1)
    expect(clampViewerPageToDocument(1_000, 10)).toBe(10)
  })

  it('contains asynchronous page setup failures and reports them locally', async () => {
    const onError = vi.fn()
    await expect(
      runPageRenderSafely(async () => {
        throw new Error('oversized page')
      }, onError),
    ).resolves.toBe(false)
    expect(onError).toHaveBeenCalledWith('oversized page')
  })

  it('cancels canvas and text rendering and cleans the page on unmount', () => {
    const render = { cancel: vi.fn() }
    const text = { cancel: vi.fn() }
    const page = { cleanup: vi.fn() }
    const scope = createPageResourceScope()
    scope.setRenderTask(render)
    scope.setTextLayer(text)
    scope.setPage(page)
    scope.dispose()
    scope.dispose()
    expect(render.cancel).toHaveBeenCalledOnce()
    expect(text.cancel).toHaveBeenCalledOnce()
    expect(page.cleanup).toHaveBeenCalledOnce()
  })

  it.each([0, 90, 180, 270])(
    'round trips reverse SyncTeX and marker coordinates at %i°',
    (rotation) => {
      const forward = (x: number, y: number): [number, number] => {
        switch (rotation) {
          case 0:
            return [x * 2, (200 - y) * 2]
          case 90:
            return [y * 2, x * 2]
          case 180:
            return [(100 - x) * 2, y * 2]
          default:
            return [(200 - y) * 2, (100 - x) * 2]
        }
      }
      const inverse = (left: number, top: number): [number, number] => {
        switch (rotation) {
          case 0:
            return [left / 2, 200 - top / 2]
          case 90:
            return [top / 2, left / 2]
          case 180:
            return [100 - left / 2, top / 2]
          default:
            return [100 - top / 2, 200 - left / 2]
        }
      }
      const viewport = {
        width: rotation % 180 === 0 ? 200 : 400,
        height: rotation % 180 === 0 ? 400 : 200,
        convertToViewportPoint: forward,
        convertToPdfPoint: inverse,
      }
      const marker = pdfPointToViewportCss(viewport, 25, 75)
      const reverse = viewportClientToPdfPoint(
        { left: 10, top: 20, width: viewport.width, height: viewport.height },
        viewport,
        marker.left + 10,
        marker.top + 20,
      )
      expect(reverse?.x).toBeCloseTo(25)
      expect(reverse?.y).toBeCloseTo(75)
    },
  )
})
