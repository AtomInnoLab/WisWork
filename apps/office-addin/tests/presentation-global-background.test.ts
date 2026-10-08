import { afterEach, expect, it, vi } from 'vitest'
import { BrowserPowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'

afterEach(() => vi.unstubAllGlobals())

it('keeps content collisions while excluding an explicit full-slide background from deck verification', async () => {
  const shapes = {
    load: vi.fn(),
    items: [
      {
        id: 'bg',
        name: 'Background',
        type: 'GeometricShape',
        left: 0,
        top: 0,
        width: 960,
        height: 540,
      },
      { id: 'a', name: 'Title', type: 'TextBox', left: 20, top: 20, width: 300, height: 80 },
      { id: 'b', name: 'Subtitle', type: 'TextBox', left: 50, top: 50, width: 300, height: 80 },
    ],
  }
  const slides = { load: vi.fn(), items: [{ id: 'slide-1', shapes }] }
  const context = {
    presentation: { slides, pageSetup: { load: vi.fn(), slideWidth: 960, slideHeight: 540 } },
    sync: vi.fn(async () => {}),
  }
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', requirements: { isSetSupported: () => true } },
  })
  vi.stubGlobal('PowerPoint', {
    run: (callback: (runtimeContext: typeof context) => Promise<unknown>) => callback(context),
  })
  const result = await new BrowserPowerPointAdapter().verifySlides()
  expect(result.slides[0]?.overlaps).toEqual([
    { shapeAId: 'a', shapeBId: 'b', overlapX: 270, overlapY: 50 },
  ])
})
