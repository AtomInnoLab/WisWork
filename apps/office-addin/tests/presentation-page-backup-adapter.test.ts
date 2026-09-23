import { afterEach, expect, it, vi } from 'vitest'
import { BrowserPowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'
afterEach(() => vi.unstubAllGlobals())
function setup() {
  const slide = {
    id: 'host-27',
    load: vi.fn(),
    exportAsBase64: vi.fn(() => ({ value: 'UEsDBAAAAAA=' })),
  }
  const slides = {
    items: [{ id: 'before' }, { id: 'host-27' }, { id: 'after' }],
    load: vi.fn(),
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
    run: vi.fn(async (cb: (c: typeof context) => Promise<unknown>) => cb(context)),
  })
  return { adapter: new BrowserPowerPointAdapter(), slide, slides, context, supports }
}
it('exports the exact stable host page and bounded presentation order without selecting or writing', async () => {
  const f = setup()
  expect(await f.adapter.exportPresentationPagePackage('host-27')).toEqual({
    slideId: 'host-27',
    slideIds: ['before', 'host-27', 'after'],
    base64: 'UEsDBAAAAAA=',
  })
  expect(f.slides.getItem).toHaveBeenCalledWith('host-27')
  expect(f.slides.getItemAt).not.toHaveBeenCalled()
  expect(f.slides.load).toHaveBeenCalledWith({ $top: 513, id: true })
  expect(f.supports).toHaveBeenCalledWith('PowerPointApi', '1.8')
})
it('rejects page order changes during export, missing targets, and mismatched identities', async () => {
  const f = setup()
  f.slide.exportAsBase64.mockImplementation(() => {
    f.slides.items.reverse()
    return { value: 'UEsDBAAAAAA=' }
  })
  await expect(f.adapter.exportPresentationPagePackage('host-27')).rejects.toThrow(
    'office_concurrent_change',
  )
  await expect(f.adapter.exportPresentationPagePackage('missing')).rejects.toThrow(
    'office_read_failed',
  )
  f.slide.id = 'different'
  await expect(f.adapter.exportPresentationPagePackage('host-27')).rejects.toThrow(
    'office_read_failed',
  )
})
it('rejects oversized or duplicate page lists before exporting', async () => {
  const f = setup()
  f.slides.items.push({ id: 'host-27' })
  await expect(f.adapter.exportPresentationPagePackage('host-27')).rejects.toThrow(
    'office_read_failed',
  )
  f.slides.items = Array.from({ length: 513 }, (_, i) => ({ id: i ? `page-${i}` : 'host-27' }))
  await expect(f.adapter.exportPresentationPagePackage('host-27')).rejects.toThrow(
    'office_read_failed',
  )
  expect(f.slide.exportAsBase64).not.toHaveBeenCalled()
})
it('stops for cancellation, unsupported hosts, and oversized output', async () => {
  const f = setup(),
    controller = new AbortController()
  controller.abort()
  await expect(
    f.adapter.exportPresentationPagePackage('host-27', controller.signal),
  ).rejects.toThrow('cancelled')
  expect(f.slide.exportAsBase64).not.toHaveBeenCalled()
  f.supports.mockReturnValue(false)
  await expect(f.adapter.exportPresentationPagePackage('host-27')).rejects.toThrow(
    'office_api_unsupported',
  )
  f.supports.mockReturnValue(true)
  f.slide.exportAsBase64.mockReturnValue({
    value: 'A'.repeat(Math.ceil((8 * 1024 * 1024) / 3) * 4 + 4),
  })
  await expect(f.adapter.exportPresentationPagePackage('host-27')).rejects.toThrow(
    'office_read_failed',
  )
})
