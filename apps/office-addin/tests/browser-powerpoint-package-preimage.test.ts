import { afterEach, describe, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { BrowserPowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'
import { presentationPackageDigest } from '../src/skills/powerpoint/powerpoint-package'

afterEach(() => vi.unstubAllGlobals())

async function host(initial: string[], changeOnLoad?: number) {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', '<p:sld/>')
  const base64 = await zip.generateAsync({ type: 'base64' })
  const remove = vi.fn(),
    insert = vi.fn()
  const items = initial.map((id) => ({
    id,
    load: vi.fn(),
    delete: remove,
    exportAsBase64: vi.fn(() => ({ value: base64 })),
  }))
  let loads = 0
  const slides = {
    items,
    load: vi.fn(() => {
      loads++
      if (loads === changeOnLoad) slides.items = [...items].reverse()
    }),
    getCount: vi.fn(() => ({ value: slides.items.length })),
    getItemAt: vi.fn((index: number) => slides.items[index]),
  }
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', requirements: { isSetSupported: vi.fn(() => true) } },
  })
  const sync = vi.fn()
  vi.stubGlobal('PowerPoint', {
    run: (callback: (value: unknown) => unknown) =>
      callback({ presentation: { slides, insertSlidesFromBase64: insert }, sync }),
  })
  return {
    adapter: new BrowserPowerPointAdapter(),
    base64,
    digest: await presentationPackageDigest(base64),
    slides,
    sync,
    insert,
    remove,
  }
}

describe('native XML complete page order preimage', () => {
  it('reads all slide identities beyond visual verification and backup budgets', async () => {
    const ids = Array.from({ length: 600 }, (_, index) => `slide-${index}`)
    const h = await host(ids)
    expect(await h.adapter.readSlideOrder()).toEqual(ids)
    expect(h.insert).not.toHaveBeenCalled()
    expect(h.remove).not.toHaveBeenCalled()
  })

  it.each([
    ['s1', 's3', 's2'],
    ['s1', 's2'],
    ['s1', 's2', 's3', 's4'],
  ])('refuses changed deck order/count before any import for %j', async (...ids) => {
    const h = await host(ids)
    await expect(
      h.adapter.replaceSlidePackage(0, h.base64, false, undefined, undefined, {
        slideId: 's1',
        packageDigest: h.digest,
        slideIds: ['s1', 's2', 's3'],
      }),
    ).rejects.toThrow('proposal_stale')
    expect(h.insert).not.toHaveBeenCalled()
    expect(h.remove).not.toHaveBeenCalled()
  })

  it('refuses page order racing after the initial read and before native mutation', async () => {
    const h = await host(['s1', 's2', 's3'], 1)
    await expect(
      h.adapter.replaceSlidePackage(0, h.base64, false, undefined, undefined, {
        slideId: 's1',
        packageDigest: h.digest,
        slideIds: ['s1', 's2', 's3'],
      }),
    ).rejects.toThrow('proposal_stale')
    expect(h.slides.load.mock.calls.length).toBeGreaterThanOrEqual(1)
    expect(h.insert).not.toHaveBeenCalled()
    expect(h.remove).not.toHaveBeenCalled()
  })

  it.each(['package', 'order'])('owns the caller preimage across SDK awaits (%s)', async (kind) => {
    const h = await host(['s1', 's3', 's2'])
    const preimage = {
      slideId: 's1',
      packageDigest: kind === 'package' ? '0'.repeat(64) : h.digest,
      slideIds: kind === 'order' ? ['s1', 's2', 's3'] : ['s1', 's3', 's2'],
    }
    h.sync.mockImplementationOnce(() => {
      preimage.packageDigest = h.digest
      preimage.slideIds.splice(0, preimage.slideIds.length, 's1', 's3', 's2')
    })
    await expect(
      h.adapter.replaceSlidePackage(0, h.base64, false, undefined, undefined, preimage),
    ).rejects.toThrow('proposal_stale')
    expect(h.insert).not.toHaveBeenCalled()
    expect(h.remove).not.toHaveBeenCalled()
  })

  it('refuses ambiguous duplicate SDK slide identities', async () => {
    const h = await host(['s1', 's1'])
    await expect(h.adapter.readSlideOrder()).rejects.toThrow()
    expect(h.insert).not.toHaveBeenCalled()
  })
})
