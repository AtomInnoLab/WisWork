import { afterEach, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import {
  BrowserPresentationPackageEditAdapter,
  readPackageSourceSlideId,
} from '../src/skills/powerpoint/browser-presentation-package-edit-adapter'
afterEach(() => vi.unstubAllGlobals())
async function fixture(count = 2) {
  const zip = new JSZip()
  zip.file(
    'ppt/presentation.xml',
    '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>',
  )
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>',
  )
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="p"/>')
  const base64 = await zip.generateAsync({ type: 'base64' }),
    adapter = new BrowserPresentationPackageEditAdapter()
  let pending: (() => void) | undefined,
    loseAck = false,
    onSync: (() => void) | undefined
  const remove = vi.fn((target: string) => {
    pending = () => {
      slides.items = slides.items.filter((s) => s.id !== target)
    }
  })
  const item = (id: string) => ({
    id,
    exportAsBase64: () => ({ value: base64 }),
    delete: () => remove(id),
  })
  const slides = {
    items: Array.from({ length: count }, (_, i) => item(`host-${i}`)),
    load: vi.fn(),
    getCount: () => ({ value: slides.items.length }),
    getItem: (id: string) => slides.items.find((s) => s.id === id)!,
  }
  const insert = vi.fn((_base64: string, options: { targetSlideId: string }) => {
    pending = () => {
      slides.items.splice(
        slides.items.findIndex((s) => s.id === options.targetSlideId) + 1,
        0,
        item('new-host'),
      )
    }
  })
  const sync = vi.fn(async () => {
    onSync?.()
    if (pending) {
      const action = pending
      pending = undefined
      action()
      if (loseAck) throw Error('raw host ACK lost')
    }
  })
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', requirements: { isSetSupported: () => true } },
  })
  vi.stubGlobal('PowerPoint', {
    run: (callback: (ctx: unknown) => unknown) =>
      callback({ presentation: { slides, insertSlidesFromBase64: insert }, sync }),
  })
  const preimage = await adapter.inspect()
  const request = { base64, sourceSlideId: 'host-0', packageSourceSlideId: '256#', preimage }
  return {
    adapter,
    slides,
    base64,
    request,
    insert,
    remove,
    sync,
    setLoseAck: () => {
      loseAck = true
    },
    setOnSync: (fn: () => void) => {
      onSync = fn
    },
  }
}
const beforeWrite = async () => {},
  writeGuard = () => {}
it('reads all 600 page packages while retaining only explicitly requested bytes', async () => {
  const f = await fixture(600),
    result = await f.adapter.inspect(['host-599'])
  expect(result.slideIds).toHaveLength(600)
  expect(result.pages).toHaveLength(600)
  expect(result.pages.filter((p) => p.base64)).toEqual([
    expect.objectContaining({ slideId: 'host-599', base64: f.base64 }),
  ])
  expect(result.pages.every((p) => /^[a-f0-9]{64}$/.test(p.digest))).toBe(true)
})
it('inserts after the host ID using a distinct package source ID and records the unique adjacent ID', async () => {
  const f = await fixture(),
    callback = vi.fn(async () => {})
  expect(await readPackageSourceSlideId(f.base64)).toBe('256#')
  await expect(f.adapter.stage(f.request, callback, beforeWrite, writeGuard)).resolves.toEqual({
    slideId: 'new-host',
  })
  expect(f.insert).toHaveBeenCalledWith(f.base64, {
    targetSlideId: 'host-0',
    sourceSlideIds: ['256#'],
    formatting: 'KeepSourceFormatting',
  })
  expect(callback).toHaveBeenCalledWith('new-host')
  expect(f.remove).not.toHaveBeenCalled()
})
it('clones all request fields before awaiting the business guard', async () => {
  const f = await fixture()
  await f.adapter.stage(
    f.request,
    async () => {},
    async () => {
      f.request.base64 = 'invalid'
      f.request.sourceSlideId = 'host-1'
      f.request.preimage.slideIds.reverse()
    },
    writeGuard,
  )
  expect(f.insert.mock.calls[0]![1].targetSlideId).toBe('host-0')
})
it('rejects package or order drift during the business guard without a host write', async () => {
  const f = await fixture()
  await expect(
    f.adapter.stage(
      f.request,
      async () => {},
      async () => {
        f.slides.items.reverse()
      },
      writeGuard,
    ),
  ).rejects.toThrow('proposal_stale')
  expect(f.insert).not.toHaveBeenCalled()
})
it('rejects whole-page changes outside the target', async () => {
  const f = await fixture(),
    changed = new JSZip()
  changed.file('ppt/slides/slide1.xml', '<changed/>')
  const base64 = await changed.generateAsync({ type: 'base64' })
  f.slides.items[1]!.exportAsBase64 = () => ({ value: base64 })
  await expect(f.adapter.stage(f.request, async () => {}, beforeWrite, writeGuard)).rejects.toThrow(
    'proposal_stale',
  )
  expect(f.insert).not.toHaveBeenCalled()
})
it('checks the last SDK order after asynchronous hashes and then a synchronous document/CAS guard', async () => {
  const f = await fixture(),
    guard = vi.fn(() => {
      throw Error('presentation_document_changed')
    })
  await expect(f.adapter.stage(f.request, async () => {}, beforeWrite, guard)).rejects.toThrow(
    'presentation_document_changed',
  )
  expect(guard).toHaveBeenCalledOnce()
  expect(f.insert).not.toHaveBeenCalled()
})
it('rejects final order drift without insertion', async () => {
  const f = await fixture()
  let loads = 0
  f.slides.load.mockImplementation(() => {
    if (++loads === 2) f.slides.items.reverse()
  })
  await expect(f.adapter.stage(f.request, async () => {}, beforeWrite, writeGuard)).rejects.toThrow(
    'proposal_stale',
  )
  expect(f.insert).not.toHaveBeenCalled()
})
it('does not delete or claim a candidate after a lost insert ACK', async () => {
  const f = await fixture(),
    callback = vi.fn(async () => {})
  f.setLoseAck()
  await expect(f.adapter.stage(f.request, callback, beforeWrite, writeGuard)).rejects.toThrow(
    'office_state_uncertain',
  )
  expect(f.slides.items.map((s) => s.id)).toEqual(['host-0', 'new-host', 'host-1'])
  expect(callback).not.toHaveBeenCalled()
  expect(f.remove).not.toHaveBeenCalled()
})
it('keeps an inserted page if persistence callback fails', async () => {
  const f = await fixture()
  await expect(
    f.adapter.stage(
      f.request,
      async () => {
        throw Error('storage_failed')
      },
      beforeWrite,
      writeGuard,
    ),
  ).rejects.toThrow('storage_failed')
  expect(f.slides.items).toHaveLength(3)
  expect(f.remove).not.toHaveBeenCalled()
})
it('removes only the requested page and refuses to repeat a lost delete ACK', async () => {
  const f = await fixture()
  f.setLoseAck()
  await expect(
    f.adapter.remove({ slideId: 'host-0', preimage: f.request.preimage }, beforeWrite, writeGuard),
  ).rejects.toThrow('office_state_uncertain')
  expect(f.slides.items.map((s) => s.id)).toEqual(['host-1'])
  expect(f.remove).toHaveBeenCalledTimes(1)
})
it('honors cancellation before any host write', async () => {
  const f = await fixture(),
    controller = new AbortController()
  await expect(
    f.adapter.stage(
      f.request,
      async () => {},
      async () => controller.abort(),
      writeGuard,
      controller.signal,
    ),
  ).rejects.toThrow('cancelled')
  expect(f.insert).not.toHaveBeenCalled()
})
it('refuses ambiguous post-insert page order instead of claiming an ID', async () => {
  const f = await fixture(),
    callback = vi.fn(async () => {})
  f.setOnSync(() => {
    if (f.insert.mock.calls.length) f.slides.items.reverse()
  })
  await expect(f.adapter.stage(f.request, callback, beforeWrite, writeGuard)).rejects.toThrow(
    'office_state_uncertain',
  )
  expect(callback).not.toHaveBeenCalled()
  expect(f.remove).not.toHaveBeenCalled()
})
it('refuses a host ID passed as the package source identity before any write', async () => {
  const f = await fixture()
  f.request.packageSourceSlideId = 'host-0'
  await expect(f.adapter.stage(f.request, async () => {}, beforeWrite, writeGuard)).rejects.toThrow(
    'invalid_tool_input',
  )
  expect(f.insert).not.toHaveBeenCalled()
})
it('removes one page and checks the complete surviving order', async () => {
  const f = await fixture(3)
  await f.adapter.remove(
    { slideId: 'host-1', preimage: f.request.preimage },
    beforeWrite,
    writeGuard,
  )
  expect(f.slides.items.map((s) => s.id)).toEqual(['host-0', 'host-2'])
  expect(f.remove).toHaveBeenCalledExactlyOnceWith('host-1')
})
it('rejects a duplicate or incomplete snapshot before any write', async () => {
  const f = await fixture()
  f.request.preimage.slideIds[1] = 'host-0'
  await expect(f.adapter.stage(f.request, async () => {}, beforeWrite, writeGuard)).rejects.toThrow(
    'invalid_tool_input',
  )
  expect(f.insert).not.toHaveBeenCalled()
})
it('refuses unsupported Office capability instead of using a weaker fallback', async () => {
  const f = await fixture()
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', requirements: { isSetSupported: () => false } },
  })
  await expect(f.adapter.inspect()).rejects.toThrow('office_api_unsupported')
  expect(f.insert).not.toHaveBeenCalled()
})
