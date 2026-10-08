import { afterEach, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { BrowserPresentationMasterXmlAdapter } from '../src/skills/powerpoint/browser-presentation-master-xml-adapter'
afterEach(() => vi.unstubAllGlobals())
async function fixture(count = 3) {
  const zip = new JSZip()
  zip.file(
    'ppt/presentation.xml',
    '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="256" r:id="r1"/></p:sldIdLst></p:presentation>',
  )
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    '<Relationships><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>',
  )
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="p"/>')
  const base64 = await zip.generateAsync({ type: 'base64' }),
    adapter = new BrowserPresentationMasterXmlAdapter()
  const layout = (id: string, masterId: string) => ({ id, masterId, name: 'Duplicate layout' })
  const collection = (items: ReturnType<typeof layout>[]) => ({
    items,
    load: vi.fn(),
    getCount: () => ({ value: items.length }),
    getItem: (id: string) => items.find((item) => item.id === id)!,
  })
  const master = (id: string) => ({
    id,
    name: 'Duplicate master',
    layouts: collection([layout(`${id}-layout-1`, id), layout(`${id}-layout-2`, id)]),
  })
  const masters = {
    items: [master('master-1'), master('master-2')],
    load: vi.fn(),
    getCount: () => ({ value: masters.items.length }),
    getItem: (id: string) => masters.items.find((item) => item.id === id)!,
  }
  let queued: (() => void) | undefined,
    loseAck = false,
    hook: (() => void) | undefined
  const remove = vi.fn((id: string) => {
    queued = () => {
      slides.items = slides.items.filter((page) => page.id !== id)
    }
  })
  const apply = vi.fn((id: string, target: ReturnType<typeof layout>) => {
    queued = () => {
      const page = slides.getItem(id)
      page.slideMaster = { id: target.masterId }
      page.layout = { id: target.id }
    }
  })
  const page = (
    id: string,
    masterId = 'master-1',
    layoutId = `${masterId}-layout-1`,
    bytes = base64,
  ) => ({
    id,
    slideMaster: { id: masterId },
    layout: { id: layoutId },
    load: vi.fn(),
    exportAsBase64: vi.fn(() => ({ value: bytes })),
    delete: () => remove(id),
    applyLayout: (target: ReturnType<typeof layout>) => apply(id, target),
  })
  const slides = {
    items: Array.from({ length: count }, (_, i) =>
      page(`host-${i}`, i % 2 ? 'master-2' : 'master-1'),
    ),
    load: vi.fn(),
    getCount: () => ({ value: slides.items.length }),
    getItem: (id: string) => slides.items.find((item) => item.id === id)!,
  }
  const insert = vi.fn((bytes: string, options: { targetSlideId: string }) => {
    queued = () => {
      masters.items.push(master('actual-import-master'))
      slides.items.splice(
        slides.items.findIndex((page) => page.id === options.targetSlideId) + 1,
        0,
        page('actual-import-page', 'actual-import-master', 'actual-import-master-layout-1', bytes),
      )
    }
  })
  const sync = vi.fn(async () => {
    hook?.()
    if (queued) {
      const write = queued
      queued = undefined
      write()
      if (loseAck) throw Error('ACK lost')
    }
  })
  let api = true
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', platform: 'PC', requirements: { isSetSupported: () => api } },
  })
  vi.stubGlobal('PowerPoint', {
    run: (callback: (context: unknown) => unknown) =>
      callback({
        presentation: { slides, slideMasters: masters, insertSlidesFromBase64: insert },
        sync,
      }),
  })
  const preimage = await adapter.inspect()
  return {
    adapter,
    masters,
    slides,
    insert,
    remove,
    apply,
    sync,
    base64,
    preimage,
    request: { base64, sourceSlideId: 'host-0', packageSourceSlideId: '256#', preimage },
    loseAck: () => {
      loseAck = true
    },
    setHook: (fn: () => void) => {
      hook = fn
      sync.mockClear()
    },
    setApi: (value: boolean) => {
      api = value
    },
  }
}
const beforeWrite = async () => {},
  writeGuard = () => {}
it('captures full600 page proofs and native dependencies with duplicate names and exports every requested original', async () => {
  const f = await fixture(600),
    actual = await f.adapter.inspect(['host-0', 'host-599'])
  expect(actual.slideIds).toHaveLength(600)
  expect(actual.pages).toHaveLength(600)
  expect(actual.dependencies).toHaveLength(600)
  expect(actual.pages.filter((page) => page.base64).map((page) => page.slideId)).toEqual([
    'host-0',
    'host-599',
  ])
  expect(actual.dependencies[599]).toEqual({
    slideId: 'host-599',
    masterId: 'master-2',
    layoutId: 'master-2-layout-1',
  })
  expect(actual.masters.map((master) => master.name)).toEqual([
    'Duplicate master',
    'Duplicate master',
  ])
  expect(
    actual.masters.flatMap((master) => master.layouts.map((layout) => layout.layoutId)),
  ).toHaveLength(4)
  expect(actual.pages.every((page) => /^[a-f0-9]{64}$/.test(page.digest))).toBe(true)
})
it('stages without deleting and persists actual page/master/layout IDs separately from package IDs', async () => {
  const f = await fixture(),
    callback = vi.fn(async () => {})
  const actual = await f.adapter.stage(f.request, callback, beforeWrite, writeGuard)
  expect(actual).toEqual({
    slideId: 'actual-import-page',
    masterId: 'actual-import-master',
    layoutId: 'actual-import-master-layout-1',
  })
  expect(callback).toHaveBeenCalledExactlyOnceWith(actual)
  expect(f.insert).toHaveBeenCalledWith(f.base64, {
    targetSlideId: 'host-0',
    sourceSlideIds: ['256#'],
    formatting: 'KeepSourceFormatting',
  })
  expect(f.remove).not.toHaveBeenCalled()
  expect(f.apply).not.toHaveBeenCalled()
})
it('uses exact master/layout membership despite duplicate names and applies only one page per call', async () => {
  const f = await fixture()
  await f.adapter.applyLayout(
    {
      slideId: 'host-0',
      masterId: 'master-2',
      layoutId: 'master-2-layout-2',
      preimage: f.preimage,
    },
    beforeWrite,
    writeGuard,
  )
  expect(f.apply).toHaveBeenCalledExactlyOnceWith(
    'host-0',
    expect.objectContaining({ id: 'master-2-layout-2', masterId: 'master-2' }),
  )
  expect(f.slides.getItem('host-0').layout.id).toBe('master-2-layout-2')
  expect(f.slides.getItem('host-2').layout.id).toBe('master-1-layout-1')
  expect(f.insert).not.toHaveBeenCalled()
  expect(f.remove).not.toHaveBeenCalled()
})
it('removes one page as a separate independently guarded phase', async () => {
  const f = await fixture()
  await f.adapter.remove({ slideId: 'host-1', preimage: f.preimage }, beforeWrite, writeGuard)
  expect(f.remove).toHaveBeenCalledExactlyOnceWith('host-1')
  expect(f.slides.items.map((page) => page.id)).toEqual(['host-0', 'host-2'])
  expect(f.apply).not.toHaveBeenCalled()
  expect(f.insert).not.toHaveBeenCalled()
})
it.each(['master_name', 'layout_name', 'dependency', 'order'] as const)(
  'rejects %s drift during the business await before any write',
  async (mode) => {
    const f = await fixture()
    const business = async () => {
      await Promise.resolve()
      if (mode === 'master_name') f.masters.items[0]!.name = 'Changed'
      if (mode === 'layout_name') f.masters.items[0]!.layouts.items[0]!.name = 'Changed'
      if (mode === 'dependency') f.slides.items[0]!.layout.id = 'master-1-layout-2'
      if (mode === 'order') f.slides.items.reverse()
    }
    await expect(f.adapter.stage(f.request, async () => {}, business, writeGuard)).rejects.toThrow(
      'proposal_stale',
    )
    expect(f.insert).not.toHaveBeenCalled()
    expect(f.apply).not.toHaveBeenCalled()
    expect(f.remove).not.toHaveBeenCalled()
  },
)
it.each(['master_name', 'layout_name', 'dependency', 'order'] as const)(
  'reloads %s in the final native SDK read batch after asynchronous package hashes',
  async (mode) => {
    const f = await fixture()
    f.setHook(() => {
      if (f.sync.mock.calls.length !== 9) return
      if (mode === 'master_name') f.masters.items[0]!.name = 'Changed'
      if (mode === 'layout_name') f.masters.items[0]!.layouts.items[0]!.name = 'Changed'
      if (mode === 'dependency') f.slides.items[0]!.layout.id = 'master-1-layout-2'
      if (mode === 'order') f.slides.items.reverse()
    })
    await expect(
      f.adapter.stage(f.request, async () => {}, beforeWrite, writeGuard),
    ).rejects.toThrow('proposal_stale')
    expect(f.insert).not.toHaveBeenCalled()
    expect(
      f.masters.load.mock.calls.some(([fields]) => String(fields).includes('items/name')),
    ).toBe(true)
  },
)
it('checks the synchronous document/capability/CAS guard immediately after the final SDK await', async () => {
  const f = await fixture(),
    guard = vi.fn(() => {
      throw Error('presentation_document_changed')
    })
  await expect(
    f.adapter.applyLayout(
      {
        slideId: 'host-0',
        masterId: 'master-2',
        layoutId: 'master-2-layout-1',
        preimage: f.preimage,
      },
      beforeWrite,
      guard,
    ),
  ).rejects.toThrow('presentation_document_changed')
  expect(guard).toHaveBeenCalledOnce()
  expect(f.apply).not.toHaveBeenCalled()
})
it('copies requests and preimage before the first await', async () => {
  const f = await fixture()
  await f.adapter.stage(
    f.request,
    async () => {},
    async () => {
      f.request.sourceSlideId = 'host-1'
      f.request.packageSourceSlideId = 'wrong'
      f.request.preimage.masters[0]!.layouts.reverse()
      f.request.preimage.dependencies[0]!.layoutId = 'wrong'
    },
    writeGuard,
  )
  expect(f.insert.mock.calls[0]![1].targetSlideId).toBe('host-0')
})
it.each(['insert', 'delete', 'apply'] as const)(
  'retains unknown %s ACK without replay or inverse writes',
  async (mode) => {
    const f = await fixture(),
      callback = vi.fn(async () => {})
    f.loseAck()
    const action =
      mode === 'insert'
        ? f.adapter.stage(f.request, callback, beforeWrite, writeGuard)
        : mode === 'delete'
          ? f.adapter.remove({ slideId: 'host-0', preimage: f.preimage }, beforeWrite, writeGuard)
          : f.adapter.applyLayout(
              {
                slideId: 'host-0',
                masterId: 'master-2',
                layoutId: 'master-2-layout-1',
                preimage: f.preimage,
              },
              beforeWrite,
              writeGuard,
            )
    await expect(action).rejects.toThrow('office_state_uncertain')
    expect(f.insert).toHaveBeenCalledTimes(mode === 'insert' ? 1 : 0)
    expect(f.remove).toHaveBeenCalledTimes(mode === 'delete' ? 1 : 0)
    expect(f.apply).toHaveBeenCalledTimes(mode === 'apply' ? 1 : 0)
    expect(callback).not.toHaveBeenCalled()
  },
)
it('keeps the inserted page if actual-ID persistence fails', async () => {
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
  expect(f.slides.items.some((page) => page.id === 'actual-import-page')).toBe(true)
  expect(f.remove).not.toHaveBeenCalled()
})
it('refuses Mac and absent capability without a hidden fallback', async () => {
  const f = await fixture()
  f.setApi(false)
  await expect(f.adapter.inspect()).rejects.toThrow('office_api_unsupported')
  f.setApi(true)
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', platform: 'Mac', requirements: { isSetSupported: () => true } },
  })
  await expect(f.adapter.stage(f.request, async () => {}, beforeWrite, writeGuard)).rejects.toThrow(
    'office_api_unsupported',
  )
  expect(f.insert).not.toHaveBeenCalled()
})
it('refuses foreign native layout membership and cancellation before any host write', async () => {
  const f = await fixture()
  await expect(
    f.adapter.applyLayout(
      {
        slideId: 'host-0',
        masterId: 'master-2',
        layoutId: 'master-1-layout-1',
        preimage: f.preimage,
      },
      beforeWrite,
      writeGuard,
    ),
  ).rejects.toThrow('invalid_tool_input')
  const controller = new AbortController()
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
  expect(f.apply).not.toHaveBeenCalled()
})

it('reads only the requested last page in a 600-page document', async () => {
  const f = await fixture(600)
  f.slides.items.forEach((page) => page.exportAsBase64.mockClear())
  const result = await f.adapter.readPage('host-599')
  expect(result).toMatchObject({
    slideId: 'host-599',
    masterId: 'master-2',
    layoutId: 'master-2-layout-1',
    base64: f.base64,
  })
  expect(result.digest).toMatch(/^[a-f0-9]{64}$/)
  expect(f.slides.items.reduce((n, page) => n + page.exportAsBase64.mock.calls.length, 0)).toBe(1)
})
it('rejects single-page identity drift after package hashing', async () => {
  const f = await fixture()
  f.setHook(() => {
    if (f.sync.mock.calls.length === 3) f.slides.getItem('host-0').layout.id = 'foreign'
  })
  await expect(f.adapter.readPage('host-0')).rejects.toThrow('proposal_stale')
})
it('keeps exported bytes immutable across the final SDK await', async () => {
  const f = await fixture()
  const exported = { value: f.base64 }
  f.slides.getItem('host-0').exportAsBase64.mockReturnValue(exported)
  f.setHook(() => {
    if (f.sync.mock.calls.length === 3) exported.value = 'changed'
  })
  expect((await f.adapter.readPage('host-0')).base64).toBe(f.base64)
})
it('cancels a single-page read and refuses an oversized package', async () => {
  const f = await fixture()
  const abort = new AbortController()
  f.setHook(() => abort.abort())
  await expect(f.adapter.readPage('host-0', abort.signal)).rejects.toThrow('cancelled')
  f.setHook(() => {})
  f.slides
    .getItem('host-0')
    .exportAsBase64.mockReturnValue({ value: 'A'.repeat(Math.ceil((8 * 1024 * 1024) / 3) * 4 + 4) })
  await expect(f.adapter.readPage('host-0')).rejects.toThrow()
  expect(f.insert).not.toHaveBeenCalled()
  expect(f.remove).not.toHaveBeenCalled()
  expect(f.apply).not.toHaveBeenCalled()
})
