import JSZip from 'jszip'
import { afterEach, expect, it, vi } from 'vitest'
import { BrowserPresentationPageReplacementAdapter } from '../src/skills/powerpoint/browser-presentation-page-replacement-adapter'
import { presentationPackageDigest } from '../src/skills/powerpoint/powerpoint-package'
import type { PresentationPageReplacement } from '../src/skills/powerpoint/presentation-page-replacement-record'

afterEach(() => vi.unstubAllGlobals())
async function pkg(text: string, compression: 'STORE' | 'DEFLATE' = 'STORE') {
  return new JSZip()
    .file('ppt/slides/slide1.xml', text)
    .generateAsync({ type: 'base64', compression })
}
async function setup() {
  const original = await pkg('original'),
    replacement = await pkg('replacement')
  const packages: Record<string, string> = { before: original, old: original, after: original }
  let ids = ['before', 'old', 'after']
  const remove = vi.fn((id: string) => {
    ids = ids.filter((item) => item !== id)
  })
  const slides = {
    get items() {
      return ids.map((id) => ({ id }))
    },
    load: vi.fn(),
    getItem: vi.fn((id: string) => ({
      id,
      load: vi.fn(),
      exportAsBase64: () => ({ value: packages[id] }),
      delete: () => remove(id),
    })),
  }
  const insert = vi.fn((base64: string) => {
    ids.splice(ids.indexOf('old') + 1, 0, 'new')
    packages.new = base64
  })
  const context = {
    presentation: { slides, insertSlidesFromBase64: insert },
    sync: vi.fn(async () => {}),
  }
  const support = vi.fn(() => true)
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', requirements: { isSetSupported: support } },
  })
  vi.stubGlobal('PowerPoint', {
    run: vi.fn(async (fn: (c: typeof context) => Promise<unknown>) => fn(context)),
  })
  const record = {
    version: 1,
    state: 'pending',
    oldSlideId: 'old',
    beforeSlideIds: [...ids],
    sourceSlideId: '256#',
    originalPackageDigest: await presentationPackageDigest(original),
    replacementPackageDigest: await presentationPackageDigest(replacement),
  } as PresentationPageReplacement
  return {
    adapter: new BrowserPresentationPageReplacementAdapter(),
    record,
    original,
    replacement,
    packages,
    context,
    insert,
    remove,
    support,
    setIds: (value: string[]) => {
      ids = value
    },
  }
}
it('fingerprints uncompressed entry content independent of ZIP compression; rejects unsafe and cancelled input', async () => {
  expect(await presentationPackageDigest(await pkg('same'))).toBe(
    await presentationPackageDigest(await pkg('same', 'DEFLATE')),
  )
  expect(await presentationPackageDigest(await pkg('same'))).not.toBe(
    await presentationPackageDigest(await pkg('different')),
  )
  await expect(presentationPackageDigest('bad')).rejects.toThrow('invalid_tool_input')
  await expect(presentationPackageDigest(await pkg('same'), AbortSignal.abort())).rejects.toThrow(
    'cancelled',
  )
})
it('stages a single page after the old page, journals identity before validation, and discards only that page', async () => {
  const f = await setup(),
    guard = vi.fn(),
    onInserted = vi.fn(async (id: string) => {
      expect(id).toBe('new')
      f.record = { ...f.record, state: 'inserted', newSlideId: id }
    })
  expect((await f.adapter.inspect(f.record)).status).toBe('baseline')
  await f.adapter.stage(f.record, f.replacement, onInserted, guard)
  expect(f.insert).toHaveBeenCalledWith(f.replacement, {
    targetSlideId: 'old',
    sourceSlideIds: ['256#'],
    formatting: 'KeepSourceFormatting',
  })
  expect(onInserted).toHaveBeenCalledOnce()
  expect((await f.adapter.inspect(f.record)).status).toBe('staged')
  expect(f.remove).not.toHaveBeenCalled()
  await f.adapter.discard({ ...f.record, state: 'discard_pending' }, guard)
  expect(f.remove).toHaveBeenCalledExactlyOnceWith('new')
  await f.adapter.discard({ ...f.record, state: 'discard_pending' }, guard)
  expect(f.remove).toHaveBeenCalledTimes(1)
})
it('never retries or deletes on partial insertion, failed identity receipt, or changed imported content', async () => {
  for (const mode of ['partial', 'receipt', 'content']) {
    const f = await setup(),
      journal = vi.fn(async () => {
        if (mode === 'receipt') throw new Error('receipt_failed')
      })
    if (mode === 'partial')
      f.insert.mockImplementation(() => {
        f.setIds(['before', 'old', 'new', 'after'])
        throw new Error('partial_write')
      })
    if (mode === 'content')
      f.insert.mockImplementation(() => {
        f.setIds(['before', 'old', 'new', 'after'])
        f.packages.new = f.original
      })
    await expect(f.adapter.stage(f.record, f.replacement, journal, vi.fn())).rejects.toThrow()
    expect(f.insert).toHaveBeenCalledTimes(1)
    expect(f.remove).not.toHaveBeenCalled()
    if (mode !== 'partial') expect(journal).toHaveBeenCalledOnce()
  }
})
it('refuses stale original content, external pages, cancellation, unsupported API, and wrong replacement bytes before writing', async () => {
  for (const mode of ['content', 'order', 'cancel', 'unsupported', 'replacement', 'guard']) {
    const f = await setup()
    if (mode === 'content') f.packages.old = f.replacement
    if (mode === 'order') f.setIds(['before', 'old', 'foreign', 'after'])
    if (mode === 'unsupported') f.support.mockReturnValue(false)
    await expect(
      f.adapter.stage(
        f.record,
        mode === 'replacement' ? f.original : f.replacement,
        vi.fn(),
        () => {
          if (mode === 'guard') throw new Error('stale')
        },
        mode === 'cancel' ? AbortSignal.abort() : undefined,
      ),
    ).rejects.toThrow()
    expect(f.insert).not.toHaveBeenCalled()
    expect(f.remove).not.toHaveBeenCalled()
  }
})
it('refuses deleting modified or moved staged pages and pending records', async () => {
  const f = await setup()
  await expect(f.adapter.discard(f.record, vi.fn())).rejects.toThrow()
  await f.adapter.stage(
    f.record,
    f.replacement,
    async (id) => {
      f.record = { ...f.record, state: 'discard_pending', newSlideId: id }
    },
    vi.fn(),
  )
  f.packages.new = f.original
  expect((await f.adapter.inspect(f.record)).status).toBe('conflict')
  await expect(f.adapter.discard(f.record, vi.fn())).rejects.toThrow()
  f.packages.new = f.replacement
  f.setIds(['before', 'new', 'old', 'after'])
  await expect(f.adapter.discard(f.record, vi.fn())).rejects.toThrow()
  expect(f.remove).not.toHaveBeenCalled()
})

it('stops on cancellation after a host write without deleting or issuing another write', async () => {
  const f = await setup(),
    controller = new AbortController(),
    journal = vi.fn()
  f.insert.mockImplementation((base64) => {
    f.setIds(['before', 'old', 'new', 'after'])
    f.packages.new = base64
    controller.abort()
  })
  await expect(
    f.adapter.stage(f.record, f.replacement, journal, vi.fn(), controller.signal),
  ).rejects.toThrow('cancelled')
  expect(f.insert).toHaveBeenCalledOnce()
  expect(journal).not.toHaveBeenCalled()
  expect(f.remove).not.toHaveBeenCalled()
})
it('retains exact inserted identity even when cancellation follows its receipt', async () => {
  const f = await setup(),
    controller = new AbortController(),
    journal = vi.fn(async (id: string) => {
      expect(id).toBe('new')
      controller.abort()
    })
  await expect(
    f.adapter.stage(f.record, f.replacement, journal, vi.fn(), controller.signal),
  ).rejects.toThrow('cancelled')
  expect(journal).toHaveBeenCalledOnce()
  expect(f.remove).not.toHaveBeenCalled()
})
it('does not guess an inserted ID with an external page or a reordered original', async () => {
  for (const ids of [
    ['before', 'old', 'new', 'external', 'after'],
    ['old', 'new', 'before', 'after'],
  ]) {
    const f = await setup(),
      journal = vi.fn()
    f.insert.mockImplementation((base64) => {
      f.setIds(ids)
      f.packages.new = base64
    })
    await expect(f.adapter.stage(f.record, f.replacement, journal, vi.fn())).rejects.toThrow(
      'office_concurrent_change',
    )
    expect(journal).not.toHaveBeenCalled()
    expect(f.remove).not.toHaveBeenCalled()
  }
})
it('recovers an acknowledged ID after deletion applied but sync failed without deleting again', async () => {
  const f = await setup()
  await f.adapter.stage(
    f.record,
    f.replacement,
    async (id) => {
      f.record = { ...f.record, state: 'discard_pending', newSlideId: id }
    },
    vi.fn(),
  )
  f.context.sync.mockImplementation(async () => {
    if (f.remove.mock.calls.length) throw new Error('connection_lost')
  })
  await expect(f.adapter.discard(f.record, vi.fn())).rejects.toThrow('connection_lost')
  expect(f.remove).toHaveBeenCalledExactlyOnceWith('new')
  f.context.sync.mockResolvedValue()
  await f.adapter.discard(f.record, vi.fn())
  expect(f.remove).toHaveBeenCalledTimes(1)
})
it('does not delete when a final binding check fails', async () => {
  const f = await setup()
  await f.adapter.stage(
    f.record,
    f.replacement,
    async (id) => {
      f.record = { ...f.record, state: 'discard_pending', newSlideId: id }
    },
    vi.fn(),
  )
  let calls = 0
  await expect(
    f.adapter.discard(f.record, () => {
      if (++calls === 2) throw new Error('document_changed')
    }),
  ).rejects.toThrow('document_changed')
  expect(f.remove).not.toHaveBeenCalled()
})

it('fingerprints every entry and path in sorted order, but ignores directories and ZIP dates', async () => {
  const first = new JSZip().file('b.xml', 'b', { date: new Date('2020-01-01') }).file('a.xml', 'a')
  const second = new JSZip().file('a.xml', 'a').file('b.xml', 'b', { date: new Date('2026-01-01') })
  second.folder('empty')
  const encode = (zip: JSZip) => zip.generateAsync({ type: 'base64' })
  expect(await presentationPackageDigest(await encode(first))).toBe(
    await presentationPackageDigest(await encode(second)),
  )
  second.remove('b.xml').file('c.xml', 'b')
  expect(await presentationPackageDigest(await encode(first))).not.toBe(
    await presentationPackageDigest(await encode(second)),
  )
  await expect(
    presentationPackageDigest(await encode(new JSZip().file('../escape.xml', 'bad'))),
  ).rejects.toThrow('invalid_tool_input')
})
