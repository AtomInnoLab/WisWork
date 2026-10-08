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
it('captures and rechecks unchanged generated pages during every replacement phase', async () => {
  const f = await setup()
  f.record.version = 2
  f.record.untouchedSlideDigests = await f.adapter.captureUnchangedPageDigests(
    ['before', 'after'],
    f.record.beforeSlideIds,
  )
  expect(f.record.untouchedSlideDigests).toEqual([
    { slideId: 'before', digest: await presentationPackageDigest(f.original) },
    { slideId: 'after', digest: await presentationPackageDigest(f.original) },
  ])
  expect((await f.adapter.inspect(f.record)).status).toBe('baseline')
  f.packages.after = f.replacement
  expect((await f.adapter.inspect(f.record)).status).toBe('conflict')
  await expect(f.adapter.stage(f.record, f.replacement, vi.fn(), vi.fn())).rejects.toThrow(
    'office_concurrent_change',
  )
  expect(f.insert).not.toHaveBeenCalled()
  f.packages.after = f.original
  await f.adapter.stage(
    f.record,
    f.replacement,
    async (newSlideId) => {
      f.record = { ...f.record, state: 'inserted', newSlideId }
    },
    vi.fn(),
  )
  expect((await f.adapter.inspect(f.record)).status).toBe('staged')
  f.packages.before = f.replacement
  expect((await f.adapter.inspect(f.record)).status).toBe('conflict')
  await expect(f.adapter.commit({ ...f.record, state: 'commit_pending' }, vi.fn())).rejects.toThrow(
    'office_concurrent_change',
  )
  expect(f.remove).not.toHaveBeenCalled()
})
it('reconciles an exact pending insertion without issuing another host write', async () => {
  const f = await setup()
  expect(await f.adapter.reconcilePending(f.record)).toEqual({ status: 'baseline' })
  f.setIds(['before', 'old', 'new', 'after'])
  f.packages.new = f.replacement
  expect(await f.adapter.reconcilePending(f.record)).toEqual({
    status: 'inserted',
    newSlideId: 'new',
  })
  expect(f.insert).not.toHaveBeenCalled()
  expect(f.remove).not.toHaveBeenCalled()
  f.packages.new = f.original
  expect(await f.adapter.reconcilePending(f.record)).toEqual({ status: 'conflict' })
  f.packages.new = f.replacement
  f.setIds(['before', 'new', 'old', 'after'])
  expect(await f.adapter.reconcilePending(f.record)).toEqual({ status: 'conflict' })
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

it('checks ZIP limits before inflation and bounds real inflated bytes even with forged sizes', async () => {
  const oversized = new JSZip().file('ppt/slides/slide1.xml', 'x'.repeat(2 * 1024 * 1024 + 1))
  const bytes = await oversized.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
  const load = vi.spyOn(JSZip, 'loadAsync')
  try {
    await expect(presentationPackageDigest(Buffer.from(bytes).toString('base64'))).rejects.toThrow(
      'invalid_tool_input',
    )
    expect(load).toHaveBeenLastCalledWith(
      expect.any(String),
      expect.objectContaining({ checkCRC32: false }),
    )
    const forged = new Uint8Array(bytes),
      view = new DataView(forged.buffer)
    for (let i = 0; i + 30 <= forged.length; i++) {
      const signature = view.getUint32(i, true)
      if (signature === 0x04034b50) view.setUint32(i + 22, 1, true)
      if (signature === 0x02014b50) view.setUint32(i + 24, 1, true)
    }
    await expect(presentationPackageDigest(Buffer.from(forged).toString('base64'))).rejects.toThrow(
      'invalid_tool_input',
    )
  } finally {
    load.mockRestore()
  }
})

it('rechecks staged content after obtaining the exact deletion proxy', async () => {
  const f = await setup()
  await f.adapter.stage(
    f.record,
    f.replacement,
    async (id) => {
      f.record = { ...f.record, state: 'discard_pending', newSlideId: id }
    },
    vi.fn(),
  )
  f.context.presentation.slides.getItem.mockClear()
  f.context.sync.mockImplementation(async () => {
    if (
      f.context.presentation.slides.getItem.mock.calls.filter(([id]) => id === 'new').length === 2
    )
      f.packages.new = f.original
  })
  await expect(f.adapter.discard(f.record, vi.fn())).rejects.toThrow('office_concurrent_change')
  expect(f.remove).not.toHaveBeenCalled()
})

async function staged() {
  const f = await setup()
  await f.adapter.stage(
    f.record,
    f.replacement,
    async (id) => {
      f.record = { ...f.record, state: 'commit_pending', newSlideId: id }
    },
    vi.fn(),
  )
  return f
}
async function applied() {
  const f = await staged()
  await f.adapter.commit(f.record, vi.fn())
  f.record = { ...f.record, state: 'undo_pending' }
  f.insert.mockImplementation((base64) => {
    f.setIds(['before', 'new', 'restored', 'after'])
    f.packages.restored = base64
  })
  return f
}
it('commits exact staged pages, restores originals before deletion, and only reads already completed writes', async () => {
  const f = await staged()
  await f.adapter.commit(f.record, vi.fn())
  await f.adapter.commit(f.record, vi.fn())
  expect(f.remove).toHaveBeenCalledExactlyOnceWith('old')
  expect((await f.adapter.inspect(f.record)).status).toBe('applied')
  const g = await applied()
  const journal = vi.fn(async (id: string) => {
    expect(g.remove).toHaveBeenCalledTimes(1)
    g.record = { ...g.record, state: 'restore_inserted', restoredSlideId: id }
  })
  await g.adapter.undo(g.record, g.original, journal, vi.fn())
  expect(journal).toHaveBeenCalledExactlyOnceWith('restored')
  expect(g.insert).toHaveBeenLastCalledWith(g.original, {
    targetSlideId: 'new',
    formatting: 'KeepSourceFormatting',
  })
  expect((await g.adapter.inspect(g.record)).status).toBe('undone')
  await g.adapter.undo(g.record, g.original, journal, vi.fn())
  expect(g.remove.mock.calls).toEqual([['old'], ['new']])
  expect(g.insert).toHaveBeenCalledTimes(2)
})
it('recovers commit after deletion applied but sync failed without deleting twice', async () => {
  const f = await staged()
  f.context.sync.mockImplementation(async () => {
    if (f.remove.mock.calls.length) throw new Error('connection_lost')
  })
  await expect(f.adapter.commit(f.record, vi.fn())).rejects.toThrow('connection_lost')
  f.context.sync.mockResolvedValue()
  await f.adapter.commit(f.record, vi.fn())
  expect(f.remove).toHaveBeenCalledExactlyOnceWith('old')
})
it('retains replacement after failed restore receipt and never guesses unknown inserted identity', async () => {
  const f = await applied()
  await expect(
    f.adapter.undo(
      f.record,
      f.original,
      async () => {
        throw new Error('receipt_failed')
      },
      vi.fn(),
    ),
  ).rejects.toThrow('receipt_failed')
  await expect(f.adapter.undo(f.record, f.original, vi.fn(), vi.fn())).rejects.toThrow(
    'office_concurrent_change',
  )
  expect(f.insert).toHaveBeenCalledTimes(2)
  expect(f.remove).toHaveBeenCalledExactlyOnceWith('old')
})
it('resumes journalled restoration and partially applied deletion without reinserting', async () => {
  const f = await applied()
  await expect(
    f.adapter.undo(
      f.record,
      f.original,
      async (id) => {
        f.record = { ...f.record, state: 'restore_inserted', restoredSlideId: id }
        throw new Error('lost_response')
      },
      vi.fn(),
    ),
  ).rejects.toThrow('lost_response')
  expect((await f.adapter.inspect(f.record)).status).toBe('restore_staged')
  f.context.sync.mockImplementation(async () => {
    if (f.remove.mock.calls.length === 2) throw new Error('connection_lost')
  })
  await expect(f.adapter.undo(f.record, f.original, vi.fn(), vi.fn())).rejects.toThrow(
    'connection_lost',
  )
  f.context.sync.mockResolvedValue()
  await f.adapter.undo(f.record, f.original, vi.fn(), vi.fn())
  expect(f.insert).toHaveBeenCalledTimes(2)
  expect(f.remove).toHaveBeenCalledTimes(2)
})
it('blocks commit on content, order, support, cancellation, state and final binding changes', async () => {
  for (const mode of ['original', 'replacement', 'order', 'support', 'cancel', 'state', 'guard']) {
    const f = await staged()
    if (mode === 'original') f.packages.old = f.replacement
    if (mode === 'replacement') f.packages.new = f.original
    if (mode === 'order') f.setIds(['old', 'new', 'before', 'after'])
    if (mode === 'support') f.support.mockReturnValue(false)
    if (mode === 'state') f.record.state = 'staged'
    let calls = 0
    await expect(
      f.adapter.commit(
        f.record,
        () => {
          if (mode === 'guard' && ++calls === 2) throw new Error('document_changed')
        },
        mode === 'cancel' ? AbortSignal.abort() : undefined,
      ),
    ).rejects.toThrow()
    expect(f.remove).not.toHaveBeenCalled()
  }
})
it('retains replacement when restored content, replacement content, position or receipt guards differ', async () => {
  for (const mode of ['original', 'replacement', 'order', 'guard', 'cancel']) {
    const f = await applied(),
      controller = new AbortController()
    await expect(
      f.adapter.undo(
        f.record,
        f.original,
        async (id) => {
          f.record = { ...f.record, state: 'restore_inserted', restoredSlideId: id }
          if (mode === 'original') f.packages.restored = f.replacement
          if (mode === 'replacement') f.packages.new = f.original
          if (mode === 'order') f.setIds(['before', 'restored', 'new', 'after'])
          if (mode === 'guard') throw new Error('document_changed')
          if (mode === 'cancel') controller.abort()
        },
        vi.fn(),
        controller.signal,
      ),
    ).rejects.toThrow()
    expect(f.remove).toHaveBeenCalledExactlyOnceWith('old')
  }
})
it('requires matching backup bytes and explicit pending states before restoration', async () => {
  const f = await staged()
  await f.adapter.commit(f.record, vi.fn())
  await expect(f.adapter.undo(f.record, f.original, vi.fn(), vi.fn())).rejects.toThrow()
  f.record = { ...f.record, state: 'undo_pending' }
  await expect(f.adapter.undo(f.record, f.replacement, vi.fn(), vi.fn())).rejects.toThrow(
    'office_concurrent_change',
  )
  expect(f.insert).toHaveBeenCalledTimes(1)
})
it('rechecks original content after loading the commit deletion proxy', async () => {
  const f = await staged()
  f.context.presentation.slides.getItem.mockClear()
  f.context.sync.mockImplementation(async () => {
    if (
      f.context.presentation.slides.getItem.mock.calls.filter(([id]) => id === 'old').length === 2
    )
      f.packages.old = f.replacement
  })
  await expect(f.adapter.commit(f.record, vi.fn())).rejects.toThrow('office_concurrent_change')
  expect(f.remove).not.toHaveBeenCalled()
})
it('checks both pages again after obtaining the undo deletion proxy', async () => {
  const f = await applied()
  await expect(
    f.adapter.undo(
      f.record,
      f.original,
      async (id) => {
        f.record = { ...f.record, state: 'restore_inserted', restoredSlideId: id }
        throw new Error('lost_response')
      },
      vi.fn(),
    ),
  ).rejects.toThrow('lost_response')
  f.context.presentation.slides.getItem.mockClear()
  f.context.sync.mockImplementation(async () => {
    if (
      f.context.presentation.slides.getItem.mock.calls.filter(([id]) => id === 'new').length === 2
    )
      f.packages.restored = f.replacement
  })
  await expect(f.adapter.undo(f.record, f.original, vi.fn(), vi.fn())).rejects.toThrow(
    'office_concurrent_change',
  )
  expect(f.remove).toHaveBeenCalledExactlyOnceWith('old')
})
it('stops after cancelled or uncertain restoration writes without inferring ownership or deleting', async () => {
  for (const mode of ['cancel', 'partial', 'foreign']) {
    const f = await applied(),
      controller = new AbortController(),
      journal = vi.fn()
    f.insert.mockImplementation((base64) => {
      f.setIds(['before', 'new', 'restored', ...(mode === 'foreign' ? ['external'] : []), 'after'])
      f.packages.restored = base64
      if (mode === 'cancel') controller.abort()
      if (mode === 'partial') throw new Error('connection_lost')
    })
    await expect(
      f.adapter.undo(f.record, f.original, journal, vi.fn(), controller.signal),
    ).rejects.toThrow()
    expect(journal).not.toHaveBeenCalled()
    expect(f.remove).toHaveBeenCalledExactlyOnceWith('old')
    await expect(f.adapter.undo(f.record, f.original, journal, vi.fn())).rejects.toThrow(
      'office_concurrent_change',
    )
    expect(f.insert).toHaveBeenCalledTimes(2)
  }
})
