import { afterEach, describe, expect, it, vi } from 'vitest'
import { createBrowserPresentationImportAdapter } from '../src/skills/powerpoint/presentation-import.js'

function host(initial: string[] = ['original-1', 'original-2']) {
  let ids = [...initial]
  let queued = false
  let afterWrite: (() => void) | undefined
  let pageCount = 2
  const insert = vi.fn((_base64?: string, options?: { sourceSlideIds?: string[] }) => {
    pageCount = options?.sourceSlideIds?.length ?? 2
    queued = true
  })
  const slides = { items: [] as Array<{ id: string }>, load: vi.fn() }
  const sync = vi.fn(async () => {
    if (queued) {
      queued = false
      ids.push(...['generated-1', 'generated-2'].slice(0, pageCount))
      afterWrite?.()
    }
    slides.items = ids.map((id) => ({ id }))
  })
  const context = { presentation: { slides, insertSlidesFromBase64: insert }, sync }
  vi.stubGlobal('Office', {
    context: {
      host: 'PowerPoint',
      requirements: {
        isSetSupported: (name: string, version: string) =>
          name === 'PowerPointApi' && version === '1.2',
      },
    },
  })
  vi.stubGlobal('PowerPoint', {
    run: async (action: (value: typeof context) => Promise<unknown>) => action(context),
  })
  return {
    insert,
    sync,
    context,
    ids: () => ids,
    replace: (value: string[]) => {
      ids = value
    },
    afterWrite: (action: () => void) => {
      afterWrite = action
    },
  }
}

afterEach(() => vi.unstubAllGlobals())

describe('PowerPoint generated-deck import', () => {
  it('exports one existing page without inserting or changing slide order', async () => {
    const runtime = host()
    const exported = { value: 'UEs=' }
    const slide = { exportAsBase64: vi.fn(() => exported) }
    Object.assign(runtime.context.presentation.slides, { getItem: vi.fn(() => slide) })
    const adapter = createBrowserPresentationImportAdapter()
    expect(adapter.supportsPageExport?.()).toBe(false)
    await expect(adapter.exportPage?.('original-1')).rejects.toThrow('office_api_unsupported')
    vi.stubGlobal('Office', {
      context: {
        host: 'PowerPoint',
        requirements: { isSetSupported: () => true },
      },
    })
    expect(adapter.supportsPageExport?.()).toBe(true)
    expect(await adapter.exportPage?.('original-1')).toBe('UEs=')
    expect(slide.exportAsBase64).toHaveBeenCalledOnce()
    expect(runtime.ids()).toEqual(['original-1', 'original-2'])
    expect(runtime.insert).not.toHaveBeenCalled()
  })
  it.each([{ original: [] }, { original: ['original-1', 'original-2'] }])(
    'appends to $original, preserving original order and verifying the receipt',
    async ({ original }) => {
      const runtime = host(original)
      const adapter = createBrowserPresentationImportAdapter()
      expect(adapter.available()).toBe(true)
      const before = await adapter.snapshot()
      const receipt = await adapter.insert('UEs=', 2, before)
      expect(runtime.insert).toHaveBeenCalledWith('UEs=', {
        formatting: 'KeepSourceFormatting',
        ...(original.length ? { targetSlideId: original.at(-1) } : {}),
      })
      expect(receipt.slideIds).toEqual(['generated-1', 'generated-2'])
      expect(runtime.ids()).toEqual([...original, ...receipt.slideIds])
      expect(await adapter.verify(receipt, before)).toBe(true)
      runtime.replace([...runtime.ids()].reverse())
      expect(await adapter.verify(receipt, before)).toBe(false)
    },
  )

  it('accepts a bounded multi-megabyte generated package without regex stack overflow', async () => {
    const runtime = host()
    const adapter = createBrowserPresentationImportAdapter()
    const before = await adapter.snapshot()
    await expect(adapter.insert('A'.repeat(8 * 1024 * 1024), 2, before)).resolves.toEqual({
      slideIds: ['generated-1', 'generated-2'],
    })
    expect(runtime.insert).toHaveBeenCalledOnce()
  })

  it('blocks stale baselines and pre-write cancellation before inserting', async () => {
    const runtime = host()
    const adapter = createBrowserPresentationImportAdapter()
    const before = await adapter.snapshot()
    runtime.replace(['other-document'])
    await expect(adapter.insert('UEs=', 2, before)).rejects.toThrow('proposal_stale')
    const controller = new AbortController()
    controller.abort()
    await expect(adapter.insert('UEs=', 2, before, controller.signal)).rejects.toThrow('cancelled')
    expect(runtime.insert).not.toHaveBeenCalled()
  })

  it('reconciles and verifies a committed insertion despite cancellation during sync', async () => {
    const runtime = host()
    const adapter = createBrowserPresentationImportAdapter()
    const before = await adapter.snapshot()
    const controller = new AbortController()
    runtime.afterWrite(() => controller.abort())
    const receipt = await adapter.insert('UEs=', 2, before, controller.signal)
    expect(controller.signal.aborted).toBe(true)
    expect(await adapter.verify(receipt, before, controller.signal)).toBe(true)
    expect(runtime.insert).toHaveBeenCalledOnce()
  })

  it('reports uncertain state after a partially committed sync failure without deleting any page', async () => {
    const runtime = host()
    const adapter = createBrowserPresentationImportAdapter()
    const before = await adapter.snapshot()
    runtime.afterWrite(() => {
      runtime.replace([...before.slideIds, 'generated-1'])
      throw new Error('Office sync failed')
    })
    await expect(adapter.insert('UEs=', 2, before)).rejects.toThrow('office_state_uncertain')
    expect(runtime.ids()).toEqual([...before.slideIds, 'generated-1'])
    expect(runtime.insert).toHaveBeenCalledOnce()
    expect(runtime.context.presentation.slides.items.map((slide) => slide.id)).toEqual(
      runtime.ids(),
    )
  })

  it.each(['count', 'order'])(
    'reports uncertain state after a committed %s mismatch',
    async (change) => {
      const runtime = host()
      const adapter = createBrowserPresentationImportAdapter()
      const before = await adapter.snapshot()
      runtime.afterWrite(() =>
        runtime.replace(
          change === 'count' ? [...runtime.ids(), 'user-added'] : [...runtime.ids()].reverse(),
        ),
      )
      await expect(adapter.insert('UEs=', 2, before)).rejects.toThrow('office_state_uncertain')
      expect(runtime.insert).toHaveBeenCalledOnce()
    },
  )

  it('requires the import API and validates bounded base64 and expected page count before writing', async () => {
    const runtime = host()
    const adapter = createBrowserPresentationImportAdapter()
    const before = await adapter.snapshot()
    for (const [data, count] of [
      ['not base64', 2],
      ['UEs=', 0],
      ['UEs=', 1.5],
      ['A'.repeat(16 * 1024 * 1024 + 4), 2],
    ] as const) {
      await expect(adapter.insert(data, count, before)).rejects.toThrow('invalid_tool_input')
    }
    delete (runtime.context.presentation as { insertSlidesFromBase64?: unknown })
      .insertSlidesFromBase64
    await expect(adapter.insert('UEs=', 2, before)).rejects.toThrow('office_api_unsupported')
    expect(runtime.insert).not.toHaveBeenCalled()
    vi.stubGlobal('Office', undefined)
    expect(adapter.available()).toBe(false)
    await expect(adapter.snapshot()).rejects.toThrow('office_api_unsupported')
  })
})

it('inserts exactly one selected source page and preserves the structural baseline', async () => {
  const runtime = host(['256'])
  const adapter = createBrowserPresentationImportAdapter()
  const before = await adapter.snapshot()
  const receipt = await adapter.insertPage!('UEs=', '257#', before)
  expect(runtime.insert).toHaveBeenCalledWith('UEs=', {
    formatting: 'KeepSourceFormatting',
    targetSlideId: '256#',
    sourceSlideIds: ['257#'],
  })
  expect(receipt.slideIds).toEqual(['generated-1'])
  expect(await adapter.verify(receipt, before)).toBe(true)
})
it.each(['0#', '255#', '4294967296#', '256', 'x#', '0256#'])(
  'rejects an invalid source page selector %s before queuing Office writes',
  async (id) => {
    const runtime = host()
    const adapter = createBrowserPresentationImportAdapter()
    await expect(adapter.insertPage!('UEs=', id, await adapter.snapshot())).rejects.toThrow(
      'invalid_tool_input',
    )
    expect(runtime.insert).not.toHaveBeenCalled()
  },
)
