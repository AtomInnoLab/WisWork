import { afterEach, expect, it, vi } from 'vitest'
import {
  BrowserPowerPointAdapter,
  type PowerPointMasterOperation,
} from '../src/skills/powerpoint/browser-powerpoint-adapter'
afterEach(() => vi.unstubAllGlobals())
async function fixture() {
  const colors: Record<string, string> = Object.fromEntries(
    [
      'Accent1',
      'Accent2',
      'Accent3',
      'Accent4',
      'Accent5',
      'Accent6',
      'Dark1',
      'Dark2',
      'Light1',
      'Light2',
      'Hyperlink',
      'FollowedHyperlink',
    ].map((k) => [k, '#FFFFFF']),
  )
  const setThemeColor = vi.fn(),
    solid = { color: '#FFFFFF', transparency: 0, isNullObject: false, load: vi.fn() }
  const layout = {
    id: 'l1',
    name: 'Layout',
    background: {
      isMasterBackgroundFollowed: true,
      areBackgroundGraphicsHidden: false,
      load: vi.fn(),
      fill: { type: 'Solid', load: vi.fn() },
    },
  }
  const master = {
    id: 'm1',
    name: 'Main',
    layouts: { items: [layout], load: vi.fn(), getItem: vi.fn(() => layout) },
    background: { fill: { type: 'Solid', load: vi.fn(), getSolidFillOrNullObject: () => solid } },
    themeColorScheme: {
      getThemeColor: (slot: string) => ({
        get value() {
          return colors[slot]
        },
      }),
      setThemeColor,
    },
  }
  const masters = { items: [master], load: vi.fn(), getItem: vi.fn(() => master) }
  const slides = {
    items: ['s1', 's2'].map((id) => ({ id, slideMaster: { id: 'm1' }, layout: { id: 'l1' } })),
    load: vi.fn(),
    getCount: () => ({ value: slides.items.length }),
  }
  const sync = vi.fn(async () => {})
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', requirements: { isSetSupported: () => true } },
  })
  vi.stubGlobal('PowerPoint', {
    run: async (callback: (ctx: unknown) => unknown) =>
      callback({ presentation: { slides, slideMasters: masters }, sync }),
  })
  const adapter = new BrowserPowerPointAdapter()
  const before = await adapter.inspectSlideMasters()
  const operation: PowerPointMasterOperation = {
    op: 'set_master_theme_color',
    master_id: 'm1',
    theme_color: 'Accent1',
    color: '#112233',
  }
  const preimage = {
    before,
    operations: [operation],
    slideIds: ['s1', 's2'],
    dependencies: {
      slides: slides.items.map((s) => ({
        slideId: s.id,
        masterId: s.slideMaster.id,
        layoutId: s.layout.id,
      })),
    },
  }
  return {
    adapter,
    colors,
    setThemeColor,
    solid,
    slides,
    sync,
    operation,
    preimage,
    master,
    masters,
  }
}
it('guards complete native field/dependency/order preimage before setters', async () => {
  const f = await fixture(),
    guard = vi.fn(async () => {})
  await f.adapter.executeMasterOperations([f.operation], undefined, f.preimage, guard)
  expect(guard).toHaveBeenCalledOnce()
  expect(f.setThemeColor).toHaveBeenCalledWith('Accent1', '#112233')
})
it.each(['target', 'unaffected', 'order', 'dependency'])(
  'rejects %s drift at the actual SDK write boundary',
  async (mode) => {
    const f = await fixture()
    if (mode === 'target') f.colors.Accent1 = '#000000'
    if (mode === 'unaffected') f.solid.color = '#000000'
    if (mode === 'order') f.slides.items.reverse()
    if (mode === 'dependency') f.slides.items[1].layout.id = 'external-layout'
    await expect(
      f.adapter.executeMasterOperations([f.operation], undefined, f.preimage),
    ).rejects.toThrow('proposal_stale')
    expect(f.setThemeColor).not.toHaveBeenCalled()
  },
)
it('runs dynamic document/capability guard inside the native context immediately before a setter', async () => {
  const f = await fixture(),
    guard = vi.fn(async () => {
      throw new Error('presentation_document_changed')
    })
  await expect(
    f.adapter.executeMasterOperations([f.operation], undefined, f.preimage, guard),
  ).rejects.toThrow('presentation_document_changed')
  expect(f.setThemeColor).not.toHaveBeenCalled()
})
it('owns native operations and preimage across asynchronous SDK reads', async () => {
  const f = await fixture()
  f.sync.mockImplementationOnce(async () => {
    f.operation.color = '#000000'
    f.preimage.slideIds.reverse()
    f.preimage.before.masters[0].themeColors.Accent1 = '#000000'
  })
  await f.adapter.executeMasterOperations([f.operation], undefined, f.preimage)
  expect(f.setThemeColor).toHaveBeenCalledWith('Accent1', '#112233')
})

it.each(['target', 'unaffected', 'order', 'dependency'])(
  'rejects %s drift during awaited business guard before any setter',
  async (mode) => {
    const f = await fixture()
    const guard = async () => {
      await Promise.resolve()
      if (mode === 'target') f.colors.Accent1 = '#000000'
      if (mode === 'unaffected') f.solid.color = '#000000'
      if (mode === 'order') f.slides.items.reverse()
      if (mode === 'dependency') f.slides.items[1].layout.id = 'foreign-layout'
    }
    await expect(
      f.adapter.executeMasterOperations([f.operation], undefined, f.preimage, guard),
    ).rejects.toThrow('proposal_stale')
    expect(f.setThemeColor).not.toHaveBeenCalled()
  },
)

it.each(['target', 'order', 'dependency'])(
  'refreshes %s state with native fields in the final SDK read batch',
  async (mode) => {
    const f = await fixture()
    let reads = 0
    f.sync.mockImplementation(async () => {
      if (++reads !== 2) return
      if (mode === 'target') f.colors.Accent1 = '#000000'
      if (mode === 'order') f.slides.items.reverse()
      if (mode === 'dependency') f.slides.items[1].layout.id = 'foreign-layout'
    })
    await expect(
      f.adapter.executeMasterOperations([f.operation], undefined, f.preimage),
    ).rejects.toThrow('proposal_stale')
    expect(f.setThemeColor).not.toHaveBeenCalled()
  },
)
it('checks synchronous document capability and CAS guard after the last asynchronous SDK read', async () => {
  const f = await fixture()
  let valid = true
  f.sync.mockImplementation(async () => {
    valid = false
  })
  const finalGuard = vi.fn(() => {
    if (!valid) throw Error('presentation_document_changed')
  })
  await expect(
    f.adapter.executeMasterOperations(
      [f.operation],
      undefined,
      f.preimage,
      async () => {},
      finalGuard,
    ),
  ).rejects.toThrow('presentation_document_changed')
  expect(finalGuard).toHaveBeenCalledOnce()
  expect(f.setThemeColor).not.toHaveBeenCalled()
})

it('reloads master names in the final read batch to detect a concurrent rename', async () => {
  const f = await fixture()
  let properties = '',
    hostName = 'Main',
    reads = 0
  f.masters.load.mockImplementation((value: string) => {
    properties = value
  })
  f.sync.mockImplementation(async () => {
    if (++reads === 2) hostName = 'Renamed by user'
    if (properties.includes('items/name')) f.master.name = hostName
  })
  await expect(
    f.adapter.executeMasterOperations([f.operation], undefined, f.preimage),
  ).rejects.toThrow('proposal_stale')
  expect(f.setThemeColor).not.toHaveBeenCalled()
})
