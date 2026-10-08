import { afterEach, expect, it } from 'vitest'
import { masterXmlFixture, cleanupMasterXmlFixtures } from './helpers/master-xml-fixture.js'
import { assertMasterXmlPagePreserved } from '../src/skills/powerpoint/presentation-master-xml-package.js'
afterEach(cleanupMasterXmlFixtures)
const edges = [
  ['forward_layout', 'before'],
  ['forward_layout', 'after'],
  ['forward_source_delete', 'before'],
  ['forward_source_delete', 'after'],
  ['undo_layout', 'before'],
  ['undo_layout', 'after'],
  ['undo_carrier_delete', 'before'],
  ['undo_carrier_delete', 'after'],
] as const
it.each(edges)(
  'requires explicit durable reconciliation for %s %s ACK uncertainty',
  async (edge, when) => {
    const f = await masterXmlFixture(3),
      p = await f.propose(),
      id = String(p.preview.changeId)
    const undo = edge.startsWith('undo_')
    if (undo) await f.confirm()
    const carrier = f.data.get(id)?.stagedSource?.slideId
    if (undo) expect(carrier).toBeDefined()
    let injected = false
    const apply = f.adapter.applyLayout.getMockImplementation()!,
      remove = f.adapter.remove.getMockImplementation()!
    const selected = (input: any) =>
      edge.endsWith('layout') ? input.slideId === 's1' : input.slideId === (undo ? carrier : 's0')
    const run = async (original: () => Promise<void>, input: any) => {
      if (selected(input) && !injected) {
        injected = true
        if (when === 'before') throw Error('synthetic_lost_ack')
        await original()
        throw Error('synthetic_lost_ack')
      }
      await original()
    }
    if (edge.endsWith('layout'))
      f.adapter.applyLayout.mockImplementation(async (...args) =>
        run(() => apply(...args), args[0]),
      )
    else f.adapter.remove.mockImplementation(async (...args) => run(() => remove(...args), args[0]))
    if (undo) expect((await f.tool('undo', { change_id: id })).isError).not.toBe(true)
    await expect(f.confirm()).rejects.toThrow('synthetic_lost_ack')
    expect(injected).toBe(true)
    const uncertain = f.data.get(id)!
    expect(uncertain.pending).toBeDefined()
    const nativeCounts = () => [
      f.adapter.stage.mock.calls.length,
      f.adapter.applyLayout.mock.calls.length,
      f.adapter.remove.mock.calls.length,
    ]
    const counts = nativeCounts(),
      observedOrder = [...f.order],
      observedPackages = [...f.packages]
    // No inverse follows an unknown native receipt, and reload cannot replay it.
    const relevant = edge.endsWith('layout')
      ? f.adapter.applyLayout.mock.calls
      : f.adapter.remove.mock.calls
    expect(relevant.filter(([input]) => selected(input))).toHaveLength(
      1 + (edge === 'undo_layout' ? 1 : 0),
    )
    f.reopen()
    const inspect = await f.tool('inspect', { change_id: id })
    expect(inspect.isError).not.toBe(true)
    expect(nativeCounts()).toEqual(counts)
    expect(f.order).toEqual(observedOrder)
    expect([...f.packages]).toEqual(observedPackages)
    expect(f.data.get(id)!.pending).toEqual(uncertain.pending)
    expect((await f.tool('resume', { change_id: id })).isError).toBe(true)
    expect(nativeCounts()).toEqual(counts)
    expect((await f.tool('reconcile', { change_id: id })).isError).not.toBe(true)
    expect(f.proposals.pending()!.impact.host).toBe('local_checkpoint')
    await f.confirm()
    expect(nativeCounts()).toEqual(counts)
    expect(f.order).toEqual(observedOrder)
    expect([...f.packages]).toEqual(observedPackages)
    expect(f.data.get(id)!.pending).toBeUndefined()
    // A proven before-write baseline never becomes an applied forward transaction.
    if (!undo && when === 'before') expect(f.data.get(id)!.state).not.toBe('applied')
    if (f.data.get(id)!.state !== 'undone') {
      expect((await f.tool('undo', { change_id: id })).isError).not.toBe(true)
      await f.confirm()
    }
    expect(f.data.get(id)!.state).toBe('undone')
    expect(f.order).toHaveLength(3)
    expect(f.order.slice(1)).toEqual(['s1', 's2'])
    expect(new Set(f.order).size).toBe(3)
    for (const page of f.order) {
      const dep = f.deps.get(page)!,
        master = f.masters.find((value) => value.masterId === dep.masterId)
      expect(master?.layouts.some((layout) => layout.layoutId === dep.layoutId)).toBe(true)
      await assertMasterXmlPagePreserved(f.original, f.packages.get(page)!, {
        expectedMasterBase64: f.original,
        targetMasterPath: 'ppt/slideMasters/slideMaster1.xml',
        packageLayoutPath: 'ppt/slideLayouts/slideLayout1.xml',
      })
    }
  },
  120000,
)
