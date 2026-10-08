import { afterEach, expect, it } from 'vitest'
import JSZip from 'jszip'
import { masterXmlFixture, cleanupMasterXmlFixtures } from './helpers/master-xml-fixture.js'
import { assertMasterXmlPagePreserved } from '../src/skills/powerpoint/presentation-master-xml-package.js'

afterEach(cleanupMasterXmlFixtures)
const preservation = (base64: string) => ({
  expectedMasterBase64: base64,
  targetMasterPath: 'ppt/slideMasters/slideMaster1.xml',
  packageLayoutPath: 'ppt/slideLayouts/slideLayout1.xml',
})
async function proveRestored(f: Awaited<ReturnType<typeof masterXmlFixture>>, original: string) {
  expect(f.order).toHaveLength(3)
  for (const id of f.order) {
    const dependency = f.deps.get(id)!
    const master = f.masters.find((m) => m.masterId === dependency.masterId)!
    expect(master).toBeDefined()
    expect(master.layouts.some((l) => l.layoutId === dependency.layoutId)).toBe(true)
    await assertMasterXmlPagePreserved(original, f.packages.get(id)!, preservation(original))
  }
  expect(new Set(f.order).size).toBe(3)
}

it('rebuilds disappeared original native master and actual layout associations after known source removal', async () => {
  const f = await masterXmlFixture(3)
  const remove = f.adapter.remove.getMockImplementation()!
  f.adapter.remove.mockImplementation(async (...args) => {
    await remove(...args)
    if (args[0].slideId === 's0') {
      const index = f.masters.findIndex((m) => m.masterId === 'm1')
      expect(index).toBeGreaterThanOrEqual(0)
      f.masters.splice(index, 1)
    }
  })
  const p = await f.propose(),
    id = String(p.preview.changeId)
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('applied')
  expect(f.masters.some((m) => m.masterId === 'm1')).toBe(false)
  f.reopen()
  expect((await f.tool('undo', { change_id: id })).isError).not.toBe(true)
  await f.confirm()
  expect(f.data.get(id)!.state).toBe('undone')
  expect(f.order.slice(1)).toEqual(['s1', 's2'])
  const restoredMaster = f.deps.get(f.order[0]!)!.masterId
  expect(restoredMaster).not.toBe('m1')
  expect(f.order.every((page) => f.deps.get(page)!.masterId === restoredMaster)).toBe(true)
  await proveRestored(f, f.original)
}, 120000)

it.each(['payload', 'notes'] as const)(
  'retains %s loss as uncertain and explicitly restores original full packages',
  async (kind) => {
    const f = await masterXmlFixture(3),
      originalOrder = [...f.order]
    const source = await JSZip.loadAsync(f.original, { base64: true })
    source.file(
      'ppt/notesSlides/notesSlide1.xml',
      '<p:notes xmlns:p="urn:p"><p:cSld>protected original notes</p:cSld></p:notes>',
    )
    const original = await source.generateAsync({ type: 'base64' })
    for (const page of f.order) f.packages.set(page, original)
    f.masters[0]!.base64 = original
    const apply = f.adapter.applyLayout.getMockImplementation()!
    let injected = false
    f.adapter.applyLayout.mockImplementation(async (...args) => {
      await apply(...args)
      if (args[0].slideId === 's1' && !injected) {
        injected = true
        const actual = await JSZip.loadAsync(f.packages.get('s1')!, { base64: true })
        if (kind === 'payload')
          actual.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"><p:cSld/></p:sld>')
        else actual.remove('ppt/notesSlides/notesSlide1.xml')
        f.packages.set('s1', await actual.generateAsync({ type: 'base64' }))
      }
    })
    const p = await f.propose(),
      id = String(p.preview.changeId)
    await expect(f.confirm()).rejects.toThrow()
    expect(injected).toBe(true)
    expect(f.data.get(id)!.state).not.toBe('applied')
    expect(f.data.get(id)!.pending).toBeDefined()
    expect(f.data.get(id)!.reviews).toEqual([])
    f.reopen()
    expect((await f.tool('resume', { change_id: id })).isError).toBe(true)
    expect((await f.tool('reconcile', { change_id: id })).isError).not.toBe(true)
    await f.confirm()
    expect(f.data.get(id)!.state).toBe('recovery_required')
    expect((await f.tool('undo', { change_id: id })).isError).not.toBe(true)
    await f.confirm()
    expect(f.data.get(id)!.state).toBe('undone')
    const fallbackIndex = f.adapter.stage.mock.calls.findIndex(
      ([input]) => input.sourceSlideId === 's1',
    )
    expect(fallbackIndex).toBeGreaterThanOrEqual(0)
    const restoredTarget = await f.adapter.stage.mock.results[fallbackIndex]!.value
    expect(restoredTarget.slideId).not.toBe('s1')
    expect(f.order).toEqual([originalOrder[0], restoredTarget.slideId, originalOrder[2]])
    await proveRestored(f, original)
    for (const page of f.order) {
      const zip = await JSZip.loadAsync(f.packages.get(page)!, { base64: true })
      expect(await zip.file('ppt/notesSlides/notesSlide1.xml')!.async('string')).toContain(
        'protected original notes',
      )
      expect(await zip.file('ppt/slides/slide1.xml')!.async('string')).toContain('original')
    }
  },
  120000,
)

it('uses monotonically fresh immutable PC review keys after repeated reviews, undo and recapture', async () => {
  const f = await masterXmlFixture(3),
    p = await f.propose(),
    id = String(p.preview.changeId)
  await f.confirm()
  const keys: string[] = []
  const review = async (notes: string) => {
    const page = f.order[0]!,
      capture = await f.tool('capture', { change_id: id, slide_id: page })
    expect(capture.isError).not.toBe(true)
    const body = JSON.parse(capture.output)
    expect(
      (
        await f.tool('review', {
          change_id: id,
          slide_id: page,
          screenshot_digest: body.screenshotDigest,
          status: 'pass',
          notes,
        })
      ).isError,
    ).not.toBe(true)
    keys.push(f.data.get(id)!.reviews[0]!.reviewRef.key)
  }
  for (let i = 0; i < 4; i++) await review(`before undo ${i}`)
  expect(f.data.get(id)!.reviewSequence).toBe(4)
  f.reopen()
  expect((await f.tool('undo', { change_id: id })).isError).not.toBe(true)
  await f.confirm()
  expect(f.data.get(id)!.reviewSequence).toBe(4)
  f.reopen()
  await review('fresh restored capture')
  expect(f.data.get(id)!.reviewSequence).toBe(5)
  expect(new Set(keys).size).toBe(5)
  const finished = f.request.mock.calls
    .map(([body]) => body as Record<string, unknown>)
    .filter((body) => body.operation === 'package_backup_finish')
  for (const key of keys) expect(finished.filter((body) => body.key === key)).toHaveLength(1)
}, 120000)
