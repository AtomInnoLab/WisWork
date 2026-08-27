import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import {
  addElement,
  createBlankPptx,
  duplicateSlide,
  ensureElementCreationId,
  fingerprintSlide,
  groupElements,
  openPptx,
  copyElementData,
  moveSlide,
  normalizeCreationId,
  pasteElements,
  reorderElement,
  resolvePresentationTarget,
  savePptx,
  setCreationIdInElementXml,
  ungroupElement,
} from '../src/index'

const OFF_A = { x: 100, y: 200, cx: 300, cy: 400 }
const OFF_B = { x: 500, y: 600, cx: 700, cy: 800 }
const here = dirname(fileURLToPath(import.meta.url))
const fixture = () => readFileSync(join(here, 'fixtures', '01_standard_business.pptx'))
const makeFactory = () => {
  let value = 0
  return () => `{00000000-0000-4000-8000-${(++value).toString().padStart(12, '0')}}`
}

describe('durable presentation targets', () => {
  it('accepts Office GUID-shaped creationIds without imposing UUID version bits', () => {
    expect(normalizeCreationId('{822E6D20-D7BC-2841-A643-D49A6EF008A2}')).toBe(
      '{822E6D20-D7BC-2841-A643-D49A6EF008A2}',
    )
  })

  it('adds creationId to the existing cNvPr extension list without creating invalid siblings', () => {
    const xml =
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="x"><a:extLst><a:ext uri="other"/></a:extLst></p:cNvPr></p:nvSpPr></p:sp>'
    const patched = setCreationIdInElementXml(xml, '{00000000-0000-4000-8000-000000000001}')
    expect(patched.match(/<a:extLst>/g)).toHaveLength(1)
    expect(patched).toContain('<a:ext uri="other"/>')
    expect(patched).toContain('<a16:creationId')
  })

  it('preserves slide part identity and shape creationId across save/reopen and reorder', async () => {
    const factory = makeFactory()
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const first = addElement(slide, { kind: 'rect', offset: OFF_A }, { creationIdFactory: factory })
    const second = addElement(
      slide,
      { kind: 'ellipse', offset: OFF_B },
      { creationIdFactory: factory },
    )
    const durableSlideId = slide.durableId
    const durableElementId = first.creationId

    expect(durableSlideId).toBe('ppt/slides/slide1.xml')
    expect(durableElementId).toBe('{00000000-0000-4000-8000-000000000001}')
    expect(reorderElement(slide, first.id, 'front')).toBe(true)

    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.deck.slides[0]!.durableId).toBe(durableSlideId)
    expect(reopened.deck.slides[0]!.elements.map((element) => element.creationId)).toEqual([
      second.creationId,
      durableElementId,
    ])
  })

  it('mints distinct creationIds when duplicating a slide', async () => {
    const factory = makeFactory()
    const opened = await openPptx(await createBlankPptx())
    const source = opened.deck.slides[0]!
    const original = addElement(
      source,
      { kind: 'rect', offset: OFF_A },
      { creationIdFactory: factory },
    )
    const copy = duplicateSlide(opened, 0, { creationIdFactory: factory })

    expect(copy?.durableId).not.toBe(source.durableId)
    expect(copy?.elements.at(-1)?.creationId).toBeDefined()
    expect(copy?.elements.at(-1)?.creationId).not.toBe(original.creationId)
    expect(moveSlide(opened, 1, 0)).toBe(true)
    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.deck.slides[0]!.durableId).toBe(copy?.durableId)
    expect(reopened.deck.slides[1]!.durableId).toBe(source.durableId)
  })

  it('does not share a creationId when an element is copied', async () => {
    const factory = makeFactory()
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const source = addElement(
      slide,
      { kind: 'rect', offset: OFF_A },
      { creationIdFactory: factory },
    )
    const copied = copyElementData(opened, slide, source)
    const result = pasteElements(
      opened,
      0,
      [copied],
      { dx: 10, dy: 10 },
      { creationIdFactory: factory },
    )
    expect(result).not.toBeNull()
    expect(result!.slide.elements.at(-1)?.creationId).toBeDefined()
    expect(result!.slide.elements.at(-1)?.creationId).not.toBe(source.creationId)
  })

  it('preserves child creationIds through group and ungroup while minting the new group', async () => {
    const factory = makeFactory()
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const first = addElement(slide, { kind: 'rect', offset: OFF_A }, { creationIdFactory: factory })
    const second = addElement(
      slide,
      { kind: 'ellipse', offset: OFF_B },
      { creationIdFactory: factory },
    )
    const childIds = [first.creationId, second.creationId]
    const grouped = groupElements(opened, 0, [first.id, second.id], { creationIdFactory: factory })
    expect(grouped).not.toBeNull()
    const group = grouped!.slide.elements.find((element) => element.id === grouped!.groupId)!
    expect(group.creationId).toBeDefined()
    expect(group.type).toBe('group')
    if (group.type !== 'group') throw new Error('expected group')
    expect(group.children.map((child) => child.creationId)).toEqual(childIds)

    const ungrouped = ungroupElement(opened, 0, group.id)
    expect(ungrouped?.elements.map((element) => element.creationId)).toEqual(childIds)
  })

  it('does not add creationIds to an untouched legacy slide', async () => {
    const bytes = await createBlankPptx()
    const opened = await openPptx(bytes)
    const path = opened.deck.slides[0]!.path
    const before = opened.archive.readText(path)!
    expect(before).not.toContain('a16:creationId')

    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.archive.readText(path)).toBe(before)
    expect(reopened.archive.readText(path)).not.toContain('a16:creationId')
  })

  it('mints an id only when a legacy element is explicitly enrolled', async () => {
    const factory = makeFactory()
    const opened = await openPptx(fixture())
    const slide = opened.deck.slides[0]!
    const legacy = slide.elements[0]!
    expect(legacy.creationId).toBeUndefined()

    const created = ensureElementCreationId(slide, legacy, factory)
    expect(created).toBe('{00000000-0000-4000-8000-000000000001}')
    expect(legacy.anchor.originalXml).toContain(
      'xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main"',
    )
    expect(legacy.anchor.originalXml).toMatch(
      /<a16:creationId\b[^>]*\bid="\{00000000-0000-4000-8000-000000000001\}"\/>/,
    )
  })

  it('resolves by durable ids and fails closed for stale, missing, and ambiguous targets', async () => {
    const factory = makeFactory()
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const element = addElement(
      slide,
      { kind: 'rect', offset: OFF_A },
      { creationIdFactory: factory },
    )
    const base = { slideId: slide.durableId, elementId: element.creationId! }
    const resolved = await resolvePresentationTarget(opened.deck, base)
    expect(resolved.status).toBe('resolved')
    if (resolved.status !== 'resolved') throw new Error('expected resolution')

    const slideFingerprint = await fingerprintSlide(slide)
    expect(
      await resolvePresentationTarget(opened.deck, {
        slideId: slide.durableId,
        expectedFingerprint: slideFingerprint,
      }),
    ).toMatchObject({ status: 'resolved', fingerprint: slideFingerprint })

    expect(
      await resolvePresentationTarget(opened.deck, { ...base, expectedType: 'text' }),
    ).toMatchObject({ status: 'conflict', code: 'target_stale' })
    expect(
      await resolvePresentationTarget(opened.deck, {
        ...base,
        expectedFingerprint: 'sha256:' + '0'.repeat(64),
      }),
    ).toMatchObject({ status: 'conflict', code: 'target_stale' })
    expect(
      await resolvePresentationTarget(opened.deck, {
        ...base,
        elementId: '{00000000-0000-4000-8000-000000000099}',
      }),
    ).toMatchObject({ status: 'conflict', code: 'target_missing' })
    expect(
      await resolvePresentationTarget(opened.deck, { ...base, slideId: 'ppt/slides/missing.xml' }),
    ).toMatchObject({ status: 'conflict', code: 'target_missing' })

    const duplicate = addElement(
      slide,
      { kind: 'ellipse', offset: OFF_B },
      { creationIdFactory: factory },
    )
    duplicate.creationId = element.creationId
    duplicate.anchor.originalXml = duplicate.anchor.originalXml.replace(
      /<a16:creationId id="[^"]+"\/>>?/,
      `<a16:creationId id="${element.creationId}"/>`,
    )
    expect(await resolvePresentationTarget(opened.deck, base)).toMatchObject({
      status: 'conflict',
      code: 'target_ambiguous',
    })
  })

  it('fails closed when slide or element durable ids are absent', async () => {
    const opened = await openPptx(fixture())
    const slide = opened.deck.slides[0]!
    const legacy = slide.elements[0]!
    expect(
      await resolvePresentationTarget(opened.deck, {
        slideId: slide.durableId,
        elementId: legacy.id,
      }),
    ).toMatchObject({ status: 'conflict', code: 'target_missing' })
  })
})
