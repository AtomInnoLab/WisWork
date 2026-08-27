import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { PassthroughElement } from '../src/types'
import {
  addElement,
  addPicture,
  createBlankPptx,
  duplicateSlide,
  ensureElementCreationId,
  fingerprintSlide,
  fingerprintSlideElement,
  groupElements,
  openPptx,
  copyElementData,
  moveSlide,
  normalizeCreationId,
  pasteElements,
  reorderElement,
  resolvePresentationTarget,
  resolvePresentationContainer,
  savePptx,
  setCreationIdInElementXml,
  setSlideHidden,
  setSlideNotes,
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
    const fingerprint = await fingerprintSlideElement(opened, slide, element)
    const base = {
      slideId: slide.durableId,
      elementId: element.creationId!,
      expectedType: 'shape' as const,
      expectedFingerprint: fingerprint,
    }
    const resolved = await resolvePresentationTarget(opened, base)
    expect(resolved.status).toBe('resolved')
    if (resolved.status !== 'resolved') throw new Error('expected resolution')

    const slideFingerprint = await fingerprintSlide(opened, slide)
    expect(
      await resolvePresentationTarget(opened, {
        slideId: slide.durableId,
        expectedFingerprint: slideFingerprint,
      }),
    ).toMatchObject({ status: 'resolved', fingerprint: slideFingerprint })

    expect(
      await resolvePresentationTarget(opened, { ...base, expectedType: 'text' }),
    ).toMatchObject({ status: 'conflict', code: 'target_stale' })
    expect(
      await resolvePresentationTarget(opened, {
        ...base,
        expectedFingerprint: 'sha256:' + '0'.repeat(64),
      }),
    ).toMatchObject({ status: 'conflict', code: 'target_stale' })
    expect(
      await resolvePresentationTarget(opened, {
        ...base,
        elementId: '{00000000-0000-4000-8000-000000000099}',
      }),
    ).toMatchObject({ status: 'conflict', code: 'target_missing' })
    expect(
      await resolvePresentationTarget(opened, { ...base, slideId: 'ppt/slides/missing.xml' }),
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
    expect(await resolvePresentationTarget(opened, base)).toMatchObject({
      status: 'conflict',
      code: 'target_ambiguous',
    })
  })

  it('fails closed when slide or element durable ids are absent', async () => {
    const opened = await openPptx(fixture())
    const slide = opened.deck.slides[0]!
    const legacy = slide.elements[0]!
    expect(
      await resolvePresentationTarget(opened, {
        slideId: slide.durableId,
        elementId: legacy.id,
        expectedType: 'shape',
        expectedFingerprint: 'sha256:' + '0'.repeat(64),
      }),
    ).toMatchObject({ status: 'conflict', code: 'target_missing' })
  })

  it('requires mutation preconditions and keeps add-only container resolution explicit', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const element = addElement(
      slide,
      { kind: 'rect', offset: OFF_A },
      { creationIdFactory: makeFactory() },
    )
    const fingerprint = await fingerprintSlideElement(opened, slide, element)

    expect(
      await resolvePresentationTarget(opened, {
        slideId: slide.durableId,
        elementId: element.creationId,
      }),
    ).toMatchObject({ status: 'conflict', code: 'target_stale' })
    expect(
      await resolvePresentationTarget(opened, {
        slideId: slide.durableId,
        elementId: element.creationId,
        expectedType: 'shape',
      }),
    ).toMatchObject({ status: 'conflict', code: 'target_stale' })
    expect(await resolvePresentationTarget(opened, { slideId: slide.durableId })).toMatchObject({
      status: 'conflict',
      code: 'target_stale',
    })
    expect(
      await resolvePresentationContainer(opened.deck, { slideId: slide.durableId }),
    ).toMatchObject({ status: 'resolved', slide })
    expect(
      await resolvePresentationTarget(opened, {
        slideId: slide.durableId,
        elementId: element.creationId,
        expectedType: 'shape',
        expectedFingerprint: fingerprint,
      }),
    ).toMatchObject({ status: 'resolved', element })
  })

  it('fingerprints visual effects and actual bytes behind picture paths', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const picture = addPicture(
      opened,
      slide,
      { bytes: new Uint8Array([1, 2, 3]), ext: 'png', offset: OFF_A },
      { creationIdFactory: makeFactory() },
    )!
    const before = await fingerprintSlideElement(opened, slide, picture)
    picture.opacity = 0.5
    const opacityChanged = await fingerprintSlideElement(opened, slide, picture)
    expect(opacityChanged).not.toBe(before)
    opened.archive.entries.set(picture.mediaRef, new Uint8Array([3, 2, 1]))
    const bytesChanged = await fingerprintSlideElement(opened, slide, picture)
    expect(bytesChanged).not.toBe(opacityChanged)
  })

  it('binds slide fingerprints to slide flags and speaker notes', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const before = await fingerprintSlide(opened, slide)
    setSlideHidden(slide, true)
    const hidden = await fingerprintSlide(opened, slide)
    expect(hidden).not.toBe(before)
    expect(setSlideNotes(opened, 0, 'durable note')).toBe(true)
    expect(await fingerprintSlide(opened, slide)).not.toBe(hidden)
  })

  it('fails closed instead of weakly fingerprinting passthrough content', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const passthrough: PassthroughElement = {
      id: 'unsafe',
      creationId: '{00000000-0000-4000-8000-000000000001}',
      type: 'passthrough',
      kind: 'unknown',
      anchor: { spIndex: 0, originalXml: '<p:graphicFrame/>', range: [0, 0] },
      transform: { offset: OFF_A, rot: 0, flipH: false, flipV: false },
    }
    await expect(fingerprintSlideElement(opened, slide, passthrough)).rejects.toThrow(
      'cannot be fingerprinted safely',
    )
  })

  it('retries creationId collisions and exhausts without mutation', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const first = addElement(
      slide,
      { kind: 'rect', offset: OFF_A },
      { creationIdFactory: makeFactory() },
    )
    let attempts = 0
    const retryFactory = () => {
      attempts += 1
      return attempts === 1 ? first.creationId! : '{00000000-0000-4000-8000-000000000099}'
    }
    const second = addElement(
      slide,
      { kind: 'ellipse', offset: OFF_B },
      { creationIdFactory: retryFactory },
    )
    expect(attempts).toBe(2)
    expect(second.creationId).toBe('{00000000-0000-4000-8000-000000000099}')

    const beforeCount = slide.elements.length
    const beforeXml = slide.elements.map((element) => element.anchor.originalXml)
    expect(() =>
      addElement(
        slide,
        { kind: 'rect', offset: OFF_A },
        { creationIdFactory: () => first.creationId! },
      ),
    ).toThrow('Unable to mint a unique creationId')
    expect(slide.elements).toHaveLength(beforeCount)
    expect(slide.elements.map((element) => element.anchor.originalXml)).toEqual(beforeXml)
  })

  it('preflights duplicate-slide ids before mutating the package', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const element = addElement(
      slide,
      { kind: 'rect', offset: OFF_A },
      { creationIdFactory: makeFactory() },
    )
    const beforeSlides = opened.deck.slides.length
    const beforeEntries = new Map(opened.archive.entries)
    expect(() =>
      duplicateSlide(opened, 0, { creationIdFactory: () => element.creationId! }),
    ).toThrow('Unable to mint a unique creationId')
    expect(opened.deck.slides).toHaveLength(beforeSlides)
    expect([...opened.archive.entries.keys()]).toEqual([...beforeEntries.keys()])
    for (const [path, bytes] of beforeEntries) expect(opened.archive.entries.get(path)).toBe(bytes)
  })

  it('preflights group identity before moving any child', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const factory = makeFactory()
    const first = addElement(slide, { kind: 'rect', offset: OFF_A }, { creationIdFactory: factory })
    const second = addElement(
      slide,
      { kind: 'ellipse', offset: OFF_B },
      { creationIdFactory: factory },
    )
    const beforeIds = slide.elements.map((element) => element.id)
    const beforeXml = slide.elements.map((element) => element.anchor.originalXml)
    expect(() =>
      groupElements(opened, 0, [first.id, second.id], {
        creationIdFactory: () => first.creationId!,
      }),
    ).toThrow('Unable to mint a unique creationId')
    expect(slide.elements.map((element) => element.id)).toEqual(beforeIds)
    expect(slide.elements.map((element) => element.anchor.originalXml)).toEqual(beforeXml)
  })

  it('preflights pasted element batches and retries source-id collisions', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const source = addElement(
      slide,
      { kind: 'rect', offset: OFF_A },
      { creationIdFactory: makeFactory() },
    )
    const item = copyElementData(opened, slide, source)
    let attempts = 0
    const pasted = pasteElements(
      opened,
      0,
      [item],
      { dx: 1, dy: 1 },
      {
        creationIdFactory: () => {
          attempts += 1
          return attempts === 1 ? source.creationId! : '{00000000-0000-4000-8000-000000000099}'
        },
      },
    )
    expect(attempts).toBe(2)
    expect(pasted?.slide.elements.at(-1)?.creationId).toBe('{00000000-0000-4000-8000-000000000099}')

    const stableSlide = pasted!.slide
    const beforeIds = stableSlide.elements.map((element) => element.id)
    expect(() =>
      pasteElements(
        opened,
        0,
        [item],
        { dx: 2, dy: 2 },
        {
          creationIdFactory: () => source.creationId!,
        },
      ),
    ).toThrow('Unable to mint a unique creationId')
    expect(stableSlide.elements.map((element) => element.id)).toEqual(beforeIds)
  })
})
