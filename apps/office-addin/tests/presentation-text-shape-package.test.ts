import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { inspectPowerPointTextShapeFingerprints } from '../src/skills/powerpoint/presentation-rich-text-package'

it('distinguishes intended text replacement from hidden run formatting drift in a real PPTX', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const zip = await JSZip.loadAsync((await compilePresentationDeck(deck)).bytes)
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const shape = slide.match(/<p:sp\b[^]*?<p:txBody\b[^]*?<\/p:sp>/)?.[0]
  const id = shape?.match(/<p:cNvPr id="(\d+)"/)?.[1]
  expect(id).toBeDefined()
  const base64 = await zip.generateAsync({ type: 'base64' })
  const before = (await inspectPowerPointTextShapeFingerprints(base64, [id!]))[id!]!
  const newText = slide.replace(/<a:t>[^<]*<\/a:t>/, '<a:t>Different</a:t>')
  expect(newText).not.toBe(slide)
  zip.file('ppt/slides/slide1.xml', newText)
  const textChanged = (
    await inspectPowerPointTextShapeFingerprints(await zip.generateAsync({ type: 'base64' }), [id!])
  )[id!]!
  expect(textChanged.content).not.toBe(before.content)
  expect(textChanged.formatting).toBe(before.formatting)
  zip.file('ppt/slides/slide1.xml', slide.replace(/<a:t>[^<]*<\/a:t>/, '<a:t> </a:t>'))
  const blanked = (
    await inspectPowerPointTextShapeFingerprints(await zip.generateAsync({ type: 'base64' }), [id!])
  )[id!]!
  expect(blanked.content).not.toBe(before.content)
  expect(blanked.formatting).toBe(before.formatting)
  const changedRun = slide.replace(
    /<a:rPr\b([^>]*?)(\/?)>/,
    (_match, attributes: string, closing: string) =>
      `<a:rPr${attributes.replace(/\slang="[^"]*"/, '')} lang="fr-FR"${closing}>`,
  )
  expect(changedRun).not.toBe(slide)
  zip.file('ppt/slides/slide1.xml', changedRun)
  const formatChanged = (
    await inspectPowerPointTextShapeFingerprints(await zip.generateAsync({ type: 'base64' }), [id!])
  )[id!]!
  expect(formatChanged.formatting).not.toBe(before.formatting)
  const linkedSlide = slide.replace(
    /<a:rPr\b([^>]*?)(\/?)>/,
    (_match, attributes: string, closing: string) =>
      closing
        ? `<a:rPr${attributes}><a:hlinkClick r:id="rIdWisLink"/></a:rPr>`
        : `<a:rPr${attributes}><a:hlinkClick r:id="rIdWisLink"/>`,
  )
  const relsPath = 'ppt/slides/_rels/slide1.xml.rels'
  const rels = await zip.file(relsPath)!.async('string')
  const linkedRels = rels.replace(
    '</Relationships>',
    '<Relationship Id="rIdWisLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/first" TargetMode="External"/></Relationships>',
  )
  zip.file('ppt/slides/slide1.xml', linkedSlide)
  zip.file(relsPath, linkedRels)
  const linkedBefore = (
    await inspectPowerPointTextShapeFingerprints(await zip.generateAsync({ type: 'base64' }), [id!])
  )[id!]!
  zip.file(relsPath, linkedRels.replace('https://example.com/first', 'https://example.com/second'))
  const linkedAfter = (
    await inspectPowerPointTextShapeFingerprints(await zip.generateAsync({ type: 'base64' }), [id!])
  )[id!]!
  expect(linkedAfter.formatting).not.toBe(linkedBefore.formatting)
})

it('maps an empty editable text box to an exact package fingerprint', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const zip = await JSZip.loadAsync((await compilePresentationDeck(deck)).bytes)
  const original = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const body = [...original.matchAll(/<p:sp\b[^]*?<\/p:sp>/g)].find((match) =>
    match[0].includes('name="body"'),
  )?.[0]
  expect(body).toBeDefined()
  const slide = original.replace(body!, body!.replace(/<a:t>[^<]*<\/a:t>/, '<a:t></a:t>'))
  zip.file('ppt/slides/slide1.xml', slide)
  const shape = [...slide.matchAll(/<p:sp\b[^]*?<\/p:sp>/g)].find((match) =>
    match[0].includes('name="body"'),
  )?.[0]
  expect(shape).toContain('<p:txBody>')
  const id = shape?.match(/<p:cNvPr id="(\d+)"/)?.[1]
  expect(id).toBeDefined()
  const inspected = await inspectPowerPointTextShapeFingerprints(
    await zip.generateAsync({ type: 'base64' }),
    [id!],
  )
  expect(inspected[id!]?.content).toMatch(/^[a-f0-9]{64}$/)
})

it('protects a native geometric shape style while permitting only its geometry change', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[3]!]
  const zip = await JSZip.loadAsync((await compilePresentationDeck(deck)).bytes)
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const shape = [...slide.matchAll(/<p:sp\b[^]*?<\/p:sp>/g)].find((match) =>
    match[0].includes('name="step"'),
  )?.[0]
  expect(shape).toBeDefined()
  const id = shape?.match(/<p:cNvPr id="(\d+)"/)?.[1]
  expect(id).toBeDefined()
  const before = (
    await inspectPowerPointTextShapeFingerprints(
      await zip.generateAsync({ type: 'base64' }),
      [id!],
      undefined,
      true,
    )
  )[id!]!
  const moved = shape!.replace(/<a:off x="\d+" y="\d+"\/>/, '<a:off x="999999" y="999999"/>')
  expect(moved).not.toBe(shape)
  zip.file('ppt/slides/slide1.xml', slide.replace(shape!, moved))
  const geometry = (
    await inspectPowerPointTextShapeFingerprints(
      await zip.generateAsync({ type: 'base64' }),
      [id!],
      undefined,
      true,
    )
  )[id!]!
  expect(geometry.exact).not.toBe(before.exact)
  expect(geometry.content).toBe(before.content)
  const recolored = shape!.replace(
    /<a:srgbClr val="[0-9A-Fa-f]{6}"\/>/,
    '<a:srgbClr val="ABCDEF"/>',
  )
  expect(recolored).not.toBe(shape)
  zip.file('ppt/slides/slide1.xml', slide.replace(shape!, recolored))
  const style = (
    await inspectPowerPointTextShapeFingerprints(
      await zip.generateAsync({ type: 'base64' }),
      [id!],
      undefined,
      true,
    )
  )[id!]!
  expect(style.content).not.toBe(before.content)
})
