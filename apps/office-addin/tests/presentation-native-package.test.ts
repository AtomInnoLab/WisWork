import { expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { BrowserPowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'

it('derives native text, picture, table and chart proofs from one exported page package', async () => {
  const deck = benchmarkDeck()
  const page = structuredClone(deck.slides[0]!)
  for (const index of [2, 5, 6])
    page.elements.push(structuredClone(deck.slides[index]!.elements[1]!))
  deck.slides = [page]
  const base64 = Buffer.from((await compilePresentationDeck(deck)).bytes).toString('base64')
  const zip = await JSZip.loadAsync(base64, { base64: true })
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const shapeIds = (tag: string, contains: string) =>
    [...slide.matchAll(new RegExp(`<${tag}\\b[^]*?<\\/${tag}>`, 'g'))]
      .filter((match) => match[0].includes(contains))
      .map((match) => match[0].match(/<p:cNvPr id="(\d+)"/)?.[1])
      .filter((id): id is string => id !== undefined)
  const ids = {
    pictures: shapeIds('p:pic', '<a:blip'),
    text: shapeIds('p:sp', '<p:txBody').slice(0, 1),
    tables: shapeIds('p:graphicFrame', '<a:tbl>'),
    charts: shapeIds('p:graphicFrame', '<c:chart'),
  }
  expect(Object.values(ids).map((values) => values.length)).toEqual([1, 1, 1, 1])
  const adapter = new BrowserPowerPointAdapter()
  const exported = vi.spyOn(adapter, 'exportPresentationPagePackage').mockResolvedValue({
    slideId: 'slide',
    slideIds: ['slide'],
    base64,
  })
  const proof = await adapter.inspectSlideNativePackage('slide', ids)
  expect(exported).toHaveBeenCalledTimes(1)
  expect(Object.keys(proof.pictures.fingerprints)).toEqual(ids.pictures)
  expect(Object.keys(proof.richText.fingerprints)).toEqual(ids.text)
  expect(Object.keys(proof.tables)).toEqual(ids.tables)
  expect(Object.keys(proof.charts)).toEqual(ids.charts)
  expect(proof.base64).toBe(base64)
})
