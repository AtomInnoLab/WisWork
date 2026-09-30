import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { inspectPowerPointChartFingerprints } from '../src/skills/powerpoint/presentation-chart-package'

it('detects bundled chart XML drift in a real compiled single-page PPTX', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[6]!]
  const zip = await JSZip.loadAsync((await compilePresentationDeck(deck)).bytes)
  const base64 = await zip.generateAsync({ type: 'base64' })
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const frame = slide.match(/<p:graphicFrame\b[^]*?<c:chart\b[^]*?<\/p:graphicFrame>/)?.[0]
  const id = frame?.match(/<p:cNvPr id="(\d+)"/)?.[1]
  expect(id).toBeDefined()
  const before = await inspectPowerPointChartFingerprints(base64, [id!])
  const chartPath = Object.keys(zip.files).find((path) => /^ppt\/charts\/chart\d+\.xml$/.test(path))
  expect(chartPath).toBeDefined()
  const chart = await zip.file(chartPath!)!.async('string')
  zip.file(chartPath!, chart.replace(/<c:v>([^<]+)<\/c:v>/, '<c:v>999</c:v>'))
  const after = await inspectPowerPointChartFingerprints(
    await zip.generateAsync({ type: 'base64' }),
    [id!],
  )
  expect(after[id!]).not.toBe(before[id!])
  await expect(inspectPowerPointChartFingerprints(base64, ['999999'])).rejects.toThrow(
    'office_api_unsupported',
  )
})

it('fingerprints 101 chart frames sharing a bounded chart resource', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[6]!]
  const zip = await JSZip.loadAsync((await compilePresentationDeck(deck)).bytes)
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const frame = slide.match(/<p:graphicFrame\b[^]*?<c:chart\b[^]*?<\/p:graphicFrame>/)?.[0]
  expect(frame).toBeDefined()
  const ids = Array.from({ length: 101 }, (_, index) => String(5000 + index))
  zip.file(
    'ppt/slides/slide1.xml',
    slide.replace(
      '</p:spTree>',
      `${ids.map((id) => frame!.replace(/<p:cNvPr id="\d+"/, `<p:cNvPr id="${id}"`)).join('')}</p:spTree>`,
    ),
  )
  expect(
    Object.keys(
      await inspectPowerPointChartFingerprints(await zip.generateAsync({ type: 'base64' }), ids),
    ),
  ).toHaveLength(101)
})

it('keeps the second chart fingerprint stable when only the first chart changes', async () => {
  const deck = benchmarkDeck()
  const page = structuredClone(deck.slides[6]!)
  const original = page.elements.find((element) => element.kind === 'chart')!
  page.elements.push({ ...original, id: 'second-chart', x: 9, y: 2.5, w: 4 })
  deck.slides = [page]
  const zip = await JSZip.loadAsync((await compilePresentationDeck(deck)).bytes)
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const ids = [...slide.matchAll(/<p:graphicFrame\b[^]*?<c:chart\b[^]*?<\/p:graphicFrame>/g)]
    .map((match) => match[0].match(/<p:cNvPr id="(\d+)"/)?.[1])
    .filter((id): id is string => id !== undefined)
  expect(ids).toHaveLength(2)
  const before = await inspectPowerPointChartFingerprints(
    await zip.generateAsync({ type: 'base64' }),
    ids,
  )
  const firstPath = Object.keys(zip.files)
    .filter((path) => /^ppt\/charts\/chart\d+\.xml$/.test(path))
    .sort()[0]!
  const firstChart = await zip.file(firstPath)!.async('string')
  zip.file(firstPath, firstChart.replace(/<c:v>([^<]+)<\/c:v>/, '<c:v>999</c:v>'))
  const after = await inspectPowerPointChartFingerprints(
    await zip.generateAsync({ type: 'base64' }),
    ids,
  )
  expect(Object.values(after).filter((value, index) => value !== before[ids[index]!])).toHaveLength(
    1,
  )
})
