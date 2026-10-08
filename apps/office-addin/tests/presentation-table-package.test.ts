import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { inspectPowerPointTableFingerprints } from '../src/skills/powerpoint/presentation-table-package'

it('detects a table style change in a real compiled single-page PPTX without changing cell text', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[5]!]
  const zip = await JSZip.loadAsync((await compilePresentationDeck(deck)).bytes)
  const base64 = await zip.generateAsync({ type: 'base64' })
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const table = slide.match(
    /<p:graphicFrame\b[^]*?<a:tbl>[^]*?<\/a:tbl>[^]*?<\/p:graphicFrame>/,
  )?.[0]
  expect(table).toBeDefined()
  const id = table!.match(/<p:cNvPr id="(\d+)"/)?.[1]
  expect(id).toBeDefined()
  const before = await inspectPowerPointTableFingerprints(base64, [id!])
  const styled = slide.replace(
    /<a:tblPr\b([^>]*?)(\/?)>/,
    (_match, attributes: string, closing: string) =>
      `<a:tblPr${attributes.replace(/\sfirstRow="[01]"/, '')} firstRow="${attributes.includes('firstRow="1"') ? '0' : '1'}"${closing}>`,
  )
  expect(styled).not.toBe(slide)
  zip.file('ppt/slides/slide1.xml', styled)
  const after = await inspectPowerPointTableFingerprints(
    await zip.generateAsync({ type: 'base64' }),
    [id!],
  )
  expect(after[id!]).not.toBe(before[id!])
  await expect(inspectPowerPointTableFingerprints(base64, ['999999'])).rejects.toThrow(
    'office_api_unsupported',
  )
})

it('fingerprints 101 tables on a bounded page', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[5]!]
  const zip = await JSZip.loadAsync((await compilePresentationDeck(deck)).bytes)
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const table = slide.match(
    /<p:graphicFrame\b[^]*?<a:tbl>[^]*?<\/a:tbl>[^]*?<\/p:graphicFrame>/,
  )?.[0]
  expect(table).toBeDefined()
  const rows = [...table!.matchAll(/<a:tr\b[^]*?<\/a:tr>/g)].map(([row]) => row)
  expect(rows.length).toBeGreaterThan(1)
  const compact = rows.slice(1).reduce((value, row) => value.replace(row, ''), table!)
  const ids = Array.from({ length: 101 }, (_, index) => String(5000 + index))
  zip.file(
    'ppt/slides/slide1.xml',
    slide.replace(
      '</p:spTree>',
      `${ids.map((id) => compact.replace(/<p:cNvPr id="\d+"/, `<p:cNvPr id="${id}"`)).join('')}</p:spTree>`,
    ),
  )
  expect(
    Object.keys(
      await inspectPowerPointTableFingerprints(await zip.generateAsync({ type: 'base64' }), ids),
    ),
  ).toHaveLength(101)
})
