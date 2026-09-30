import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { validatePresentationImportSourcePage } from '../src/skills/powerpoint/presentation-import-source-package'
import { inspectPowerPointChartSourcePackage } from '../src/skills/powerpoint/presentation-chart-source-package'

it('accepts a compiled native image page and rejects an external image relationship', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const { bytes } = await compilePresentationDeck(deck)
  const source = Buffer.from(bytes).toString('base64')
  await expect(validatePresentationImportSourcePage(source, '256#')).resolves.toBeUndefined()

  const zip = await JSZip.loadAsync(bytes)
  const path = 'ppt/slides/_rels/slide1.xml.rels'
  const xml = await zip.file(path)!.async('string')
  expect(xml).toContain('/image')
  const external = xml.replace(
    /(Type="[^"]*\/image"[^>]*?)Target="[^"]+"/,
    '$1Target="https://example.test/image.png" TargetMode="External"',
  )
  expect(external).not.toBe(xml)
  zip.file(path, external)
  const changed = (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64')
  await expect(validatePresentationImportSourcePage(changed, '256#')).rejects.toThrow(
    'presentation_import_state_invalid',
  )
})

it('rejects an external chart workbook even when the slide relationship remains local', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[6]!]
  const { bytes } = await compilePresentationDeck(deck)
  await expect(
    validatePresentationImportSourcePage(Buffer.from(bytes).toString('base64'), '256#'),
  ).resolves.toBeUndefined()

  const zip = await JSZip.loadAsync(bytes)
  const path = Object.keys(zip.files).find((name) =>
    /^ppt\/charts\/_rels\/chart\d+\.xml\.rels$/.test(name),
  )
  expect(path).toBeDefined()
  const xml = await zip.file(path!)!.async('string')
  const external = xml.replace(
    /Target="[^"]+"/,
    'Target="https://example.test/workbook.xlsx" TargetMode="External"',
  )
  expect(external).not.toBe(xml)
  zip.file(path!, external)
  const changed = (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64')
  await expect(validatePresentationImportSourcePage(changed, '256#')).rejects.toThrow(
    'presentation_import_state_invalid',
  )
})

it('rejects a referenced chart workbook whose bytes are not an XLSX package', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[6]!]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  const path = Object.keys(zip.files).find((name) => /^ppt\/embeddings\/[^/]+\.xlsx$/.test(name))
  expect(path).toBeDefined()
  zip.file(path!, 'not a workbook')
  await expect(
    validatePresentationImportSourcePage(
      (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64'),
      '256#',
    ),
  ).rejects.toThrow('presentation_import_state_invalid')
})

it('rejects a referenced XLSX package without a workbook definition', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[6]!]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  const path = Object.keys(zip.files).find((name) => /^ppt\/embeddings\/[^/]+\.xlsx$/.test(name))
  expect(path).toBeDefined()
  const workbook = await JSZip.loadAsync(await zip.file(path!)!.async('uint8array'))
  expect(workbook.file('xl/workbook.xml')).not.toBeNull()
  workbook.remove('xl/workbook.xml')
  zip.file(path!, await workbook.generateAsync({ type: 'uint8array' }))
  await expect(
    validatePresentationImportSourcePage(
      (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64'),
      '256#',
    ),
  ).rejects.toThrow('presentation_import_state_invalid')
})

it('rejects a chart whose embedded workbook data disagrees with its visible cache', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[6]!]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const shapeId = /<p:graphicFrame>[\s\S]*?<p:cNvPr id="(\d+)"/.exec(slideXml)?.[1]
  expect(shapeId).toBeDefined()
  const baseline = await inspectPowerPointChartSourcePackage(
    Buffer.from(bytes).toString('base64'),
    shapeId!,
    undefined,
    { allowAbsoluteChartTarget: true },
  )
  expect(baseline).toMatchObject({
    verification: 'matches',
    series: [{ categories: ['甲', '乙'] }],
  })
  const path = Object.keys(zip.files).find((name) => /^ppt\/embeddings\/[^/]+\.xlsx$/.test(name))!
  const workbook = await JSZip.loadAsync(await zip.file(path)!.async('uint8array'))
  const sheetXml = await workbook.file('xl/worksheets/sheet1.xml')!.async('string')
  const changedSheet = sheetXml.replace(
    /(<c r="B2"[^>]*><v>)[^<]+(<\/v>)/,
    (_match, open: string, close: string) => `${open}999999${close}`,
  )
  expect(changedSheet).not.toBe(sheetXml)
  workbook.file('xl/worksheets/sheet1.xml', changedSheet)
  zip.file(path, await workbook.generateAsync({ type: 'uint8array' }))
  const changed = (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64')
  expect(
    (
      await inspectPowerPointChartSourcePackage(changed, shapeId!, undefined, {
        allowAbsoluteChartTarget: true,
      })
    ).verification,
  ).toBe('mismatch')
  await expect(validatePresentationImportSourcePage(changed, '256#')).rejects.toThrow(
    'presentation_import_state_invalid',
  )
})

it('rejects a local image relationship redirected to an unrelated XML part', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  const path = 'ppt/slides/_rels/slide1.xml.rels'
  const xml = await zip.file(path)!.async('string')
  const changedXml = xml.replace(
    /(Type="[^"]*\/image"[^>]*?)Target="[^"]+"/,
    '$1Target="../theme/theme1.xml"',
  )
  expect(changedXml).not.toBe(xml)
  expect(zip.file('ppt/theme/theme1.xml')).not.toBeNull()
  zip.file(path, changedXml)
  const changed = (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64')
  await expect(validatePresentationImportSourcePage(changed, '256#')).rejects.toThrow(
    'presentation_import_state_invalid',
  )
})

it('rejects a chart relationship redirected to a different existing XML part', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[6]!]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  const path = 'ppt/slides/_rels/slide1.xml.rels'
  const xml = await zip.file(path)!.async('string')
  const changedXml = xml.replace(
    /(Type="[^"]*\/chart"[^>]*?)Target="[^"]+"/,
    '$1Target="../theme/theme1.xml"',
  )
  expect(changedXml).not.toBe(xml)
  zip.file(path, changedXml)
  const changed = (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64')
  await expect(validatePresentationImportSourcePage(changed, '256#')).rejects.toThrow(
    'presentation_import_state_invalid',
  )
})

it('rejects an empty image part even when its local relationship is intact', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  const path = Object.keys(zip.files).find((name) => /^ppt\/media\/[^/]+\.png$/.test(name))
  expect(path).toBeDefined()
  zip.file(path!, new Uint8Array())
  const changed = (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64')
  await expect(validatePresentationImportSourcePage(changed, '256#')).rejects.toThrow(
    'presentation_import_state_invalid',
  )
})

it('rejects non-image bytes stored under a referenced PNG filename', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  const path = Object.keys(zip.files).find((name) => /^ppt\/media\/[^/]+\.png$/.test(name))
  expect(path).toBeDefined()
  zip.file(path!, 'not a PNG image')
  const changed = (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64')
  await expect(validatePresentationImportSourcePage(changed, '256#')).rejects.toThrow(
    'presentation_import_state_invalid',
  )
})

it('rejects a chart part whose XML no longer contains a chart', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[6]!]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  const path = Object.keys(zip.files).find((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name))
  expect(path).toBeDefined()
  zip.file(
    path!,
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>',
  )
  const changed = (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64')
  await expect(validatePresentationImportSourcePage(changed, '256#')).rejects.toThrow(
    'presentation_import_state_invalid',
  )
})
