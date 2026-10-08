import { expect, it, vi } from 'vitest'
import { PNG } from 'pngjs'
import JSZip from 'jszip'
import { readFileSync } from 'node:fs'
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
it('rejects a referenced PNG with damaged image data before host import', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  const path = Object.keys(zip.files).find((name) => /^ppt\/media\/[^/]+\.png$/.test(name))
  expect(path).toBeDefined()
  const image = await zip.file(path!)!.async('nodebuffer')
  const damaged = Buffer.from(image)
  const idat = damaged.indexOf(Buffer.from('IDAT'))
  expect(idat).toBeGreaterThan(0)
  damaged[idat + 4] ^= 1
  zip.file(path!, damaged)
  await expect(
    validatePresentationImportSourcePage(
      (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64'),
      '256#',
    ),
  ).rejects.toThrow('presentation_import_state_invalid')
})
it('validates a compiled PNG through browser decoding when decompression streams are unavailable', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const { bytes } = await compilePresentationDeck(deck)
  vi.stubGlobal('DecompressionStream', undefined)
  vi.stubGlobal('createImageBitmap', async (blob: Blob) => {
    const image = PNG.sync.read(Buffer.from(await blob.arrayBuffer()))
    return { width: image.width, height: image.height, close() {} }
  })
  try {
    await expect(
      validatePresentationImportSourcePage(Buffer.from(bytes).toString('base64'), '256#'),
    ).resolves.toBeUndefined()
  } finally {
    vi.unstubAllGlobals()
  }
})
it('rejects a referenced image without its package content type', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  const path = '[Content_Types].xml'
  const xml = await zip.file(path)!.async('string')
  const changed = xml.replace(/<Default\b[^>]*Extension="png"[^>]*\/>/, '')
  expect(changed).not.toBe(xml)
  zip.file(path, changed)
  await expect(
    validatePresentationImportSourcePage(
      (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64'),
      '256#',
    ),
  ).rejects.toThrow('presentation_import_state_invalid')
})
it('rejects mismatched chart and workbook package content types', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[6]!]
  const { bytes } = await compilePresentationDeck(deck)
  const original = await JSZip.loadAsync(bytes)
  const xml = await original.file('[Content_Types].xml')!.async('string')
  for (const mime of [
    'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  ]) {
    const zip = await JSZip.loadAsync(bytes)
    const changed = xml.replace(mime, 'application/octet-stream')
    expect(changed).not.toBe(xml)
    zip.file('[Content_Types].xml', changed)
    await expect(
      validatePresentationImportSourcePage(
        (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64'),
        '256#',
      ),
    ).rejects.toThrow('presentation_import_state_invalid')
  }
})
it('rejects a slide without its presentation slide content type', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  const path = '[Content_Types].xml'
  const xml = await zip.file(path)!.async('string')
  const changed = xml.replace(/<Override\b[^>]*PartName="\/ppt\/slides\/slide1\.xml"[^>]*\/>/, '')
  expect(changed).not.toBe(xml)
  zip.file(path, changed)
  await expect(
    validatePresentationImportSourcePage(
      (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64'),
      '256#',
    ),
  ).rejects.toThrow('presentation_import_state_invalid')
})
it('rejects an unrelated embedded object relationship before host import', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  const path = 'ppt/slides/_rels/slide1.xml.rels'
  const xml = await zip.file(path)!.async('string')
  zip.file('ppt/embeddings/unrelated.bin', new Uint8Array([1, 2, 3]))
  zip.file(
    path,
    xml.replace(
      '</Relationships>',
      '<Relationship Id="rIdUnrelated" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/unrelated.bin"/></Relationships>',
    ),
  )
  await expect(
    validatePresentationImportSourcePage(
      (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64'),
      '256#',
    ),
  ).rejects.toThrow('presentation_import_state_invalid')
})
it('resolves single-quoted image references and rejects a missing relationship', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  const path = 'ppt/slides/slide1.xml'
  const original = await zip.file(path)!.async('string')
  const quoted = original.replace(/r:embed="([^"]+)"/, "r:embed='$1'")
  expect(quoted).not.toBe(original)
  zip.file(path, quoted)
  await expect(
    validatePresentationImportSourcePage(
      (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64'),
      '256#',
    ),
  ).resolves.toBeUndefined()
  zip.file(path, quoted.replace(/r:embed='[^']+'/, "r:embed='rIdMissing'"))
  await expect(
    validatePresentationImportSourcePage(
      (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64'),
      '256#',
    ),
  ).rejects.toThrow('presentation_import_state_invalid')
})
it('accepts a referenced HTTPS citation link but rejects executable or unused external links', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  const slidePath = 'ppt/slides/slide1.xml',
    relsPath = 'ppt/slides/_rels/slide1.xml.rels',
    slideXml = await zip.file(slidePath)!.async('string'),
    relsXml = await zip.file(relsPath)!.async('string')
  const linkedSlide = slideXml.replace(
    /<p:cNvPr([^>]*)\/>/,
    '<p:cNvPr$1><a:hlinkClick r:id="rId999"/></p:cNvPr>',
  )
  expect(linkedSlide).not.toBe(slideXml)
  const relationship = (url: string) =>
    `<Relationship Id="rId999" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${url}" TargetMode="External"/>`
  const source = async (url: string, slide = linkedSlide) => {
    zip.file(slidePath, slide)
    zip.file(relsPath, relsXml.replace('</Relationships>', `${relationship(url)}</Relationships>`))
    return (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64')
  }
  await expect(
    validatePresentationImportSourcePage(await source('https://example.test/source'), '256#'),
  ).resolves.toBeUndefined()
  await expect(
    validatePresentationImportSourcePage(await source('javascript:alert(1)'), '256#'),
  ).rejects.toThrow('presentation_import_state_invalid')
  await expect(
    validatePresentationImportSourcePage(
      await source('https://example.test/source', slideXml),
      '256#',
    ),
  ).rejects.toThrow('presentation_import_state_invalid')
  await expect(
    validatePresentationImportSourcePage(
      await source(
        'https://example.test/source',
        linkedSlide.replace('</p:spTree>', '<a:blip r:embed="rId999"/></p:spTree>'),
      ),
      '256#',
    ),
  ).rejects.toThrow('presentation_import_state_invalid')
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
it('rejects a native chart when its visible cache has no verified workbook source', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[6]!]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  const path = Object.keys(zip.files).find((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name))!
  const xml = await zip.file(path)!.async('string')
  const cacheOnly = xml.replace(/<c:externalData\b[\s\S]*?<\/c:externalData>/, '')
  expect(cacheOnly).not.toBe(xml)
  zip.file(path, cacheOnly)
  const changed = (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64')
  await expect(validatePresentationImportSourcePage(changed, '256#')).rejects.toThrow(
    'presentation_import_state_invalid',
  )
})
it('rejects a chart shape the bounded source reader cannot verify', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[6]!]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  const path = Object.keys(zip.files).find((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name))!
  const xml = await zip.file(path)!.async('string')
  const series = /<c:ser>[\s\S]*?<\/c:ser>/.exec(xml)?.[0]
  expect(series).toBeDefined()
  zip.file(path, xml.replace(series!, series!.repeat(9)))
  await expect(
    validatePresentationImportSourcePage(
      (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64'),
      '256#',
    ),
  ).rejects.toThrow('presentation_import_state_invalid')
})

it('rejects a chart relationship after the native chart frame disappears', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[6]!]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  const path = 'ppt/slides/slide1.xml'
  const xml = await zip.file(path)!.async('string')
  const removed = xml.replace(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/, '')
  expect(removed).not.toBe(xml)
  zip.file(path, removed)
  await expect(
    validatePresentationImportSourcePage(
      (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64'),
      '256#',
    ),
  ).rejects.toThrow('presentation_import_state_invalid')
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
it('rejects a referenced PNG truncated after its IHDR prefix', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  const path = Object.keys(zip.files).find((name) => /^ppt\/media\/[^/]+\.png$/.test(name))!
  zip.file(path, (await zip.file(path)!.async('uint8array')).slice(0, 24))
  await expect(
    validatePresentationImportSourcePage(
      (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64'),
      '256#',
    ),
  ).rejects.toThrow('presentation_import_state_invalid')
})
it('rejects a JPEG containing only start and end markers', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  const path = 'ppt/slides/_rels/slide1.xml.rels'
  const xml = await zip.file(path)!.async('string')
  const target = /Target="\.\.\/media\/([^"/]+)\.png"/.exec(xml)
  expect(target).not.toBeNull()
  const changed = xml.replace(/Target="\.\.\/media\/([^"/]+)\.png"/, 'Target="../media/$1.jpg"')
  expect(changed).not.toBe(xml)
  zip.file(path, changed)
  zip.file(`ppt/media/${target![1]}.jpg`, new Uint8Array([0xff, 0xd8, 0xff, 0xd9]))
  await expect(
    validatePresentationImportSourcePage(
      (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64'),
      '256#',
    ),
  ).rejects.toThrow('presentation_import_state_invalid')
})
it('accepts a real JPEG image in a compiled native page', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  deck.assets[0] = {
    ...deck.assets[0]!,
    mime: 'image/jpeg',
    width: 1200,
    height: 2520,
    base64: readFileSync(
      new URL(
        '../../../docs/product/ppt-benchmark-materials/PPT-P0-11/naca-rm-l50b01-page-index.jpg',
        import.meta.url,
      ),
    ).toString('base64'),
  }
  const { bytes } = await compilePresentationDeck(deck)
  await expect(
    validatePresentationImportSourcePage(Buffer.from(bytes).toString('base64'), '256#'),
  ).resolves.toBeUndefined()
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
