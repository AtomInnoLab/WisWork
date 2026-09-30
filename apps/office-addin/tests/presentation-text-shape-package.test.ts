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
  zip.file(
    'ppt/slides/slide1.xml',
    slide.replace(
      shape!,
      shape!
        .replace(/<a:t>[^<]*<\/a:t>/, '<a:t>Different</a:t>')
        .replace(/<a:off x="\d+" y="\d+"\/>/, '<a:off x="12345" y="67890"/>'),
    ),
  )
  const textAndGeometryChanged = (
    await inspectPowerPointTextShapeFingerprints(await zip.generateAsync({ type: 'base64' }), [id!])
  )[id!]!
  expect(textAndGeometryChanged.exact).not.toBe(before.exact)
  expect(textAndGeometryChanged.formatting).not.toBe(before.formatting)
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

it('fingerprints a connector and a group with their children in a real PPTX package', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const zip = await JSZip.loadAsync((await compilePresentationDeck(deck)).bytes)
  const original = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const connector =
    '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="9001" name="connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr></p:cxnSp>'
  const group =
    '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9002" name="group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/><a:chOff x="0" y="0"/><a:chExt cx="300" cy="400"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="9003" name="child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp></p:grpSp>'
  const slide = original.replace('</p:spTree>', `${connector}${group}</p:spTree>`)
  const inspect = async (xml: string) => {
    zip.file('ppt/slides/slide1.xml', xml)
    return inspectPowerPointTextShapeFingerprints(
      await zip.generateAsync({ type: 'base64' }),
      ['9001', '9002'],
      undefined,
      true,
    )
  }
  const before = await inspect(slide)
  const moved = await inspect(
    slide.replace(connector, connector.replace('x="100" y="200"', 'x="101" y="201"')),
  )
  expect(moved['9001']!.exact).not.toBe(before['9001']!.exact)
  expect(moved['9001']!.content).toBe(before['9001']!.content)
  expect(moved['9002']!.exact).toBe(before['9002']!.exact)
  const movedGroup = await inspect(
    slide.replace(group, group.replace('x="100" y="200"', 'x="101" y="201"')),
  )
  expect(movedGroup['9002']!.content).toBe(before['9002']!.content)
  const changedChild = await inspect(
    slide.replace(group, group.replace('prst="rect"', 'prst="ellipse"')),
  )
  expect(changedChild['9002']!.content).not.toBe(before['9002']!.content)
  const changedConnector = await inspect(
    slide.replace(connector, connector.replace('val="000000"', 'val="FF0000"')),
  )
  expect(changedConnector['9001']!.content).not.toBe(before['9001']!.content)
  const childPicture =
    '<p:pic><p:nvPicPr><p:cNvPr id="9004" name="group-picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdWisGroupImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
  const withPicture = slide.replace(group, group.replace('</p:grpSp>', `${childPicture}</p:grpSp>`))
  const relsPath = 'ppt/slides/_rels/slide1.xml.rels'
  const rels = await zip.file(relsPath)!.async('string')
  zip.file(
    relsPath,
    rels.replace(
      '</Relationships>',
      '<Relationship Id="rIdWisGroupImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/wiswork-group.png"/></Relationships>',
    ),
  )
  zip.file('ppt/media/wiswork-group.png', new Uint8Array([1, 2, 3]))
  const pictureBefore = (await inspect(withPicture))['9002']!
  zip.file('ppt/media/wiswork-group.png', new Uint8Array([1, 2, 4]))
  const pictureAfter = (await inspect(withPicture))['9002']!
  expect(pictureAfter.exact).not.toBe(pictureBefore.exact)
  expect(pictureAfter.content).not.toBe(pictureBefore.content)
  await expect(
    inspectPowerPointTextShapeFingerprints(
      await zip.generateAsync({ type: 'base64' }),
      ['9003'],
      undefined,
      true,
    ),
  ).rejects.toThrow('office_api_unsupported')
})

it('protects an unsupported graphic frame and fails closed when its package mapping is missing', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const zip = await JSZip.loadAsync((await compilePresentationDeck(deck)).bytes)
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const graphic =
    '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9010" name="diagram"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rIdWisDiagram"/></a:graphicData></a:graphic></p:graphicFrame>'
  zip.file('ppt/slides/slide1.xml', slide.replace('</p:spTree>', `${graphic}</p:spTree>`))
  const relsPath = 'ppt/slides/_rels/slide1.xml.rels'
  const rels = await zip.file(relsPath)!.async('string')
  zip.file(
    relsPath,
    rels.replace(
      '</Relationships>',
      '<Relationship Id="rIdWisDiagram" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/dataWiswork.xml"/></Relationships>',
    ),
  )
  zip.file('ppt/diagrams/dataWiswork.xml', '<diagram>first</diagram>')
  const inspect = async () =>
    (
      await inspectPowerPointTextShapeFingerprints(
        await zip.generateAsync({ type: 'base64' }),
        ['9010'],
        undefined,
        true,
      )
    )['9010']!
  const before = await inspect()
  zip.file('ppt/diagrams/dataWiswork.xml', '<diagram>second</diagram>')
  const after = await inspect()
  expect(after.exact).not.toBe(before.exact)
  zip.file(
    'ppt/diagrams/_rels/dataWiswork.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdWisColor" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="colorsWiswork.xml"/></Relationships>',
  )
  zip.file('ppt/diagrams/colorsWiswork.xml', '<colors>blue</colors>')
  const withColors = await inspect()
  zip.file('ppt/diagrams/colorsWiswork.xml', '<colors>red</colors>')
  const changedColors = await inspect()
  expect(changedColors.exact).not.toBe(withColors.exact)
  zip.remove('ppt/diagrams/colorsWiswork.xml')
  await expect(inspect()).rejects.toThrow('office_api_unsupported')
  zip.file('ppt/diagrams/colorsWiswork.xml', '<colors>red</colors>')
  zip.file(
    'ppt/diagrams/_rels/colorsWiswork.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdWisBack" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="dataWiswork.xml"/></Relationships>',
  )
  expect((await inspect()).exact).toMatch(/^[a-f0-9]{64}$/)
  await expect(
    inspectPowerPointTextShapeFingerprints(
      await zip.generateAsync({ type: 'base64' }),
      ['9011'],
      undefined,
      true,
    ),
  ).rejects.toThrow('office_api_unsupported')
})

it('detects a shared theme change even when the shape XML is unchanged', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const zip = await JSZip.loadAsync((await compilePresentationDeck(deck)).bytes)
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const id = slide.match(/<p:sp\b[^]*?<p:cNvPr id="(\d+)"/)?.[1]
  const themePath = Object.keys(zip.files).find((path) => /^ppt\/theme\/theme\d+\.xml$/.test(path))
  expect(id).toBeDefined()
  expect(themePath).toBeDefined()
  const inspect = async () =>
    (
      await inspectPowerPointTextShapeFingerprints(
        await zip.generateAsync({ type: 'base64' }),
        [id!],
        undefined,
        true,
      )
    )[id!]!
  const before = await inspect()
  const theme = await zip.file(themePath!)!.async('string')
  const changed = theme.replace(/<a:srgbClr val="[0-9A-Fa-f]{6}"\/>/, '<a:srgbClr val="ABCDEF"/>')
  expect(changed).not.toBe(theme)
  zip.file(themePath!, changed)
  const after = await inspect()
  expect(after.exact).not.toBe(before.exact)
  expect(after.formatting).not.toBe(before.formatting)
})
