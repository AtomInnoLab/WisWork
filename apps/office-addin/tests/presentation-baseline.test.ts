import JSZip from 'jszip'
import PptxGenJS from 'pptxgenjs'
import { expect, it, vi } from 'vitest'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { createPresentationBaselineSkill } from '../src/skills/powerpoint/presentation-baseline'
import { inspectPowerPointPageNotes } from '../src/skills/powerpoint/presentation-notes-package'
import { inspectPowerPointSourceLinks } from '../src/skills/powerpoint/presentation-source-links-package'
import { inspectPowerPointRichText } from '../src/skills/powerpoint/presentation-rich-text-package'
const page = (slideId = 's1', text = 'original') => ({
  slideId,
  masterId: 'm',
  layoutId: 'l',
  shapes: [
    {
      id: 'shape',
      name: 'Title',
      type: 'TextBox',
      left: 10,
      top: 10,
      width: 100,
      height: 40,
      text,
      font: { name: 'Arial', size: 20, color: '#000000' },
    },
  ],
})
function fixture() {
  let documentId = 'doc',
    context = {
      slideIds: ['s1', 's2'],
      selectedSlideIds: ['s2'],
      selectedShapeIds: ['shape'],
      slideWidth: 960,
      slideHeight: 540,
    }
  const pages = new Map([
    ['s1', page()],
    ['s2', page('s2')],
  ])
  const adapter = {
    readContext: vi.fn(async () => structuredClone(context)),
    readPage: vi.fn(async (id: string) => structuredClone(pages.get(id)!)),
  }
  const inspectPage = vi.fn(async (slideId: string) => ({
    slideId,
    slideWidth: 960,
    slideHeight: 540,
    shapes: pages.get(slideId)!.shapes.map(({ text: _text, font: _font, ...shape }) => shape),
    shapesTruncated: false,
    overflows: [],
    overlaps: [],
    overlapsTruncated: false,
    screenshot: {
      mime: 'image/png' as const,
      base64:
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII=',
    },
  }))
  let packageBase64 = ''
  const exportPagePackage = vi.fn(async (slideId: string) => ({
    slideId,
    slideIds: [...context.slideIds],
    base64: packageBase64,
  }))
  const skill = createPresentationBaselineSkill({
    adapter,
    documentId: async () => documentId,
    inspectPage,
    exportPagePackage,
  })
  const call = (name: string, input: Record<string, unknown> = {}, signal?: AbortSignal) =>
    skill.executeTool({ id: 'call', name, input }, signal)
  const read = (scope = 'current') => call('read_presentation_baseline', { scope })
  return {
    skill,
    adapter,
    pages,
    inspectPage,
    exportPagePackage,
    call,
    read,
    setDocument: (id: string) => {
      documentId = id
    },
    setContext: (next: typeof context) => {
      context = next
    },
    getContext: () => context,
    setPackage: (base64: string) => {
      packageBase64 = base64
    },
  }
}
async function complexPagePackage(cell = 'North') {
  const zip = new JSZip()
  zip.file(
    'ppt/slides/slide1.xml',
    `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="Table"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>${cell}</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>`,
  )
  return zip.generateAsync({ type: 'base64' })
}
async function chartPagePackage() {
  const zip = new JSZip()
  zip.file(
    'ppt/slides/slide1.xml',
    '<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:c="urn:c" xmlns:r="urn:r"><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="Chart"/></p:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="rId5"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>',
  )
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    '<Relationships><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>',
  )
  zip.file(
    'ppt/charts/chart1.xml',
    '<c:chartSpace xmlns:c="urn:c"><c:chart><c:plotArea><c:barChart><c:ser><c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>12</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>',
  )
  return zip.generateAsync({ type: 'base64' })
}
async function notesPagePackage(
  note = 'First &amp; second',
  target = '../notesSlides/notesSlide1.xml',
) {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', '<p:sld><p:cSld><p:spTree/></p:cSld></p:sld>')
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    `<Relationships><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="${target}"/></Relationships>`,
  )
  zip.file(
    'ppt/notesSlides/notesSlide1.xml',
    `<p:notes><p:cSld><p:spTree>
    <p:sp><p:nvSpPr><p:nvPr><p:ph type="ftr"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Footer ignored</a:t></a:r></a:p></p:txBody></p:sp>
    <p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>${note}</a:t></a:r></a:p><a:p><a:r><a:t>Second paragraph</a:t></a:r></a:p></p:txBody></p:sp>
  </p:spTree></p:cSld></p:notes>`,
  )
  return zip.generateAsync({ type: 'base64' })
}
async function sourceLinksPackage(target = 'https://example.org/paper?x=1&amp;y=2') {
  const zip = new JSZip()
  zip.file(
    'ppt/slides/slide1.xml',
    `<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="urn:r"><p:cSld><p:spTree>
      <p:sp><p:nvSpPr><p:cNvPr id="7" name="Citation"><a:hlinkClick r:id="rId2"/></p:cNvPr></p:nvSpPr>
      <p:txBody><a:p><a:r><a:rPr><a:hlinkClick r:id="rId1"/></a:rPr><a:t>Original paper</a:t></a:r></a:p></p:txBody></p:sp>
    </p:spTree></p:cSld></p:sld>`,
  )
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    `<Relationships>
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" TargetMode="External" Target="${target}"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" TargetMode="External" Target="https://example.org/home"/>
    </Relationships>`,
  )
  return zip.generateAsync({ type: 'base64' })
}
it('detects hidden package changes when page integrity is requested', async () => {
  const f = fixture()
  f.setPackage(await complexPagePackage('North'))
  const read = await f.call('read_presentation_baseline', {
    scope: 'current',
    package_integrity: true,
  })
  expect(read.isError).toBeFalsy()
  const baseline = JSON.parse(read.output)
  expect(baseline.coverage.pagePackages).toBe('read')
  expect(baseline.pagePackageDigests.s2).toMatch(/^[a-f0-9]{64}$/)
  expect(
    JSON.parse(
      (
        await f.call('check_presentation_baseline', {
          baseline_id: baseline.baselineId,
        })
      ).output,
    ),
  ).toMatchObject({ unchanged: true, changedPackageSlideIds: [] })
  f.setPackage(await complexPagePackage('South'))
  expect(
    JSON.parse(
      (
        await f.call('check_presentation_baseline', {
          baseline_id: baseline.baselineId,
        })
      ).output,
    ),
  ).toMatchObject({ unchanged: false, changedPackageSlideIds: ['s2'] })
})
it('fingerprints every page in a selected multi-page baseline', async () => {
  const f = fixture()
  f.setContext({ ...f.getContext(), selectedSlideIds: ['s1', 's2'], selectedShapeIds: [] })
  f.setPackage(await complexPagePackage('North'))
  const baseline = JSON.parse(
    (
      await f.call('read_presentation_baseline', {
        scope: 'selected',
        package_integrity: true,
      })
    ).output,
  )
  expect(Object.keys(baseline.pagePackageDigests)).toEqual(['s1', 's2'])
  expect(f.exportPagePackage).toHaveBeenCalledTimes(4)
  const checked = JSON.parse(
    (
      await f.call('check_presentation_baseline', {
        baseline_id: baseline.baselineId,
      })
    ).output,
  )
  expect(checked).toMatchObject({ unchanged: true, changedPackageSlideIds: [] })
  expect(f.exportPagePackage).toHaveBeenCalledTimes(8)
})
it('retains a package digest for a host slide ID matching an object property name', async () => {
  const f = fixture()
  f.pages.set('__proto__', page('__proto__'))
  f.setContext({ ...f.getContext(), slideIds: ['__proto__'], selectedSlideIds: ['__proto__'] })
  f.setPackage(await complexPagePackage())
  const baseline = JSON.parse(
    (
      await f.call('read_presentation_baseline', {
        scope: 'current',
        package_integrity: true,
      })
    ).output,
  )
  expect(Object.hasOwn(baseline.pagePackageDigests, '__proto__')).toBe(true)
  expect(baseline.pagePackageDigests['__proto__']).toMatch(/^[a-f0-9]{64}$/)
})
it('rejects an unstable or unavailable requested page package without saving a baseline', async () => {
  const f = fixture()
  f.setPackage(await complexPagePackage('North'))
  const changed = await complexPagePackage('South')
  f.exportPagePackage
    .mockImplementationOnce(async (slideId: string) => ({
      slideId,
      slideIds: [...f.getContext().slideIds],
      base64: await complexPagePackage('North'),
    }))
    .mockImplementationOnce(async (slideId: string) => ({
      slideId,
      slideIds: [...f.getContext().slideIds],
      base64: changed,
    }))
  expect(
    (
      await f.call('read_presentation_baseline', {
        scope: 'current',
        package_integrity: true,
      })
    ).output,
  ).toBe('presentation_baseline_changed')
  expect(f.skill.snapshot('any')).toBeUndefined()
  const noExport = createPresentationBaselineSkill({
    adapter: f.adapter,
    documentId: async () => 'doc',
  })
  expect(
    (
      await noExport.executeTool({
        id: 'call',
        name: 'read_presentation_baseline',
        input: { scope: 'current', package_integrity: true },
      })
    ).output,
  ).toBe('office_api_unsupported')
})
it('reads slide source links with labels and package IDs without verifying targets', async () => {
  const f = fixture()
  f.setPackage(await sourceLinksPackage())
  const baseline = JSON.parse((await f.read()).output)
  const result = await f.call('read_presentation_baseline_source_links', {
    baseline_id: baseline.baselineId,
    slide_id: 's2',
  })
  expect(result.isError, result.output).not.toBe(true)
  expect(JSON.parse(result.output)).toMatchObject({
    status: 'read',
    sourceVerified: false,
    links: [
      {
        packageShapeId: '7',
        target: 'https://example.org/paper?x=1&y=2',
        label: 'Original paper',
        location: 'text_run',
        sourceVerified: false,
      },
      { packageShapeId: '7', location: 'shape_action', sourceVerified: false },
    ],
  })
  expect(f.exportPagePackage).toHaveBeenCalledTimes(2)
  expect(
    (
      await f.call('read_presentation_baseline_source_links', {
        baseline_id: baseline.baselineId,
        slide_id: 's1',
      })
    ).output,
  ).toBe('presentation_baseline_scope_mismatch')
})
it('ignores non-web links and rejects duplicate relationship IDs or changing exports', async () => {
  expect(
    await inspectPowerPointSourceLinks(await sourceLinksPackage('javascript:alert(1)')),
  ).toMatchObject({
    status: 'read',
    links: [{ target: 'https://example.org/home' }],
  })
  const duplicate = await sourceLinksPackage()
  const zip = await JSZip.loadAsync(duplicate, { base64: true })
  const path = 'ppt/slides/_rels/slide1.xml.rels'
  zip.file(path, (await zip.file(path)!.async('string')).replace('Id="rId2"', 'Id="rId1"'))
  await expect(
    inspectPowerPointSourceLinks(await zip.generateAsync({ type: 'base64' })),
  ).rejects.toThrow('office_api_unsupported')
  const f = fixture()
  const before = await sourceLinksPackage()
  const after = await sourceLinksPackage('https://example.org/changed')
  f.setPackage(before)
  const baseline = JSON.parse((await f.read()).output)
  f.exportPagePackage.mockImplementationOnce(async (slideId) => ({
    slideId,
    slideIds: ['s1', 's2'],
    base64: before,
  }))
  f.exportPagePackage.mockImplementationOnce(async (slideId) => ({
    slideId,
    slideIds: ['s1', 's2'],
    base64: after,
  }))
  expect(
    (
      await f.call('read_presentation_baseline_source_links', {
        baseline_id: baseline.baselineId,
        slide_id: 's2',
      })
    ).output,
  ).toBe('presentation_baseline_changed')
})
it('reads a PptxGenJS-authored hyperlink from a real slide package', async () => {
  const deck = new PptxGenJS()
  deck.addSlide().addText('Read the source', {
    x: 1,
    y: 1,
    w: 3,
    h: 0.5,
    hyperlink: { url: 'https://example.org/research' },
  })
  const bytes = await deck.write({ outputType: 'nodebuffer' })
  const report = await inspectPowerPointSourceLinks(
    Buffer.from(bytes as Uint8Array).toString('base64'),
  )
  expect(report.links).toEqual([
    expect.objectContaining({ target: 'https://example.org/research', sourceVerified: false }),
  ])
})
it('reads mixed direct run formatting and paragraph alignment on a stable baseline page', async () => {
  const zip = new JSZip()
  zip.file(
    'ppt/slides/slide1.xml',
    `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree>
    <p:sp><p:nvSpPr><p:cNvPr id="7" name="Citation"/></p:nvSpPr><p:txBody>
      <a:p><a:pPr algn="ctr"><a:defRPr sz="1400"/></a:pPr><a:r><a:rPr b="1" sz="1800"><a:latin typeface="Arial"/><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill></a:rPr><a:t>Bold</a:t></a:r><a:r><a:rPr i="1"/><a:t> &amp; italic</a:t></a:r><a:endParaRPr sz="1600"/></a:p>
      <a:p><a:r><a:t>Inherited</a:t></a:r><a:tab/><a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:rPr><a:t>Theme</a:t></a:r></a:p>
    </p:txBody></p:sp>
  </p:spTree></p:cSld></p:sld>`,
  )
  const packageBase64 = await zip.generateAsync({ type: 'base64' })
  const f = fixture()
  f.setPackage(packageBase64)
  const baseline = JSON.parse((await f.read()).output)
  const result = await f.call('read_presentation_baseline_rich_text', {
    baseline_id: baseline.baselineId,
    slide_id: 's2',
  })
  expect(result.isError, result.output).not.toBe(true)
  expect(JSON.parse(result.output)).toMatchObject({
    inheritanceResolved: false,
    qaPassed: false,
    shapes: [
      {
        packageShapeId: '7',
        paragraphs: [
          {
            alignment: 'ctr',
            paragraphDefaultFont: { sizePt: 14 },
            endParagraphFont: { sizePt: 16 },
            runs: [
              {
                text: 'Bold',
                directFont: { bold: true, sizePt: 18, color: '#FF0000', typeface: 'Arial' },
              },
              { text: ' & italic', directFont: { italic: true } },
            ],
          },
          {
            runs: [
              { text: 'Inherited', directFont: {} },
              { text: '\t', directFont: {} },
              { text: 'Theme', directFont: { themeColor: 'accent1' } },
            ],
          },
        ],
      },
    ],
  })
  expect(f.exportPagePackage).toHaveBeenCalledTimes(2)
  expect(
    (
      await f.call('read_presentation_baseline_rich_text', {
        baseline_id: baseline.baselineId,
        slide_id: 's1',
      })
    ).output,
  ).toBe('presentation_baseline_scope_mismatch')
})
it('merges local list, paragraph and run font evidence without claiming master inheritance', async () => {
  const zip = new JSZip()
  zip.file(
    'ppt/slides/slide1.xml',
    `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree>
    <p:sp><p:nvSpPr><p:cNvPr id="7" name="Styled"/></p:nvSpPr><p:txBody>
      <a:lstStyle><a:defPPr algn="r"><a:defRPr><a:latin typeface="Arial"/></a:defRPr></a:defPPr><a:lvl1pPr><a:defRPr sz="1200" b="1"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>
      <a:p><a:pPr lvl="0"><a:defRPr sz="1400"/></a:pPr>
        <a:r><a:t>Inherited</a:t></a:r>
        <a:r><a:rPr i="1"><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill></a:rPr><a:t>Direct</a:t></a:r>
        <a:br><a:rPr u="sng"/></a:br>
      </a:p>
    </p:txBody></p:sp>
  </p:spTree></p:cSld></p:sld>`,
  )
  const report = await inspectPowerPointRichText(await zip.generateAsync({ type: 'base64' }))
  expect(report.inheritanceResolved).toBe(false)
  expect(report.shapes[0]?.paragraphs[0]).toMatchObject({
    knownAlignment: 'r',
    listStyleFont: { sizePt: 12, bold: true, themeColor: 'accent1', typeface: 'Arial' },
    paragraphDefaultFont: { sizePt: 14 },
    runs: [
      {
        text: 'Inherited',
        directFont: {},
        knownFont: { sizePt: 14, bold: true, themeColor: 'accent1', typeface: 'Arial' },
      },
      {
        text: 'Direct',
        directFont: { italic: true, color: '#FF0000' },
        knownFont: { sizePt: 14, bold: true, italic: true, color: '#FF0000', typeface: 'Arial' },
      },
      {
        text: '\n',
        directFont: { underline: 'sng' },
        knownFont: {
          sizePt: 14,
          bold: true,
          underline: 'sng',
          themeColor: 'accent1',
          typeface: 'Arial',
        },
      },
    ],
  })
  expect(report.shapes[0]?.paragraphs[0]?.runs[1]?.knownFont).not.toHaveProperty('themeColor')
})
it('resolves only explicitly linked and unmodified theme colors', async () => {
  const zip = new JSZip()
  zip.file(
    'ppt/slides/slide1.xml',
    `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="7" name="Theme"/></p:nvSpPr><p:txBody><a:p><a:r><a:rPr><a:latin typeface="+mj-lt"/><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:rPr><a:t>Exact</a:t></a:r><a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"><a:tint val="50000"/></a:schemeClr></a:solidFill></a:rPr><a:t>Tinted</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`,
  )
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    `<Relationships><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
  )
  zip.file('ppt/slideLayouts/slideLayout1.xml', `<p:sldLayout xmlns:p="urn:p"/>`)
  zip.file(
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    `<Relationships><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
  )
  zip.file(
    'ppt/slideMasters/slideMaster1.xml',
    `<p:sldMaster xmlns:p="urn:p"><p:clrMap accent1="accent2"/></p:sldMaster>`,
  )
  zip.file(
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    `<Relationships><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>`,
  )
  zip.file(
    'ppt/theme/theme1.xml',
    `<a:theme xmlns:a="urn:a"><a:themeElements><a:clrScheme><a:accent2><a:srgbClr val="123456"/></a:accent2></a:clrScheme><a:fontScheme><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont></a:fontScheme></a:themeElements></a:theme>`,
  )
  const base64 = await zip.generateAsync({ type: 'base64' })
  const report = await inspectPowerPointRichText(base64)
  expect(report.shapes[0]?.paragraphs[0]?.runs).toEqual([
    expect.objectContaining({
      text: 'Exact',
      resolvedThemeColor: '#123456',
      resolvedThemeTypeface: 'Aptos Display',
    }),
    expect.objectContaining({
      text: 'Tinted',
      directFont: expect.objectContaining({ themeColor: 'accent1' }),
    }),
  ])
  expect(report.shapes[0]?.paragraphs[0]?.runs[1]).not.toHaveProperty('resolvedThemeColor')
  zip.file(
    'ppt/slideLayouts/slideLayout1.xml',
    `<p:sldLayout xmlns:p="urn:p" xmlns:a="urn:a"><p:clrMapOvr><a:overrideClrMapping accent1="accent1"/></p:clrMapOvr></p:sldLayout>`,
  )
  zip.file(
    'ppt/theme/theme1.xml',
    `<a:theme xmlns:a="urn:a"><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="ABCDEF"/></a:accent1><a:accent2><a:srgbClr val="123456"/></a:accent2></a:clrScheme><a:fontScheme><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont></a:fontScheme></a:themeElements></a:theme>`,
  )
  const overridden = await inspectPowerPointRichText(await zip.generateAsync({ type: 'base64' }))
  expect(overridden.shapes[0]?.paragraphs[0]?.runs[0]?.resolvedThemeColor).toBe('#ABCDEF')
  zip.file(
    'ppt/theme/theme1.xml',
    `<a:theme xmlns:a="urn:a"><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="ABCDEF"><a:tint val="50000"/></a:srgbClr></a:accent1></a:clrScheme></a:themeElements></a:theme>`,
  )
  const transformed = await inspectPowerPointRichText(await zip.generateAsync({ type: 'base64' }))
  expect(transformed.shapes[0]?.paragraphs[0]?.runs[0]).not.toHaveProperty('resolvedThemeColor')
  zip.remove('ppt/slideMasters/_rels/slideMaster1.xml.rels')
  const unlinked = await inspectPowerPointRichText(await zip.generateAsync({ type: 'base64' }))
  expect(unlinked.shapes[0]?.paragraphs[0]?.runs[0]).not.toHaveProperty('resolvedThemeColor')
  expect(unlinked.shapes[0]?.paragraphs[0]?.runs[0]).not.toHaveProperty('resolvedThemeTypeface')
})
it('reads formatted runs from a real PptxGenJS slide without claiming inherited style', async () => {
  const deck = new PptxGenJS()
  deck.addSlide().addText(
    [
      { text: 'First', options: { bold: true, color: '112233' } },
      { text: 'Second', options: { italic: true } },
    ],
    { x: 1, y: 1, w: 3, h: 0.5 },
  )
  const bytes = await deck.write({ outputType: 'nodebuffer' })
  const report = await inspectPowerPointRichText(
    Buffer.from(bytes as Uint8Array).toString('base64'),
  )
  expect(report.inheritanceResolved).toBe(false)
  expect(report.shapes[0]?.paragraphs[0]?.runs).toEqual([
    expect.objectContaining({
      text: 'First',
      directFont: expect.objectContaining({ bold: true, color: '#112233' }),
    }),
    expect.objectContaining({
      text: 'Second',
      directFont: expect.objectContaining({ italic: true }),
    }),
  ])
})
it('follows theme relationships in a complete generated PPTX package', async () => {
  const deck = new PptxGenJS()
  deck.addSlide().addText('Theme-linked', { x: 1, y: 1, w: 3, h: 1, color: '123456' })
  const bytes = await deck.write({ outputType: 'nodebuffer' })
  const zip = await JSZip.loadAsync(bytes as Uint8Array)
  const slidePath = 'ppt/slides/slide1.xml'
  const original = await zip.file(slidePath)!.async('string')
  expect(original).toContain('123456')
  zip.file(
    slidePath,
    original.replace(/<a:srgbClr val="123456"\s*\/>/, '<a:schemeClr val="accent1"/>'),
  )
  const report = await inspectPowerPointRichText(await zip.generateAsync({ type: 'base64' }))
  const run = report.shapes
    .flatMap((shape) => shape.paragraphs.flatMap((p) => p.runs))
    .find((entry) => entry.text === 'Theme-linked')
  expect(run?.directFont.themeColor).toBe('accent1')
  expect(run?.resolvedThemeColor).toMatch(/^#[0-9A-F]{6}$/)
})
it('captures arbitrary current/selected/deck scopes without a generation artifact or a host write', async () => {
  const f = fixture()
  const result = await f.read()
  expect(result.isError, result.output).not.toBe(true)
  const b = JSON.parse(result.output)
  expect(b.scope).toEqual({ kind: 'current', slideIds: ['s2'] })
  expect(b.pages[0].shapes[0].text).toBe('original')
  expect(b.coverage.notes).toBe('not_read')
  expect(b.coverage.sources).toBe('not_read')
  expect(b.qaPassed).toBe(false)
  expect(b.contentDigest).toMatch(/^[a-f0-9]{64}$/)
  expect(JSON.parse((await f.read('selected')).output).scope).toEqual({
    kind: 'selected',
    slideIds: ['s2'],
    shapeIds: ['shape'],
  })
  expect(JSON.parse((await f.read('deck')).output).scope.slideIds).toEqual(['s1', 's2'])
  expect(f.inspectPage).not.toHaveBeenCalled()
})
it('captures and rechecks a 120-object current page baseline', async () => {
  const f = fixture()
  const current = f.pages.get('s2')!
  current.shapes.push(
    ...Array.from({ length: 119 }, (_, index) => ({
      ...current.shapes[0]!,
      id: `extra-${index}`,
      name: `Extra ${index}`,
    })),
  )
  const read = await f.read()
  expect(read.isError, read.output).not.toBe(true)
  const baseline = JSON.parse(read.output)
  expect(baseline.pages[0].shapes).toHaveLength(120)
  const checked = await f.call('check_presentation_baseline', { baseline_id: baseline.baselineId })
  expect(checked.isError, checked.output).not.toBe(true)
})
it('reports manual edits and selection/order drift without silently replacing the baseline', async () => {
  const f = fixture(),
    b = JSON.parse((await f.read('deck')).output)
  expect(
    JSON.parse((await f.call('check_presentation_baseline', { baseline_id: b.baselineId })).output)
      .unchanged,
  ).toBe(true)
  f.pages.set('s1', page('s1', 'manual change'))
  f.setContext({ ...f.getContext(), slideIds: ['s2', 's1'], selectedShapeIds: [] })
  const r = JSON.parse(
    (await f.call('check_presentation_baseline', { baseline_id: b.baselineId })).output,
  )
  expect(r).toMatchObject({
    unchanged: false,
    changedSlideIds: ['s1'],
    orderChanged: true,
    selectionChanged: true,
  })
  expect(
    JSON.parse((await f.call('check_presentation_baseline', { baseline_id: b.baselineId })).output)
      .unchanged,
  ).toBe(false)
})
it('reads large decks in explicit bounded windows with an honest coverage marker', async () => {
  const f = fixture()
  const slideIds = Array.from({ length: 43 }, (_, i) => `s${i}`)
  for (const id of slideIds) f.pages.set(id, page(id))
  f.setContext({ ...f.getContext(), slideIds })
  const first = JSON.parse((await f.read('deck')).output)
  expect(first.scope.slideIds).toEqual(slideIds.slice(0, 20))
  expect(first.scope.deckWindow).toEqual({ start: 0, end: 20, total: 43, hasMore: true })
  const last = JSON.parse(
    (
      await f.call('read_presentation_baseline', {
        scope: 'deck',
        page_offset: 40,
        page_limit: 20,
      })
    ).output,
  )
  expect(last.scope.slideIds).toEqual(slideIds.slice(40))
  expect(last.scope.deckWindow).toEqual({ start: 40, end: 43, total: 43, hasMore: false })
  expect(last.pages).toHaveLength(3)
  expect(
    JSON.parse(
      (
        await f.call('check_presentation_baseline', {
          baseline_id: last.baselineId,
        })
      ).output,
    ).unchanged,
  ).toBe(true)
})
it('retains deck windows and verifies complete coverage without calling them atomic', async () => {
  const f = fixture()
  const slideIds = Array.from({ length: 43 }, (_, i) => `s${i}`)
  for (const id of slideIds) f.pages.set(id, page(id))
  f.setContext({ ...f.getContext(), slideIds })
  const windows: string[] = []
  for (const page_offset of [0, 20, 40]) {
    const result = await f.call('read_presentation_baseline', { scope: 'deck', page_offset })
    expect(result.isError, result.output).not.toBe(true)
    windows.push(JSON.parse(result.output).baselineId)
  }
  expect(f.skill.snapshot(windows[0]!)).toBeDefined()
  const complete = await f.call('check_presentation_baseline_windows', { baseline_ids: windows })
  expect(complete.isError, complete.output).not.toBe(true)
  expect(JSON.parse(complete.output)).toMatchObject({
    coveredPages: 43,
    totalPages: 43,
    complete: true,
    unchanged: true,
    atomicSnapshot: false,
    qaPassed: false,
  })
  const partial = JSON.parse(
    (
      await f.call('check_presentation_baseline_windows', {
        baseline_ids: windows.slice(1),
      })
    ).output,
  )
  expect(partial).toMatchObject({
    coveredPages: 23,
    totalPages: 43,
    complete: false,
    unchanged: true,
    atomicSnapshot: false,
  })
  f.pages.set('s1', page('s1', 'manual edit'))
  const changed = JSON.parse(
    (
      await f.call('check_presentation_baseline_windows', {
        baseline_ids: windows,
      })
    ).output,
  )
  expect(changed).toMatchObject({ complete: true, unchanged: false })
  expect(changed.changedSlideIds).toContain('s1')
})
it('rejects overlapping windows and windows captured with different global context', async () => {
  const f = fixture()
  const slideIds = Array.from({ length: 25 }, (_, i) => `s${i}`)
  for (const id of slideIds) f.pages.set(id, page(id))
  f.setContext({ ...f.getContext(), slideIds })
  const first = JSON.parse(
    (
      await f.call('read_presentation_baseline', {
        scope: 'deck',
        page_offset: 0,
      })
    ).output,
  )
  const overlap = JSON.parse(
    (
      await f.call('read_presentation_baseline', {
        scope: 'deck',
        page_offset: 10,
      })
    ).output,
  )
  expect(
    (
      await f.call('check_presentation_baseline_windows', {
        baseline_ids: [first.baselineId, overlap.baselineId],
      })
    ).output,
  ).toBe('presentation_baseline_windows_inconsistent')
  f.setContext({ ...f.getContext(), selectedSlideIds: ['s1'], selectedShapeIds: [] })
  const next = JSON.parse(
    (
      await f.call('read_presentation_baseline', {
        scope: 'deck',
        page_offset: 20,
      })
    ).output,
  )
  expect(
    (
      await f.call('check_presentation_baseline_windows', {
        baseline_ids: [first.baselineId, next.baselineId],
      })
    ).output,
  ).toBe('presentation_baseline_windows_inconsistent')
})
it('clears retained windows on Save As and session reset', async () => {
  const f = fixture()
  const first = JSON.parse((await f.read()).output).baselineId as string
  const second = JSON.parse((await f.read('deck')).output).baselineId as string
  expect(f.skill.snapshot(first)).toBeDefined()
  expect(f.skill.snapshot(second)).toBeDefined()
  f.setDocument('save-as')
  expect((await f.call('check_presentation_baseline', { baseline_id: first })).output).toBe(
    'presentation_document_changed',
  )
  expect(f.skill.snapshot(first)).toBeUndefined()
  expect(f.skill.snapshot(second)).toBeUndefined()
  f.skill.clear()
  expect(f.skill.snapshot(first)).toBeUndefined()
})
it('rejects empty current selection, invalid deck windows, forged IDs and unsupported input', async () => {
  const f = fixture()
  f.setContext({ ...f.getContext(), selectedSlideIds: [], selectedShapeIds: [] })
  expect((await f.read()).output).toBe('presentation_selection_empty')
  f.setContext({ ...f.getContext(), slideIds: Array.from({ length: 21 }, (_, i) => `s${i}`) })
  for (const id of f.getContext().slideIds) f.pages.set(id, page(id))
  for (const input of [
    { scope: 'deck', page_offset: 21 },
    { scope: 'deck', page_offset: -1 },
    { scope: 'deck', page_limit: 21 },
    { scope: 'deck', page_limit: 0 },
    { scope: 'current', page_offset: 1 },
  ])
    expect((await f.call('read_presentation_baseline', input)).isError).toBe(true)
  expect(
    (await f.call('read_presentation_baseline', { scope: 'current', extra: true })).output,
  ).toBe('invalid_tool_input')
  expect((await f.call('check_presentation_baseline', { baseline_id: 'forged' })).output).toBe(
    'presentation_baseline_missing',
  )
})
it('discards a torn read and a read across Save As', async () => {
  const f = fixture()
  f.adapter.readPage.mockImplementation(async (id) => {
    const value = structuredClone(f.pages.get(id)!)
    f.pages.set(id, page(id, value.shapes[0]!.text + '!'))
    return value
  })
  expect((await f.read()).output).toBe('presentation_baseline_changed')
  f.adapter.readPage.mockImplementation(async (id) => {
    f.setDocument('other')
    return page(id)
  })
  expect((await f.read()).output).toBe('presentation_document_changed')
})
it('clear and cancellation suppress late reads and invalidate saved baselines', async () => {
  const f = fixture(),
    b = JSON.parse((await f.read()).output)
  f.skill.clear()
  expect((await f.call('check_presentation_baseline', { baseline_id: b.baselineId })).isError).toBe(
    true,
  )
  let release!: (v: ReturnType<typeof f.getContext>) => void
  f.adapter.readContext.mockImplementationOnce(
    () =>
      new Promise((resolve) => {
        release = resolve
      }),
  )
  const pending = f.read()
  await vi.waitFor(() => expect(release).toBeDefined())
  f.skill.clear()
  release(f.getContext())
  expect((await pending).output).toBe('cancelled')
  const controller = new AbortController()
  controller.abort()
  expect((await f.call('read_presentation_baseline', {}, controller.signal)).output).toBe(
    'cancelled',
  )
})
it('captures a screenshot by exact baseline page ID and refuses stale or out-of-scope content', async () => {
  const f = fixture(),
    b = JSON.parse((await f.read()).output)
  const r = await f.call('read_presentation_baseline_page', {
    baseline_id: b.baselineId,
    slide_id: 's2',
  })
  expect(r.isError, r.output).not.toBe(true)
  expect(r.modelContent?.[0]?.type).toBe('image')
  expect(f.inspectPage).toHaveBeenCalledWith('s2', undefined)
  expect(
    (await f.call('read_presentation_baseline_page', { baseline_id: b.baselineId, slide_id: 's1' }))
      .isError,
  ).toBe(true)
  f.pages.set('s2', page('s2', 'manual'))
  expect(
    (await f.call('read_presentation_baseline_page', { baseline_id: b.baselineId, slide_id: 's2' }))
      .output,
  ).toBe('presentation_baseline_changed')
})
it.each(['png', 'geometry'])(
  'rejects an invalid or mismatched screenshot response (%s)',
  async (kind) => {
    const f = fixture(),
      b = JSON.parse((await f.read()).output)
    const original = await f.inspectPage('s2')
    if (kind === 'png') original.screenshot.base64 = 'iVBORw0KGgo_invalid'
    else original.shapes[0]!.left++
    f.inspectPage.mockResolvedValue(original)
    expect(
      (
        await f.call('read_presentation_baseline_page', {
          baseline_id: b.baselineId,
          slide_id: 's2',
        })
      ).isError,
    ).toBe(true)
  },
)
it('keeps the latest completed request and drops a superseded concurrent read', async () => {
  const f = fixture()
  let release!: (v: ReturnType<typeof f.getContext>) => void
  f.adapter.readContext.mockImplementationOnce(
    () =>
      new Promise((resolve) => {
        release = resolve
      }),
  )
  const pending = f.read()
  await vi.waitFor(() => expect(release).toBeDefined())
  const latest = JSON.parse((await f.read('deck')).output)
  release(f.getContext())
  expect((await pending).output).toBe('cancelled')
  expect(
    JSON.parse(
      (await f.call('check_presentation_baseline', { baseline_id: latest.baselineId })).output,
    ).unchanged,
  ).toBe(true)
})
it('bounds combined page output and refuses stale document IDs', async () => {
  const f = fixture()
  const b = JSON.parse((await f.read()).output)
  f.setDocument('save-as')
  expect((await f.call('check_presentation_baseline', { baseline_id: b.baselineId })).output).toBe(
    'presentation_document_changed',
  )
  f.pages.set('s2', page('s2', '中'.repeat(100_000)))
  expect((await f.read()).output).toBe('presentation_baseline_size_limit')
})
it('reports deleted baseline pages instead of retargeting by slide index', async () => {
  const f = fixture(),
    b = JSON.parse((await f.read()).output)
  f.setContext({
    ...f.getContext(),
    slideIds: ['s1'],
    selectedSlideIds: ['s1'],
    selectedShapeIds: [],
  })
  expect(
    JSON.parse((await f.call('check_presentation_baseline', { baseline_id: b.baselineId })).output),
  ).toMatchObject({ unchanged: false, changedSlideIds: ['s2'], orderChanged: true })
})
it('preserves unsupported theme status and detects a changed master snapshot', async () => {
  const f = fixture()
  let theme = 'black'
  const skill = createPresentationBaselineSkill({
    adapter: f.adapter,
    documentId: async () => 'doc',
    readMasters: async () => ({
      masters: [
        {
          id: 'm',
          name: 'master',
          background: { type: 'solid', color: theme },
          themeColors: { Accent1: theme },
          layouts: [],
        },
      ],
    }),
  })
  const b = JSON.parse(
    (await skill.executeTool({ id: 'read', name: 'read_presentation_baseline', input: {} })).output,
  )
  expect(b.coverage.theme).toBe('read')
  theme = 'white'
  const r = await skill.executeTool({
    id: 'check',
    name: 'check_presentation_baseline',
    input: { baseline_id: b.baselineId },
  })
  expect(JSON.parse(r.output)).toMatchObject({ unchanged: false, stylesChanged: true })
})
it('reads bounded native complex objects only from the exact scoped baseline slide', async () => {
  const f = fixture()
  f.setPackage(await complexPagePackage())
  const baseline = JSON.parse((await f.read()).output)
  const result = await f.call('read_presentation_baseline_complex_page', {
    baseline_id: baseline.baselineId,
    slide_id: 's2',
  })
  expect(result.isError, result.output).not.toBe(true)
  expect(JSON.parse(result.output)).toMatchObject({
    baselineId: baseline.baselineId,
    slideId: 's2',
    tables: [{ shapeId: '7', rows: [['North']] }],
    charts: [],
    cacheOnly: true,
    qaPassed: false,
  })
  expect(result.mutated).toBe(false)
  expect(f.exportPagePackage).toHaveBeenCalledTimes(2)
  expect(f.exportPagePackage).toHaveBeenCalledWith('s2', undefined)
  expect(
    (
      await f.call('read_presentation_baseline_complex_page', {
        baseline_id: baseline.baselineId,
        slide_id: 's1',
      })
    ).output,
  ).toBe('presentation_baseline_scope_mismatch')
})
it('rejects complex page reads when the package or host baseline changes during export', async () => {
  const f = fixture()
  const first = await complexPagePackage('North')
  const changed = await complexPagePackage('South')
  f.setPackage(first)
  const baseline = JSON.parse((await f.read()).output)
  f.exportPagePackage
    .mockImplementationOnce(async (slideId) => ({
      slideId,
      slideIds: [...f.getContext().slideIds],
      base64: first,
    }))
    .mockImplementationOnce(async (slideId) => ({
      slideId,
      slideIds: [...f.getContext().slideIds],
      base64: changed,
    }))
  expect(
    (
      await f.call('read_presentation_baseline_complex_page', {
        baseline_id: baseline.baselineId,
        slide_id: 's2',
      })
    ).output,
  ).toBe('presentation_baseline_changed')
  f.exportPagePackage.mockReset()
  f.exportPagePackage.mockImplementation(async (slideId) => {
    f.pages.set('s2', page('s2', 'manual change'))
    return { slideId, slideIds: [...f.getContext().slideIds], base64: first }
  })
  expect(
    (
      await f.call('read_presentation_baseline_complex_page', {
        baseline_id: baseline.baselineId,
        slide_id: 's2',
      })
    ).output,
  ).toBe('presentation_baseline_changed')
})
it('reads one exact scoped chart source without treating cache as verified data', async () => {
  const f = fixture()
  f.pages.get('s2')!.shapes[0]!.type = 'Chart'
  f.pages.get('s2')!.shapes[0]!.id = '7'
  f.setContext({ ...f.getContext(), selectedShapeIds: ['7'] })
  f.setPackage(await chartPagePackage())
  const baseline = JSON.parse((await f.read()).output)
  const result = await f.call('read_presentation_baseline_chart_source', {
    baseline_id: baseline.baselineId,
    slide_id: 's2',
    shape_id: '7',
  })
  expect(result.isError, result.output).not.toBe(true)
  expect(JSON.parse(result.output)).toMatchObject({
    baselineId: baseline.baselineId,
    slideId: 's2',
    shapeId: '7',
    sourceKind: 'cache_only',
    verification: 'not_verified',
    qaPassed: false,
    writeAuthorized: false,
  })
  expect(result.mutated).toBe(false)
  expect(f.exportPagePackage).toHaveBeenCalledTimes(2)
  expect(
    (
      await f.call('read_presentation_baseline_chart_source', {
        baseline_id: baseline.baselineId,
        slide_id: 's1',
        shape_id: '7',
      })
    ).output,
  ).toBe('presentation_baseline_scope_mismatch')
})
it('refuses chart source evidence if the chart package changes during the read', async () => {
  const f = fixture()
  f.pages.get('s2')!.shapes[0]!.type = 'Chart'
  f.pages.get('s2')!.shapes[0]!.id = '7'
  f.setContext({ ...f.getContext(), selectedShapeIds: ['7'] })
  const before = await chartPagePackage()
  const edited = await JSZip.loadAsync(before, { base64: true })
  edited.file(
    'ppt/charts/chart1.xml',
    (await edited.file('ppt/charts/chart1.xml')!.async('string')).replace('12', '13'),
  )
  const after = await edited.generateAsync({ type: 'base64' })
  f.setPackage(before)
  const baseline = JSON.parse((await f.read()).output)
  f.exportPagePackage
    .mockImplementationOnce(async (slideId) => ({
      slideId,
      slideIds: ['s1', 's2'],
      base64: before,
    }))
    .mockImplementationOnce(async (slideId) => ({ slideId, slideIds: ['s1', 's2'], base64: after }))
  expect(
    (
      await f.call('read_presentation_baseline_chart_source', {
        baseline_id: baseline.baselineId,
        slide_id: 's2',
        shape_id: '7',
      })
    ).output,
  ).toBe('presentation_baseline_changed')
})
it('reads bounded notes from an exact stable baseline page without treating them as a source', async () => {
  const f = fixture()
  f.setPackage(await notesPagePackage())
  const baseline = JSON.parse((await f.read()).output)
  const result = await f.call('read_presentation_baseline_notes', {
    baseline_id: baseline.baselineId,
    slide_id: 's2',
  })
  expect(result.isError, result.output).not.toBe(true)
  expect(JSON.parse(result.output)).toMatchObject({
    baselineId: baseline.baselineId,
    slideId: 's2',
    status: 'read',
    text: 'First & second\nSecond paragraph',
    sourceVerified: false,
    qaPassed: false,
    writeAuthorized: false,
  })
  expect(result.mutated).toBe(false)
  expect(f.exportPagePackage).toHaveBeenCalledTimes(2)
  expect(
    (
      await f.call('read_presentation_baseline_notes', {
        baseline_id: baseline.baselineId,
        slide_id: 's1',
      })
    ).output,
  ).toBe('presentation_baseline_scope_mismatch')
})
it('reads notes from a real compiled one-page PowerPoint package', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const packageBase64 = Buffer.from((await compilePresentationDeck(deck)).bytes).toString('base64')
  const f = fixture()
  f.setPackage(packageBase64)
  const baseline = JSON.parse((await f.read()).output)
  const result = await f.call('read_presentation_baseline_notes', {
    baseline_id: baseline.baselineId,
    slide_id: 's2',
  })
  expect(result.isError, result.output).not.toBe(true)
  expect(JSON.parse(result.output)).toMatchObject({ status: 'read', sourceVerified: false })
  expect(JSON.parse(result.output).text).toContain('演讲备注')
})
it('distinguishes an absent notes relationship from an empty body', async () => {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', '<p:sld><p:cSld/></p:sld>')
  expect(await inspectPowerPointPageNotes(await zip.generateAsync({ type: 'base64' }))).toEqual({
    status: 'not_present',
    text: '',
  })
})
it('rejects changed and unsafe notes packages instead of returning stale text', async () => {
  const f = fixture()
  const before = await notesPagePackage()
  const after = await notesPagePackage('Changed')
  f.setPackage(before)
  const baseline = JSON.parse((await f.read()).output)
  f.exportPagePackage
    .mockImplementationOnce(async (slideId) => ({
      slideId,
      slideIds: ['s1', 's2'],
      base64: before,
    }))
    .mockImplementationOnce(async (slideId) => ({ slideId, slideIds: ['s1', 's2'], base64: after }))
  expect(
    (
      await f.call('read_presentation_baseline_notes', {
        baseline_id: baseline.baselineId,
        slide_id: 's2',
      })
    ).output,
  ).toBe('presentation_baseline_changed')
  f.setPackage(await notesPagePackage('Text', '../../outside.xml'))
  expect(
    (
      await f.call('read_presentation_baseline_notes', {
        baseline_id: baseline.baselineId,
        slide_id: 's2',
      })
    ).output,
  ).toBe('office_api_unsupported')
})

it('reads a final-page baseline for a 512-page generic restore without enlarging the page window', async () => {
  const f = fixture()
  const ids = Array.from({ length: 512 }, (_, index) => `large-${index}`)
  f.setContext({
    ...f.getContext(),
    slideIds: ids,
    selectedSlideIds: [ids[511]!],
    selectedShapeIds: [],
  })
  f.pages.set(ids[511]!, page(ids[511]!))
  const result = await f.call('read_presentation_baseline', {
    scope: 'deck',
    page_offset: 511,
    page_limit: 1,
  })
  expect(result.isError, result.output).not.toBe(true)
  expect(JSON.parse(result.output).scope.slideIds).toEqual([ids[511]])
  expect(f.adapter.readPage).toHaveBeenCalledTimes(2)
  expect(f.adapter.readPage.mock.calls.every(([id]) => id === ids[511])).toBe(true)
})

it('checks all 512 pages with the final twenty-sixth baseline window', async () => {
  const f = fixture()
  const ids = Array.from({ length: 512 }, (_, index) => `large-${index}`)
  for (const id of ids) f.pages.set(id, page(id))
  f.setContext({
    ...f.getContext(),
    slideIds: ids,
    selectedSlideIds: [ids[0]!],
    selectedShapeIds: [],
  })
  const windows: string[] = []
  for (let page_offset = 0; page_offset < ids.length; page_offset += 20) {
    const result = await f.call('read_presentation_baseline', { scope: 'deck', page_offset })
    expect(result.isError, result.output).not.toBe(true)
    windows.push(JSON.parse(result.output).baselineId)
  }
  expect(windows).toHaveLength(26)
  const checked = await f.call('check_presentation_baseline_windows', { baseline_ids: windows })
  expect(checked.isError, checked.output).not.toBe(true)
  expect(JSON.parse(checked.output)).toMatchObject({
    totalPages: 512,
    coveredPages: 512,
    complete: true,
    unchanged: true,
    atomicSnapshot: false,
    writeAuthorized: false,
  })
  f.pages.set(ids[511]!, page(ids[511]!, 'last-page edit'))
  const drift = await f.call('check_presentation_baseline_windows', { baseline_ids: windows })
  expect(drift.isError, drift.output).not.toBe(true)
  expect(JSON.parse(drift.output)).toMatchObject({
    complete: true,
    unchanged: false,
    changedSlideIds: [ids[511]],
  })
})
