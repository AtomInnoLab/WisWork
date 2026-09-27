import { expect, it } from 'vitest'
import { createHash } from 'node:crypto'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { openPptx } from '@wiswork/pptx-engine'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { comparePresentationPageStructure } from '../src/skills/powerpoint/presentation-structure-comparison'
import { createPresentationQaSkill } from '../src/skills/powerpoint/presentation-qa'
import { presentationArtifactContent } from '../src/skills/powerpoint/presentation-page-delivery'
import { InMemoryVfs } from '../src/skills/shared/vfs'

it('detects changed native text in an exported host page package', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const { bytes } = await compilePresentationDeck(deck)
  const source = (await openPptx(bytes)).deck.slides[0]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  const hostPackage = await JSZip.loadAsync(bytes)
  const slide = await hostPackage.file('ppt/slides/slide1.xml')!.async('string')
  hostPackage.file('ppt/slides/slide1.xml', slide.replace('科研汇报', '替换标题'))
  const hostBase64 = await hostPackage.generateAsync({ type: 'base64' })
  const result = await comparePresentationPageStructure(
    Buffer.from(bytes).toString('base64'),
    0,
    {
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes,
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: '' },
    },
    hostBase64,
  )
  expect(result.content).toMatchObject({ status: 'warning', changed: ['title'] })
})

it('detects changed table cells and leaves chart data explicitly unchecked', async () => {
  for (const [pageIndex, original, replacement, nativeType] of [
    [5, '120', '999', 'Table'],
    [6, '120', '999', 'Chart'],
  ] as const) {
    const deck = benchmarkDeck()
    deck.slides = [deck.slides[pageIndex]!]
    const { bytes } = await compilePresentationDeck(deck)
    const source = (await openPptx(bytes)).deck.slides[0]!
    const shapes = source.elements.map((element, index) => ({
      id: String(index),
      name: element.name!,
      type: element.name === 'table' || element.name === 'chart' ? nativeType : 'TextBox',
      left: (element.transform.offset.x * 72) / 914400,
      top: (element.transform.offset.y * 72) / 914400,
      width: (element.transform.offset.cx * 72) / 914400,
      height: (element.transform.offset.cy * 72) / 914400,
    }))
    const hostPackage = await JSZip.loadAsync(bytes)
    const target = nativeType === 'Table' ? 'ppt/slides/slide1.xml' : 'ppt/charts/chart1.xml'
    const content = await hostPackage.file(target)!.async('string')
    const valueTag = nativeType === 'Table' ? 'a:t' : 'c:v'
    hostPackage.file(
      target,
      content.replace(
        `<${valueTag}>${original}</${valueTag}>`,
        `<${valueTag}>${replacement}</${valueTag}>`,
      ),
    )
    const result = await comparePresentationPageStructure(
      Buffer.from(bytes).toString('base64'),
      0,
      {
        slideId: 'host',
        slideWidth: 960,
        slideHeight: 540,
        shapes,
        shapesTruncated: false,
        overflows: [],
        overlaps: [],
        overlapsTruncated: false,
        screenshot: { mime: 'image/png', base64: '' },
      },
      await hostPackage.generateAsync({ type: 'base64' }),
    )
    expect(result.structureStatus).toBe('passed')
    if (nativeType === 'Table')
      expect(result).toMatchObject({
        status: 'warning',
        content: { status: 'warning', changed: ['table'] },
      })
    else
      expect(result).toMatchObject({
        status: 'warning',
        content: { status: 'warning', cacheChanged: ['chart'], unchecked: ['chart'] },
      })
  }
})

it('compares a selected chart cache from a multi-page source deck', async () => {
  const { bytes } = await compilePresentationDeck(benchmarkDeck())
  const source = (await openPptx(bytes)).deck.slides[6]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: element.name === 'chart' ? 'Chart' : 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  const one = benchmarkDeck()
  one.slides = [one.slides[6]!]
  const host = await compilePresentationDeck(one)
  const zip = await JSZip.loadAsync(host.bytes)
  const hostInspection = {
    slideId: 'host',
    slideWidth: 960,
    slideHeight: 540,
    shapes,
    shapesTruncated: false,
    overflows: [],
    overlaps: [],
    overlapsTruncated: false,
    screenshot: { mime: 'image/png' as const, base64: '' },
  }
  const unchanged = await comparePresentationPageStructure(
    Buffer.from(bytes).toString('base64'),
    6,
    hostInspection,
    Buffer.from(host.bytes).toString('base64'),
  )
  expect(unchanged.content).toMatchObject({
    status: 'incomplete',
    cacheChanged: [],
    unchecked: ['chart'],
  })
  const chartPath = Object.keys(zip.files).find((path) =>
    /^ppt\/charts\/chart\d+\.xml$/.test(path),
  )!
  const chart = await zip.file(chartPath)!.async('string')
  zip.file(chartPath, chart.replace('<c:v>120</c:v>', '<c:v>999</c:v>'))
  const result = await comparePresentationPageStructure(
    Buffer.from(bytes).toString('base64'),
    6,
    hostInspection,
    await zip.generateAsync({ type: 'base64' }),
  )
  expect(result.content).toMatchObject({
    status: 'warning',
    cacheChanged: ['chart'],
    unchecked: ['chart'],
  })
})

it('detects replaced embedded picture media in an exported page', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const { bytes } = await compilePresentationDeck(deck)
  const source = (await openPptx(bytes)).deck.slides[0]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: element.type === 'picture' ? 'Image' : 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  const zip = await JSZip.loadAsync(bytes)
  const mediaPath = Object.keys(zip.files).find((path) =>
    /^ppt\/media\/image[^/]+\.png$/.test(path),
  )!
  zip.file(
    mediaPath,
    Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII=',
      'base64',
    ),
  )
  const result = await comparePresentationPageStructure(
    Buffer.from(bytes).toString('base64'),
    0,
    {
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes,
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: '' },
    },
    await zip.generateAsync({ type: 'base64' }),
  )
  expect(result.content).toMatchObject({
    status: 'warning',
    mediaChanged: ['image'],
    unchecked: ['image'],
  })
})

it('compares a selected picture from a multi-page source deck', async () => {
  const full = await compilePresentationDeck(benchmarkDeck())
  const source = (await openPptx(full.bytes)).deck.slides[2]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: element.type === 'picture' ? 'Image' : 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  const one = benchmarkDeck()
  one.slides = [one.slides[2]!]
  const host = await compilePresentationDeck(one)
  const result = await comparePresentationPageStructure(
    Buffer.from(full.bytes).toString('base64'),
    2,
    {
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes,
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: '' },
    },
    Buffer.from(host.bytes).toString('base64'),
  )
  expect(result.content).toMatchObject({
    status: 'incomplete',
    mediaChanged: [],
    unchecked: ['image'],
  })
})

it('matches text, shapes, images, tables and charts from compiled pages to Office shape readback', async () => {
  const { bytes } = await compilePresentationDeck(benchmarkDeck())
  const opened = await openPptx(bytes)
  for (const [pageIndex, source] of opened.deck.slides.entries()) {
    const shapes = source.elements.map((element, index) => ({
      id: String(index + 1),
      name: element.name!,
      type:
        element.type === 'table'
          ? 'Table'
          : element.type === 'chart'
            ? 'Chart'
            : element.type === 'picture'
              ? 'Image'
              : 'TextBox',
      left: (element.transform.offset.x * 72) / 914400,
      top: (element.transform.offset.y * 72) / 914400,
      width: (element.transform.offset.cx * 72) / 914400,
      height: (element.transform.offset.cy * 72) / 914400,
    }))
    const result = await comparePresentationPageStructure(
      Buffer.from(bytes).toString('base64'),
      pageIndex,
      {
        slideId: 'host',
        slideWidth: 960,
        slideHeight: 540,
        shapes,
        shapesTruncated: false,
        overflows: [],
        overlaps: [],
        overlapsTruncated: false,
        screenshot: { mime: 'image/png', base64: '' },
      },
    )
    expect(result).toMatchObject({
      structureStatus: 'passed',
      sourceCount: 3,
      hostCount: 3,
      issues: [],
    })
  }
})

it('reports missing, retyped and moved editable objects without a false pass', async () => {
  const { bytes } = await compilePresentationDeck(benchmarkDeck())
  const opened = await openPptx(bytes)
  const source = opened.deck.slides[6]!
  const chart = source.elements.find((element) => element.type === 'chart')!
  const result = await comparePresentationPageStructure(Buffer.from(bytes).toString('base64'), 6, {
    slideId: 'host',
    slideWidth: 960,
    slideHeight: 540,
    shapes: [
      { id: '1', name: chart.name!, type: 'Image', left: 100, top: 100, width: 100, height: 100 },
    ],
    shapesTruncated: false,
    overflows: [],
    overlaps: [],
    overlapsTruncated: false,
    screenshot: { mime: 'image/png', base64: '' },
  })
  expect(result.structureStatus).toBe('warning')
  expect(result.issues).toEqual(
    expect.arrayContaining([
      { name: 'chart', kind: 'type_changed', sourceType: 'chart', hostType: 'Image' },
      { name: 'title', kind: 'missing' },
    ]),
  )
})

it('marks a truncated host readback incomplete even when known objects match', async () => {
  const { bytes } = await compilePresentationDeck(benchmarkDeck())
  const source = (await openPptx(bytes)).deck.slides[0]!
  const result = await comparePresentationPageStructure(Buffer.from(bytes).toString('base64'), 0, {
    slideId: 'host',
    slideWidth: 960,
    slideHeight: 540,
    shapes: source.elements.map((element, index) => ({
      id: String(index),
      name: element.name!,
      type: 'TextBox',
      left: (element.transform.offset.x * 72) / 914400,
      top: (element.transform.offset.y * 72) / 914400,
      width: (element.transform.offset.cx * 72) / 914400,
      height: (element.transform.offset.cy * 72) / 914400,
    })),
    shapesTruncated: true,
    overflows: [],
    overlaps: [],
    overlapsTruncated: false,
    screenshot: { mime: 'image/png', base64: '' },
  })
  expect(result.status).toBe('incomplete')
})

it('flags a native object moved beyond the host conversion tolerance', async () => {
  const { bytes } = await compilePresentationDeck(benchmarkDeck())
  const source = (await openPptx(bytes)).deck.slides[0]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  shapes[0]!.left += 5
  const result = await comparePresentationPageStructure(Buffer.from(bytes).toString('base64'), 0, {
    slideId: 'host',
    slideWidth: 960,
    slideHeight: 540,
    shapes,
    shapesTruncated: false,
    overflows: [],
    overlaps: [],
    overlapsTruncated: false,
    screenshot: { mime: 'image/png', base64: '' },
  })
  expect(result).toMatchObject({
    structureStatus: 'warning',
    issues: [{ name: 'title', kind: 'geometry_changed' }],
  })
})

it('does not combine inconsistent Office shape and package readbacks into a pass', async () => {
  const { bytes } = await compilePresentationDeck(benchmarkDeck())
  const source = (await openPptx(bytes)).deck.slides[0]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  shapes[0]!.left += 5
  const base64 = Buffer.from(bytes).toString('base64')
  const result = await comparePresentationPageStructure(
    base64,
    0,
    {
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes,
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: '' },
    },
    base64,
  )
  expect(result).toMatchObject({ status: 'incomplete', readbackConsistent: false })
  expect(result.content.status).toBe('incomplete')
})

it('exposes the comparison for the exact imported production page without changing QA records', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[5]!]
  const { bytes } = await compilePresentationDeck(deck)
  const base64 = Buffer.from(bytes).toString('base64')
  const source = (await openPptx(bytes)).deck.slides[0]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index + 1),
    name: element.name!,
    type: element.type === 'table' ? 'Table' : 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  const artifact = {
    documentId: 'doc',
    projectId: 'project',
    requestId: 'request',
    pptxBase64: '',
    pagePptxBase64: [base64],
    planRevision: 1,
    slideCount: 1,
    pages: [{ id: 'page', title: 'Table', sourceSlideId: '256#' }],
  }
  const receipt = {
    state: 'complete' as const,
    documentId: 'doc',
    slideIds: ['host'],
    checkpoint: {
      version: 2 as const,
      artifactDigest: createHash('sha256')
        .update(presentationArtifactContent(artifact))
        .digest('hex'),
      pageIds: ['page'],
      sourceSlideIds: ['256#'],
      baselineSlideIds: [],
      completed: [{ sourceSlideId: '256#', slideId: 'host' }],
    },
  }
  const skill = createPresentationQaSkill({
    available: () => true,
    artifact: () => artifact,
    documentId: async () => 'doc',
    readReceipt: () => receipt,
    inspectPage: async () => ({
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes,
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: '' },
    }),
    exportPage: async (slideId) => ({ slideId, base64 }),
    readQa: () => undefined,
    writeQa: async () => {
      throw new Error('unexpected write')
    },
    vfs: new InMemoryVfs(),
  })
  const result = await skill.executeTool({
    id: 'compare',
    name: 'compare_presentation_page_structure',
    input: { page_id: 'page' },
  })
  expect(result.isError).not.toBe(true)
  expect(JSON.parse(result.output)).toMatchObject({
    status: 'passed',
    pageId: 'page',
    hostSlideId: 'host',
    readbackConsistent: true,
    content: { status: 'passed', changed: [], unchecked: [] },
  })
})
