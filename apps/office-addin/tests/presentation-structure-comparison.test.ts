import { expect, it } from 'vitest'
import { createHash } from 'node:crypto'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { openPptx } from '@wiswork/pptx-engine'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { comparePresentationPageStructure } from '../src/skills/powerpoint/presentation-structure-comparison'
import { createPresentationQaSkill } from '../src/skills/powerpoint/presentation-qa'
import { presentationArtifactContent } from '../src/skills/powerpoint/presentation-page-delivery'
import { InMemoryVfs } from '../src/skills/shared/vfs'

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
    expect(result).toMatchObject({ status: 'passed', sourceCount: 3, hostCount: 3, issues: [] })
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
  expect(result.status).toBe('warning')
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
    status: 'warning',
    issues: [{ name: 'title', kind: 'geometry_changed' }],
  })
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
  })
})
