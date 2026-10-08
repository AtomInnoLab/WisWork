import { createHash } from 'node:crypto'
import JSZip from 'jszip'
import { expect, it, vi } from 'vitest'
import { openPptx } from '@wiswork/pptx-engine'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark.js'
import { createPresentationQaSkill } from '../src/skills/powerpoint/presentation-qa.js'
import { presentationArtifactContent } from '../src/skills/powerpoint/presentation-page-delivery.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'

it.each(['bar', 'line', 'pie'] as const)(
  'actual QA tool detects %s chart palette drift without changing saved visual QA',
  async (type) => {
    const deck = benchmarkDeck()
    deck.slides = [deck.slides[6]!]
    const chart = deck.slides[0]!.elements.find((e) => e.kind === 'chart')!
    if (chart.kind !== 'chart') throw Error('invalid fixture')
    chart.chartType = type
    deck.style.background = 'F2E7D5'
    deck.style.textColor = '314259'
    deck.style.accentColor = 'A05273'
    const { bytes } = await compilePresentationDeck(deck)
    const base64 = Buffer.from(bytes).toString('base64')
    const source = (await openPptx(bytes)).deck.slides[0]!
    const shapes = source.elements.map((element, index) => ({
      id: String(index + 1),
      name: element.name!,
      type: element.type === 'chart' ? 'Chart' : 'TextBox',
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
      pages: [{ id: 'page', title: 'Chart', sourceSlideId: '256#' }],
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
    let exported = base64
    const writeQa = vi.fn(async () => {
      throw Error('unexpected QA write')
    })
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
        screenshot: { mime: 'image/png' as const, base64: '' },
      }),
      exportPage: async (slideId) => ({ slideId, base64: exported }),
      readQa: () => undefined,
      writeQa,
      vfs: new InMemoryVfs(),
    })
    const call = {
      id: 'compare',
      name: 'compare_presentation_page_structure',
      input: { page_id: 'page' },
    }
    const before = await skill.executeTool(call)
    expect(before.isError).not.toBe(true)
    expect(JSON.parse(before.output).content).toMatchObject({
      chartStyleChanged: [],
      cacheChanged: [],
      unchecked: ['chart'],
    })
    const host = await JSZip.loadAsync(bytes)
    const path = Object.keys(host.files).find((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))!
    const original = await host.file(path)!.async('string')
    const changed = original.replace(/(<a:srgbClr val=")A05273(?=")/, '$1EE1122')
    expect(changed).not.toBe(original)
    host.file(path, changed)
    exported = await host.generateAsync({ type: 'base64' })
    const after = await skill.executeTool(call)
    expect(after.isError).not.toBe(true)
    expect(JSON.parse(after.output)).toMatchObject({
      status: 'warning',
      pageId: 'page',
      hostSlideId: 'host',
      structureStatus: 'passed',
      readbackConsistent: true,
      content: {
        status: 'warning',
        chartStyleChanged: ['chart'],
        cacheChanged: [],
        unchecked: ['chart'],
      },
    })
    if (type === 'bar') {
      const pointOverride = original.replace(
        '</c:ser>',
        '<c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="EE1122"/></a:solidFill></c:spPr></c:dPt></c:ser>',
      )
      expect(pointOverride).not.toBe(original)
      host.file(path, pointOverride)
      exported = await host.generateAsync({ type: 'base64' })
      const pointResult = await skill.executeTool(call)
      expect(pointResult.isError).not.toBe(true)
      expect(JSON.parse(pointResult.output).content).toMatchObject({
        chartStyleChanged: ['chart'],
        cacheChanged: [],
        unchecked: ['chart'],
      })
    }
    expect(writeQa).not.toHaveBeenCalled()
  },
)
