import { expect, it, vi } from 'vitest'
import { createHash } from 'node:crypto'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { openPptx } from '@wiswork/pptx-engine'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { createPresentationQaSkill } from '../src/skills/powerpoint/presentation-qa'
import { presentationArtifactContent } from '../src/skills/powerpoint/presentation-page-delivery'
import { InMemoryVfs } from '../src/skills/shared/vfs'

it.each(['column', 'row', 'merge', 'transparency'] as const)(
  'does not certify an imported table with changed %s structure when text and host geometry match',
  async (change) => {
    const deck = benchmarkDeck()
    deck.slides = [deck.slides[5]!]
    const { bytes } = await compilePresentationDeck(deck)
    const base64 = Buffer.from(bytes).toString('base64')
    const parsed = (await openPptx(bytes)).deck.slides[0]!
    const shapes = parsed.elements.map((element, index) => ({
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
    let hostBase64 = base64
    const writeQa = vi.fn(async () => {})
    const inspectPage = vi.fn(async () => ({
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes,
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png' as const, base64: '' },
    }))
    const exportPage = vi.fn(async (slideId: string) => ({ slideId, base64: hostBase64 }))
    const skill = createPresentationQaSkill({
      available: () => true,
      artifact: () => artifact,
      documentId: async () => 'doc',
      readReceipt: () => receipt,
      inspectPage,
      exportPage,
      readQa: () => undefined,
      writeQa,
      vfs: new InMemoryVfs(),
    })
    const compare = async (unsupported = false) => {
      const result = await skill.executeTool({
        id: 'compare',
        name: 'compare_presentation_page_structure',
        input: { page_id: 'page' },
      })
      if (unsupported) {
        expect(result).toMatchObject({
          isError: true,
          mutated: false,
          output: 'presentation_qa_structure_unavailable',
        })
        return undefined
      }
      expect(result.isError).not.toBe(true)
      expect(result.mutated).toBe(false)
      return JSON.parse(result.output)
    }
    expect(await compare()).toMatchObject({
      status: 'passed',
      content: { status: 'passed', changed: [] },
    })
    const zip = await JSZip.loadAsync(bytes)
    const path = 'ppt/slides/slide1.xml'
    const original = await zip.file(path)!.async('string')
    let index = 0
    const altered =
      change === 'transparency'
        ? original.replace(
            /(<a:lnL[^]*?<a:solidFill><a:srgbClr val=")([^"]+)("\/>)/,
            '$1$2"><a:alpha val="0"/></a:srgbClr>',
          )
        : change === 'merge'
          ? original.replace(/<a:tc>/g, (cell) =>
              index++ === 0 ? '<a:tc gridSpan="2">' : index === 2 ? '<a:tc hMerge="1">' : cell,
            )
          : original.replace(
              change === 'column' ? /(<a:gridCol\b[^>]*\bw=")(\d+)/g : /(<a:tr\b[^>]*\bh=")(\d+)/g,
              (match, prefix: string, size: string) =>
                index++ < 2 ? `${prefix}${Number(size) + (index === 1 ? 91440 : -91440)}` : match,
            )
    expect(altered).not.toBe(original)
    zip.file(path, altered)
    hostBase64 = await zip.generateAsync({ type: 'base64' })
    if (change === 'transparency') {
      await compare(true)
    } else {
      const changed = await compare()
      expect(changed).toMatchObject({
        status: 'warning',
        readbackConsistent: true,
        pageId: 'page',
        hostSlideId: 'host',
      })
      expect(changed.content).toMatchObject({ status: 'warning', tableStructureChanged: ['table'] })
      expect(changed.content.changed).toEqual([])
      expect(changed.issues).toEqual([])
    }
    expect(writeQa).not.toHaveBeenCalled()
    expect(inspectPage).toHaveBeenCalledTimes(2)
    expect(exportPage).toHaveBeenCalledTimes(2)
    expect(artifact.pagePptxBase64).toEqual([base64])
  },
)
