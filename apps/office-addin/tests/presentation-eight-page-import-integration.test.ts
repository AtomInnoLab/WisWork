import { expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { openPptx } from '@wiswork/pptx-engine'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { createStructuredProposalController } from '../src/agent/proposal-controller'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import { createPresentationProductionDeliverySkill } from '../src/skills/powerpoint/presentation-page-delivery'
import type { CompiledPresentationArtifact } from '../src/skills/powerpoint/presentation-delivery'
import { comparePresentationPageStructure } from '../src/skills/powerpoint/presentation-structure-comparison'

it('imports eight mixed native pages and resumes after a known pre-write interruption', async () => {
  const deck = benchmarkDeck()
  const compiled = await Promise.all(
    deck.slides.map((slide) => compilePresentationDeck({ ...deck, slides: [slide] })),
  )
  expect(compiled.map((page) => page.sourceSlideIds)).toEqual(deck.slides.map(() => ['256#']))
  const pagePackages = compiled.map((page) => Buffer.from(page.bytes).toString('base64'))
  const slideXml = await Promise.all(
    compiled.map(async (page) => {
      const zip = await JSZip.loadAsync(page.bytes)
      return { xml: await zip.file('ppt/slides/slide1.xml')!.async('string'), zip }
    }),
  )
  expect(slideXml[2]!.xml).toContain('<p:pic>')
  expect(Object.keys(slideXml[2]!.zip.files).some((path) => path.startsWith('ppt/media/'))).toBe(
    true,
  )
  expect(slideXml[3]!.xml).toContain('prst="roundRect"')
  expect(slideXml[4]!.xml).toContain('prst="rect"')
  expect(slideXml[5]!.xml).toContain('<a:tbl>')
  expect(slideXml[6]!.xml).toContain('graphicFrame')
  expect(Object.keys(slideXml[6]!.zip.files).some((path) => path.startsWith('ppt/charts/'))).toBe(
    true,
  )

  const values = new Map<string, unknown>()
  const settings = {
    get: (key: string) => values.get(key),
    set: (key: string, value: unknown) => {
      values.set(key, value)
    },
    save: async () => {},
    location: () => '',
  }
  let binding = createPresentationDocumentBinding(settings)
  const documentId = await binding.documentId()
  const artifact: CompiledPresentationArtifact = {
    documentId,
    projectId: 'eight-page-project',
    requestId: 'eight-page-run',
    planRevision: 1,
    pptxBase64: '',
    pagePptxBase64: pagePackages,
    slideCount: deck.slides.length,
    pages: deck.slides.map((slide, index) => ({
      id: slide.id,
      title: slide.title,
      sourceSlideId: compiled[index]!.sourceSlideIds![0]!,
    })),
  }
  const host = ['original']
  const written = new Map<string, string>()
  let interrupted = false
  const adapter = {
    available: () => true,
    snapshot: async () => ({ slideIds: [...host], fingerprint: JSON.stringify(host) }),
    insert: vi.fn(),
    insertPage: vi.fn(async (bytes: string) => {
      if (host.length === 4 && !interrupted) {
        interrupted = true
        throw new Error('cancelled')
      }
      const id = `host-${host.length}`
      host.push(id)
      written.set(id, bytes)
      return { slideIds: [id] }
    }),
    verify: async () => true,
    exportPage: async (id: string) => written.get(id)!,
  }
  const proposals = createStructuredProposalController()
  const options = {
    adapter,
    proposals,
    available: () => true,
    artifact: () => artifact,
    documentId: () => binding.documentId(),
    readReceipt: (key: string) => binding.readReceipt(key),
    writeReceipt: (key: string, value: Parameters<typeof binding.writeReceipt>[1]) =>
      binding.writeReceipt(key, value),
  }
  const confirm = async () => {
    const result = await createPresentationProductionDeliverySkill(options).executeTool({
      id: 'import',
      name: 'import_presentation_production',
      input: { project_id: artifact.projectId },
    })
    expect(result.isError).not.toBe(true)
    await proposals.confirm(proposals.pending()!.id)
  }
  await expect(confirm()).rejects.toThrow('cancelled')
  expect(host).toHaveLength(4)
  expect(
    binding.readReceipt('production/eight-page-project/eight-page-run')?.checkpoint?.completed,
  ).toHaveLength(3)
  binding = createPresentationDocumentBinding(settings)
  await confirm()
  expect(host).toHaveLength(9)
  expect(adapter.insertPage).toHaveBeenCalledTimes(9)
  expect([...written.values()]).toEqual(pagePackages)
  for (const [index, slideId] of host.slice(1).entries()) {
    const exported = written.get(slideId)!
    const native = (await openPptx(Buffer.from(exported, 'base64'))).deck.slides[0]!
    const shapes = native.elements.map((element, shapeIndex) => {
      const type = (
        {
          shape: 'GeometricShape',
          picture: 'Image',
          table: 'Table',
          chart: 'Chart',
        } as Record<string, string>
      )[element.type]
      if (!type || !element.name) throw new Error('unexpected_native_object')
      return {
        id: `host-shape-${shapeIndex}`,
        name: element.name,
        type,
        left: (element.transform.offset.x * 72) / 914400,
        top: (element.transform.offset.y * 72) / 914400,
        width: (element.transform.offset.cx * 72) / 914400,
        height: (element.transform.offset.cy * 72) / 914400,
      }
    })
    const inspection = {
      slideId,
      slideWidth: 960,
      slideHeight: 540,
      shapes,
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png' as const, base64: '' },
    }
    const comparison = await comparePresentationPageStructure(
      pagePackages[index]!,
      0,
      inspection,
      exported,
    )
    expect(comparison.structureStatus, `page ${index + 1} native structure`).toBe('passed')
    expect(comparison.readbackConsistent).toBe(true)
    expect(comparison.content.changed).toEqual([])
    expect(comparison.content.tableStructureChanged).toEqual([])
    expect(comparison.content.chartSourceChanged).toEqual([])
    expect(comparison.content.mediaChanged).toEqual([])
    if (index === 6) {
      const altered = await JSZip.loadAsync(Buffer.from(exported, 'base64'))
      const chartPath = Object.keys(altered.files).find((path) =>
        /^ppt\/charts\/chart\d+\.xml$/.test(path),
      )!
      const chartXml = await altered.file(chartPath)!.async('string')
      expect(chartXml).toContain('<c:v>120</c:v>')
      altered.file(chartPath, chartXml.replace('<c:v>120</c:v>', '<c:v>121</c:v>'))
      const changed = await comparePresentationPageStructure(
        pagePackages[index]!,
        0,
        inspection,
        await altered.generateAsync({ type: 'base64' }),
      )
      expect(changed.content.cacheChanged).toContain('chart')
      expect(changed.content.status).toBe('warning')
    }
  }
  expect(binding.readReceipt('production/eight-page-project/eight-page-run')).toMatchObject({
    state: 'complete',
    checkpoint: { pageIds: deck.slides.map((slide) => slide.id) },
  })
  binding = createPresentationDocumentBinding(settings)
  const repeated = await createPresentationProductionDeliverySkill(options).executeTool({
    id: 'repeat',
    name: 'import_presentation_production',
    input: { project_id: artifact.projectId },
  })
  expect(repeated.output).toContain('already_imported')
  expect(adapter.insertPage).toHaveBeenCalledTimes(9)
})

it('rejects a delayed page receipt after Save As without writing into the new document', async () => {
  const deck = benchmarkDeck()
  const compiled = await compilePresentationDeck({ ...deck, slides: [deck.slides[0]!] })
  const values = new Map<string, unknown>()
  let location = 'file://legal-original.pptx'
  const binding = createPresentationDocumentBinding({
    get: (key) => values.get(key),
    set: (key, value) => values.set(key, value),
    save: async () => {},
    location: () => location,
  })
  const artifact: CompiledPresentationArtifact = {
    documentId: await binding.documentId(),
    projectId: 'p0-17-legal',
    requestId: 'p0-17-legal-run',
    planRevision: 1,
    pptxBase64: '',
    pagePptxBase64: [Buffer.from(compiled.bytes).toString('base64')],
    slideCount: 1,
    pages: [{ id: 'p01', title: deck.slides[0]!.title, sourceSlideId: '256#' }],
  }
  const oldHost = ['old-baseline']
  const newHost = ['new-copy-baseline']
  let release!: () => void
  let started!: () => void
  const gate = new Promise<void>((resolve) => {
    release = resolve
  })
  const inserting = new Promise<void>((resolve) => {
    started = resolve
  })
  const adapter = {
    available: () => true,
    insert: vi.fn(),
    snapshot: async () => {
      const ids = location.includes('original') ? oldHost : newHost
      return { slideIds: [...ids], fingerprint: JSON.stringify(ids) }
    },
    insertPage: vi.fn(async () => {
      started()
      await gate
      oldHost.push('old-host-page')
      return { slideIds: ['old-host-page'] }
    }),
    verify: async () => true,
    exportPage: async () => artifact.pagePptxBase64![0]!,
  }
  const proposals = createStructuredProposalController()
  const skill = createPresentationProductionDeliverySkill({
    adapter,
    proposals,
    available: () => true,
    artifact: () => artifact,
    documentId: () => binding.documentId(),
    readReceipt: (key) => binding.readReceipt(key),
    writeReceipt: (key, value) => binding.writeReceipt(key, value),
  })
  const call = {
    id: 'import',
    name: 'import_presentation_production',
    input: { project_id: artifact.projectId },
  }
  expect((await skill.executeTool(call)).isError).not.toBe(true)
  const confirmation = proposals.confirm(proposals.pending()!.id)
  await inserting
  location = 'file://legal-saved-copy.pptx'
  release()
  await expect(confirmation).rejects.toThrow('presentation_document_changed')
  expect(newHost).toEqual(['new-copy-baseline'])
  expect(adapter.insertPage).toHaveBeenCalledTimes(1)
  expect(binding.readReceipt('production/p0-17-legal/p0-17-legal-run')).toMatchObject({
    state: 'pending',
    checkpoint: { completed: [], inFlight: { sourceSlideId: '256#' } },
  })
  expect((await skill.executeTool(call)).output).toBe('presentation_document_changed')
})
