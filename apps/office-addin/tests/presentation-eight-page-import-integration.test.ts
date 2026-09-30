import { expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { createStructuredProposalController } from '../src/agent/proposal-controller'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import { createPresentationProductionDeliverySkill } from '../src/skills/powerpoint/presentation-page-delivery'
import type { CompiledPresentationArtifact } from '../src/skills/powerpoint/presentation-delivery'

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
      if (host.length === 5 && !interrupted) {
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
  expect(host).toHaveLength(5)
  expect(
    binding.readReceipt('production/eight-page-project/eight-page-run')?.checkpoint?.completed,
  ).toHaveLength(4)
  binding = createPresentationDocumentBinding(settings)
  await confirm()
  expect(host).toHaveLength(9)
  expect(adapter.insertPage).toHaveBeenCalledTimes(9)
  expect([...written.values()]).toEqual(pagePackages)
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
