import { expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { createStructuredProposalController } from '../src/agent/proposal-controller'
import {
  createPresentationAgentRunCheckpoint,
  createPresentationDocumentBinding,
} from '../src/skills/powerpoint/presentation-document'
import {
  createPresentationDeliverySkill,
  type CompiledPresentationArtifact,
  type PresentationImportRecord,
} from '../src/skills/powerpoint/presentation-delivery'
import {
  createPresentationProductionDeliverySkill,
  presentationImportKey,
  presentationArtifactContent,
  presentationPageMapping,
  validPresentationImportRecord,
} from '../src/skills/powerpoint/presentation-page-delivery'
const fixtureDeck = benchmarkDeck()
fixtureDeck.slides = [fixtureDeck.slides[0]!]
const fixturePage = Buffer.from((await compilePresentationDeck(fixtureDeck)).bytes).toString(
  'base64',
)
function fixture() {
  const artifact: CompiledPresentationArtifact = {
    documentId: 'doc',
    projectId: 'project',
    requestId: 'request',
    pptxBase64: '',
    planRevision: 1,
    pagePptxBase64: [fixturePage, fixturePage, fixturePage],
    slideCount: 3,
    pages: [0, 1, 2].map((i) => ({ id: `page${i}`, title: `Page ${i}`, sourceSlideId: '256#' })),
  }
  const host = ['old'],
    packages = new Map<string, string>(),
    receipts = new Map<string, PresentationImportRecord>(),
    proposals = createStructuredProposalController()
  const adapter = {
    available: () => true,
    snapshot: vi.fn(async () => ({ slideIds: [...host], fingerprint: JSON.stringify(host) })),
    insert: vi.fn(),
    insertPage: vi.fn(async (bytes: string, _source: string) => {
      const id = `host${host.length}`
      host.push(id)
      packages.set(id, bytes)
      return { slideIds: [id] }
    }),
    verify: vi.fn(async () => true),
    exportPage: vi.fn(async (id: string) => packages.get(id)!),
  }
  const options = {
    adapter,
    proposals,
    available: () => true,
    artifact: () => artifact,
    documentId: async () => 'doc',
    readReceipt: (key: string) => receipts.get(key),
    writeReceipt: async (key: string, record: PresentationImportRecord | undefined) => {
      if (record) receipts.set(key, structuredClone(record))
      else receipts.delete(key)
    },
  }
  const skill = createPresentationProductionDeliverySkill(options),
    call = {
      id: 'import',
      name: 'import_presentation_production',
      input: { project_id: 'project' },
    }
  const confirm = async () => {
    const result = await skill.executeTool(call)
    expect(result.isError).not.toBe(true)
    await proposals.confirm(proposals.pending()!.id)
  }
  return { artifact, host, receipts, proposals, adapter, options, skill, call, confirm }
}
it('checks each imported production page package before recording completion', async () => {
  const f = fixture()
  const source = fixturePage
  const otherZip = await JSZip.loadAsync(Buffer.from(source, 'base64'))
  const xml = await otherZip.file('ppt/slides/slide1.xml')!.async('string')
  otherZip.file('ppt/slides/slide1.xml', xml.replace('科研汇报', '内容不同'))
  const other = await otherZip.generateAsync({ type: 'base64' })
  f.artifact.pagePptxBase64 = [source, source, source]
  const skill = createPresentationProductionDeliverySkill({
    ...f.options,
    adapter: { ...f.adapter, exportPage: vi.fn(async () => other) },
  })
  expect((await skill.executeTool(f.call)).isError).not.toBe(true)
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow(
    'office_state_uncertain',
  )
  expect(f.receipts.get('production/project/request')?.checkpoint).toMatchObject({
    completed: [],
    inFlight: { sourceSlideId: '256#' },
  })
  expect(f.adapter.insertPage).toHaveBeenCalledOnce()
  expect((await skill.executeTool(f.call)).output).toBe('presentation_import_uncertain')
})

it('refuses production import before writing when the host lacks page export', async () => {
  const f = fixture()
  const skill = createPresentationProductionDeliverySkill({
    ...f.options,
    adapter: { ...f.adapter, supportsPageExport: () => false },
  })
  expect(skill.tools.map((tool) => tool.name)).toEqual([
    'read_presentation_production_import_status',
  ])
  expect(await skill.executeTool(f.call)).toMatchObject({
    isError: true,
    output: 'presentation_unavailable',
  })
  expect(f.adapter.insertPage).not.toHaveBeenCalled()
  const status = await skill.executeTool({
    id: 'status',
    name: 'read_presentation_production_import_status',
    input: {},
  })
  expect(status.isError).not.toBe(true)
  expect(JSON.parse(status.output).status).toBe('not_started')
})

it('requires an export adapter before proposing production import', async () => {
  const f = fixture()
  const skill = createPresentationProductionDeliverySkill({
    ...f.options,
    adapter: { ...f.adapter, exportPage: undefined },
  })
  expect(skill.tools.map((tool) => tool.name)).toEqual([
    'read_presentation_production_import_status',
  ])
  expect(await skill.executeTool(f.call)).toMatchObject({
    isError: true,
    output: 'presentation_unavailable',
  })
  expect(f.adapter.insertPage).not.toHaveBeenCalled()
})

it('rejects an oversized inflated page package before any host write', async () => {
  const f = fixture()
  const zip = new JSZip()
  zip.file('ppt/media/image1.png', new Uint8Array(11 * 1024 * 1024))
  f.artifact.pagePptxBase64![1] = await zip.generateAsync({
    type: 'base64',
    compression: 'DEFLATE',
  })
  expect(await f.skill.executeTool(f.call)).toMatchObject({
    isError: true,
    output: 'presentation_import_state_invalid',
  })
  expect(f.adapter.insertPage).not.toHaveBeenCalled()
  expect(f.proposals.pending()).toBeUndefined()
})

it('rejects a source selector absent from an otherwise valid single-page PPTX', async () => {
  const f = fixture()
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const compiled = await compilePresentationDeck(deck)
  f.artifact.pagePptxBase64 = [Buffer.from(compiled.bytes).toString('base64')]
  f.artifact.pages = [{ id: 'page0', title: 'Page 0', sourceSlideId: '999#' }]
  f.artifact.slideCount = 1
  expect(await f.skill.executeTool(f.call)).toMatchObject({
    isError: true,
    output: 'presentation_import_state_invalid',
  })
  expect(f.adapter.insertPage).not.toHaveBeenCalled()
})

it('rejects ZIP content that is not a selectable PowerPoint page before proposing', async () => {
  for (const mutate of [
    async () =>
      new JSZip().file('ppt/slides/slide1.xml', '<page/>').generateAsync({ type: 'base64' }),
    async () => {
      const zip = await JSZip.loadAsync(Buffer.from(fixturePage, 'base64'))
      zip.remove('ppt/_rels/presentation.xml.rels')
      return zip.generateAsync({ type: 'base64' })
    },
  ]) {
    const f = fixture()
    f.artifact.pagePptxBase64![0] = await mutate()
    expect(await f.skill.executeTool(f.call)).toMatchObject({
      isError: true,
      output: 'presentation_import_state_invalid',
    })
    expect(f.proposals.pending()).toBeUndefined()
    expect(f.adapter.insertPage).not.toHaveBeenCalled()
  }
})

it('rejects a source page with missing image or chart parts before proposing', async () => {
  const deck = benchmarkDeck()
  for (const slide of [deck.slides[2]!, deck.slides[6]!]) {
    const compiled = await compilePresentationDeck({ ...deck, slides: [slide] })
    const zip = await JSZip.loadAsync(compiled.bytes)
    const missing = Object.keys(zip.files).find((path) =>
      slide === deck.slides[2]
        ? path.startsWith('ppt/media/')
        : /^ppt\/charts\/chart\d+\.xml$/.test(path),
    )
    expect(missing).toBeDefined()
    zip.remove(missing!)
    const f = fixture()
    f.artifact.slideCount = 1
    f.artifact.pages = [{ id: slide.id, title: slide.title, sourceSlideId: '256#' }]
    f.artifact.pagePptxBase64 = [await zip.generateAsync({ type: 'base64' })]
    expect(await f.skill.executeTool(f.call)).toMatchObject({
      isError: true,
      output: 'presentation_import_state_invalid',
    })
    expect(f.proposals.pending()).toBeUndefined()
    expect(f.adapter.insertPage).not.toHaveBeenCalled()
  }
})

it('rejects a picture page whose relationship definition is missing', async () => {
  const deck = benchmarkDeck()
  const compiled = await compilePresentationDeck({ ...deck, slides: [deck.slides[2]!] })
  const zip = await JSZip.loadAsync(compiled.bytes)
  zip.remove('ppt/slides/_rels/slide1.xml.rels')
  const f = fixture()
  f.artifact.slideCount = 1
  f.artifact.pages = [{ id: 'slide-3', title: '研究图文', sourceSlideId: '256#' }]
  f.artifact.pagePptxBase64 = [await zip.generateAsync({ type: 'base64' })]
  expect(await f.skill.executeTool(f.call)).toMatchObject({
    isError: true,
    output: 'presentation_import_state_invalid',
  })
  expect(f.adapter.insertPage).not.toHaveBeenCalled()
})

it('records an imported production page after its exported package matches', async () => {
  const f = fixture()
  const source = fixturePage
  f.artifact.pagePptxBase64 = [source, source, source]
  const exportPage = vi.fn(async () => source)
  const skill = createPresentationProductionDeliverySkill({
    ...f.options,
    adapter: { ...f.adapter, exportPage },
  })
  expect((await skill.executeTool(f.call)).isError).not.toBe(true)
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(exportPage).toHaveBeenCalledTimes(3)
  expect(f.receipts.get('production/project/request')?.state).toBe('complete')
})

it('keeps the page uncertain if host slide order changes during package readback', async () => {
  const f = fixture()
  const source = fixturePage
  f.artifact.pagePptxBase64 = [source, source, source]
  const skill = createPresentationProductionDeliverySkill({
    ...f.options,
    adapter: {
      ...f.adapter,
      exportPage: async () => {
        f.host.push('concurrent-edit')
        return source
      },
    },
  })
  expect((await skill.executeTool(f.call)).isError).not.toBe(true)
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow(
    'office_state_uncertain',
  )
  expect(f.receipts.get('production/project/request')?.checkpoint).toMatchObject({
    completed: [],
    inFlight: { sourceSlideId: '256#' },
  })
})

it('detects changed native image bytes in a compiled SlideIR page after import', async () => {
  const f = fixture()
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const compiled = await compilePresentationDeck(deck)
  const source = Buffer.from(compiled.bytes).toString('base64')
  const changed = await JSZip.loadAsync(compiled.bytes)
  const mediaPath = Object.keys(changed.files).find((path) =>
    /^ppt\/media\/[^/]+\.png$/.test(path),
  )!
  expect(mediaPath).toBeDefined()
  changed.file(mediaPath, new Uint8Array([1, 2, 3]))
  f.artifact.pagePptxBase64 = [source, source, source]
  const skill = createPresentationProductionDeliverySkill({
    ...f.options,
    adapter: { ...f.adapter, exportPage: async () => changed.generateAsync({ type: 'base64' }) },
  })
  expect((await skill.executeTool(f.call)).isError).not.toBe(true)
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow(
    'office_state_uncertain',
  )
  expect(f.receipts.get('production/project/request')?.checkpoint?.completed).toEqual([])
})

it('records a compiled image page when PowerPoint renumbers an internal shape ID', async () => {
  const f = fixture()
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const compiled = await compilePresentationDeck(deck)
  const source = Buffer.from(compiled.bytes).toString('base64')
  const renumbered = await JSZip.loadAsync(compiled.bytes)
  const path = 'ppt/slides/slide1.xml'
  const original = await renumbered.file(path)!.async('string')
  const changed = original.replace(/(<p:cNvPr\s+id=")\d+/, (_, prefix: string) => `${prefix}99`)
  expect(changed).not.toBe(original)
  renumbered.file(path, changed)
  f.artifact.pagePptxBase64 = [source, source, source]
  const skill = createPresentationProductionDeliverySkill({
    ...f.options,
    adapter: { ...f.adapter, exportPage: async () => renumbered.generateAsync({ type: 'base64' }) },
  })
  expect((await skill.executeTool(f.call)).isError).not.toBe(true)
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(f.receipts.get('production/project/request')?.state).toBe('complete')
})
it('reconciles an interrupted append only after exact package and host-order proof', async () => {
  const f = fixture()
  const source = fixturePage
  f.artifact.pagePptxBase64![0] = source
  f.adapter.insertPage.mockImplementationOnce(async () => {
    f.host.push('host1')
    throw new Error('office_state_uncertain')
  })
  await expect(f.confirm()).rejects.toThrow('office_state_uncertain')
  const skill = createPresentationProductionDeliverySkill({
    ...f.options,
    adapter: { ...f.adapter, exportPage: vi.fn(async () => source) },
  })
  const result = await skill.executeTool({
    id: 'reconcile',
    name: 'reconcile_presentation_production_import',
    input: {},
  })
  expect(result.isError).not.toBe(true)
  expect(JSON.parse(result.output)).toMatchObject({ completed: 1, status: 'partial' })
  expect(f.receipts.get('production/project/request')?.checkpoint).toMatchObject({
    completed: [{ slideId: 'host1', sourceSlideId: '256#' }],
  })
  expect(f.receipts.get('production/project/request')?.checkpoint?.inFlight).toBeUndefined()
  expect(f.adapter.insertPage).toHaveBeenCalledOnce()
  await f.confirm()
  expect(f.host).toEqual(['old', 'host1', 'host2', 'host3'])
})
it('keeps an interrupted append uncertain when package content differs', async () => {
  const f = fixture()
  f.artifact.pagePptxBase64![0] = fixturePage
  const otherZip = await JSZip.loadAsync(Buffer.from(fixturePage, 'base64'))
  const xml = await otherZip.file('ppt/slides/slide1.xml')!.async('string')
  otherZip.file('ppt/slides/slide1.xml', xml.replace('科研汇报', '内容不同'))
  const other = await otherZip.generateAsync({ type: 'base64' })
  f.adapter.insertPage.mockImplementationOnce(async () => {
    f.host.push('host1')
    throw new Error('office_state_uncertain')
  })
  await expect(f.confirm()).rejects.toThrow('office_state_uncertain')
  const skill = createPresentationProductionDeliverySkill({
    ...f.options,
    adapter: { ...f.adapter, exportPage: vi.fn(async () => other) },
  })
  expect(
    await skill.executeTool({
      id: 'reconcile',
      name: 'reconcile_presentation_production_import',
      input: {},
    }),
  ).toMatchObject({ isError: true, output: 'presentation_import_uncertain' })
  expect(f.receipts.get('production/project/request')?.checkpoint?.inFlight).toBeDefined()
  expect(f.adapter.insertPage).toHaveBeenCalledOnce()
})
it('correlates a real page import receipt with the interrupted AgentRun without replay', async () => {
  const f = fixture()
  const values = new Map<string, string>()
  const binding = createPresentationDocumentBinding(
    {
      get: (key) => values.get(key),
      set: (key, value) => {
        values.set(key, value)
      },
      save: async () => undefined,
      location: () => 'file:///deck.pptx',
    },
    () => 'doc-id',
  )
  const documentId = await binding.documentId()
  f.artifact.documentId = documentId
  const checkpoint = createPresentationAgentRunCheckpoint(binding, documentId)
  await checkpoint.begin('run-1')
  await checkpoint.tool('run-1', 'tool_pending', f.call.name, false, f.call.id)
  const skill = createPresentationProductionDeliverySkill({
    ...f.options,
    documentId: async () => documentId,
    readReceipt: binding.readReceipt,
    writeReceipt: binding.writeReceipt,
  })
  expect((await skill.executeTool(f.call)).isError).not.toBe(true)
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(f.adapter.insertPage).toHaveBeenCalledTimes(3)
  expect(checkpoint.recovery()?.importReceipt).toEqual({
    state: 'complete',
    completed: 3,
    total: 3,
  })
  expect(binding.readReceipt('production/project/request')).toMatchObject({
    toolCallId: f.call.id,
    agentRunId: 'run-1',
    state: 'complete',
  })
})
it('imports separate PPTX files with duplicate source IDs and isolates their checkpoint namespace', async () => {
  const f = fixture()
  f.receipts.set('project/request', { state: 'complete', documentId: 'doc', slideIds: ['legacy'] })
  await f.confirm()
  expect(f.adapter.insert).not.toHaveBeenCalled()
  expect(f.adapter.insertPage.mock.calls.map((call) => call.slice(0, 2))).toEqual(
    f.artifact.pagePptxBase64!.map((bytes) => [bytes, '256#']),
  )
  expect(presentationImportKey(f.artifact)).toBe('production/project/request')
  expect(f.receipts.get('production/project/request')).toMatchObject({
    state: 'complete',
    toolCallId: 'import',
    checkpoint: {
      version: 2,
      pageIds: ['page0', 'page1', 'page2'],
      sourceSlideIds: ['256#', '256#', '256#'],
    },
  })
  expect(
    f.receipts
      .get('production/project/request')
      ?.checkpoint?.completed.every((page) => typeof page.completedAt === 'string'),
  ).toBe(true)
  expect(f.receipts.get('project/request')!.slideIds).toEqual(['legacy'])
  expect(await f.skill.executeTool(f.call)).toMatchObject({
    output: expect.stringContaining('already_imported'),
  })
})
it('resumes only the remaining page files after a known pre-write failure', async () => {
  const f = fixture(),
    insert = f.adapter.insertPage.getMockImplementation()!
  f.adapter.insertPage.mockImplementationOnce(insert).mockRejectedValueOnce(new Error('cancelled'))
  await expect(f.confirm()).rejects.toThrow('cancelled')
  expect(f.receipts.get(presentationImportKey(f.artifact))?.checkpoint?.completed).toHaveLength(1)
  await f.confirm()
  expect(f.host).toEqual(['old', 'host1', 'host2', 'host3'])
  expect(f.adapter.insertPage.mock.calls.map((call) => call[0])).toEqual([
    f.artifact.pagePptxBase64![0],
    f.artifact.pagePptxBase64![1],
    f.artifact.pagePptxBase64![1],
    f.artifact.pagePptxBase64![2],
  ])
})
it('does not replay an uncertain page or permit bundle identity changes', async () => {
  const f = fixture()
  f.adapter.insertPage.mockRejectedValueOnce(new Error('office_state_uncertain'))
  await expect(f.confirm()).rejects.toThrow('office_state_uncertain')
  expect(await f.skill.executeTool(f.call)).toMatchObject({
    isError: true,
    output: 'presentation_import_uncertain',
  })
  expect(f.adapter.insertPage).toHaveBeenCalledTimes(1)
  for (const change of [
    () => {
      f.artifact.pagePptxBase64![0] = 'UEsDBAMAAAA='
    },
    () => {
      f.artifact.pages![0]!.id = 'changed'
    },
    () => {
      f.artifact.planRevision = 2
    },
  ]) {
    change()
    expect(await f.skill.executeTool(f.call)).toMatchObject({
      isError: true,
      output: 'presentation_import_state_invalid',
    })
  }
})
it('rejects a changed page order after a proposal and never sends bundles to bulk fallback', async () => {
  const f = fixture()
  await f.skill.executeTool(f.call)
  f.artifact.pages!.reverse()
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('proposal_stale')
  expect(f.adapter.insertPage).not.toHaveBeenCalled()
  const old = createPresentationDeliverySkill({
    ...f.options,
    adapter: { ...f.adapter, insertPage: undefined },
  })
  expect(await old.executeTool({ ...f.call, name: 'import_generated_presentation' })).toMatchObject(
    { isError: true, output: 'presentation_import_state_invalid' },
  )
  expect(f.adapter.insert).not.toHaveBeenCalled()
})
it('strictly separates checkpoint versions and persists v2 only under production namespace', async () => {
  const f = fixture()
  await f.confirm()
  const saved = f.receipts.get(presentationImportKey(f.artifact))!
  expect(validPresentationImportRecord(saved)).toBe(true)
  expect(
    validPresentationImportRecord({ ...saved, checkpoint: { ...saved.checkpoint, version: 1 } }),
  ).toBe(false)
  expect(
    validPresentationImportRecord({
      ...saved,
      checkpoint: { ...saved.checkpoint, pageIds: ['same', 'same', 'same'] },
    }),
  ).toBe(false)
  const values = new Map<string, string>(),
    settings = {
      get: (key: string) => values.get(key),
      set: (key: string, value: string) => {
        values.set(key, value)
      },
      save: async () => {},
      location: () => 'test',
    }
  const binding = createPresentationDocumentBinding(settings, () => 'doc')
  await binding.writeReceipt('production/project/request', saved)
  expect(
    createPresentationDocumentBinding(settings, () => 'doc').readReceipt(
      'production/project/request',
    ),
  ).toEqual(saved)
  await expect(binding.writeReceipt('project/request', saved)).rejects.toThrow(
    'presentation_import_state_invalid',
  )
  await expect(
    binding.writeReceipt('production/project/legacy', { state: 'pending', documentId: 'doc' }),
  ).rejects.toThrow('presentation_import_state_invalid')
})
it('rejects v2 checkpoints persisted under the legacy namespace on settings reload', async () => {
  const f = fixture()
  await f.confirm()
  const saved = f.receipts.get(presentationImportKey(f.artifact))!
  const values = new Map<string, string>(),
    settings = {
      get: (key: string) => values.get(key),
      set: (key: string, value: string) => {
        values.set(key, value)
      },
      save: async () => {},
      location: () => 'test',
    }
  values.set('wiswork.presentation.imports.v1', JSON.stringify({ 'project/request': saved }))
  expect(() => createPresentationDocumentBinding(settings).readReceipt('project/request')).toThrow(
    'presentation_import_state_invalid',
  )
  values.set(
    'wiswork.presentation.imports.v1',
    JSON.stringify({ 'production/project/request': { state: 'pending', documentId: 'doc' } }),
  )
  expect(() =>
    createPresentationDocumentBinding(settings).readReceipt('production/project/request'),
  ).toThrow('presentation_import_state_invalid')
})
it('blocks a byte change during awaited document checks and identifies progress as production', async () => {
  const f = fixture()
  const progress = await f.skill.executeTool({
    id: 'status',
    name: 'read_presentation_production_import_status',
    input: {},
  })
  expect(JSON.parse(progress.output)).toMatchObject({ source: 'production', status: 'not_started' })
  await f.skill.executeTool(f.call)
  const original = f.options.documentId
  let checks = 0
  f.options.documentId = async () => {
    if (++checks === 1) f.artifact.pagePptxBase64![0] = 'UEsDBAMAAAA='
    return original()
  }
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('proposal_stale')
  expect(f.adapter.insertPage).not.toHaveBeenCalled()
})
it('rejects production bundles with missing revision, mismatched files or invalid base64', async () => {
  for (const mutate of [
    (artifact: CompiledPresentationArtifact) => {
      delete artifact.planRevision
    },
    (artifact: CompiledPresentationArtifact) => {
      artifact.pagePptxBase64!.pop()
    },
    (artifact: CompiledPresentationArtifact) => {
      artifact.pagePptxBase64![0] = '%%%invalid'
    },
    (artifact: CompiledPresentationArtifact) => {
      artifact.pptxBase64 = 'UEsDBAAAAAA='
    },
  ]) {
    const f = fixture()
    mutate(f.artifact)
    expect(await f.skill.executeTool(f.call)).toMatchObject({
      isError: true,
      output: 'presentation_import_state_invalid',
    })
    expect(f.adapter.insertPage).not.toHaveBeenCalled()
  }
})
it('preserves legacy two-segment keys whose project ID is production', async () => {
  const values = new Map<string, string>(),
    settings = {
      get: (key: string) => values.get(key),
      set: (key: string, value: string) => {
        values.set(key, value)
      },
      save: async () => {},
      location: () => 'test',
    }
  const binding = createPresentationDocumentBinding(settings)
  const record: PresentationImportRecord = { state: 'pending', documentId: 'doc' }
  await binding.writeReceipt('production/request', record)
  expect(createPresentationDocumentBinding(settings).readReceipt('production/request')).toEqual(
    record,
  )
  const f = fixture()
  await f.confirm()
  const source = f.receipts.get(presentationImportKey(f.artifact))!
  const v1: PresentationImportRecord = {
    ...source,
    checkpoint: {
      version: 1,
      artifactDigest: source.checkpoint!.artifactDigest,
      sourceSlideIds: ['256#', '257#', '258#'],
      baselineSlideIds: ['old'],
      completed: source.checkpoint!.completed.map((page, i) => ({
        ...page,
        sourceSlideId: `${256 + i}#`,
      })),
    },
  }
  await binding.writeReceipt('production/request-v1', v1)
  expect(createPresentationDocumentBinding(settings).readReceipt('production/request-v1')).toEqual(
    v1,
  )
})

it('shares the exact persisted digest and maps duplicate selectors by business page identity', async () => {
  const f = fixture()
  await f.confirm()
  const record = f.receipts.get(presentationImportKey(f.artifact))!
  const digest = Array.from(
    new Uint8Array(
      await crypto.subtle.digest(
        'SHA-256',
        new TextEncoder().encode(presentationArtifactContent(f.artifact)),
      ),
    ),
    (b) => b.toString(16).padStart(2, '0'),
  ).join('')
  expect(digest).toBe(record.checkpoint!.artifactDigest)
  expect(presentationPageMapping(f.artifact, record, 'page1')).toEqual({
    sourceSlideId: '256#',
    slideId: 'host2',
  })
  const partial = structuredClone(record)
  partial.state = 'pending'
  delete partial.slideIds
  partial.checkpoint!.completed = partial.checkpoint!.completed.slice(0, 1)
  partial.checkpoint!.inFlight = { sourceSlideId: '256#' }
  expect(presentationPageMapping(f.artifact, partial, 'page1')).toBeUndefined()
  expect(presentationPageMapping(f.artifact, partial, 'page0')?.slideId).toBe('host1')
  const changed = structuredClone(f.artifact)
  changed.pages!.reverse()
  expect(presentationPageMapping(changed, record, 'page1')).toBeUndefined()
  expect(
    presentationPageMapping(f.artifact, { ...record, documentId: 'other' }, 'page1'),
  ).toBeUndefined()
})

it('accepts legacy undated import prefixes and validates later completion time order', async () => {
  const f = fixture()
  await f.confirm()
  const saved = structuredClone(f.receipts.get(presentationImportKey(f.artifact))!)
  delete saved.checkpoint!.completed[0]!.completedAt
  delete saved.checkpoint!.completed[0]!.startedAt
  expect(validPresentationImportRecord(saved)).toBe(true)
  const reordered = structuredClone(saved)
  reordered.checkpoint!.completed[2]!.completedAt = '2020-01-01T00:00:00.000Z'
  expect(validPresentationImportRecord(reordered)).toBe(false)
  const missingSuffix = structuredClone(saved)
  delete missingSuffix.checkpoint!.completed[2]!.completedAt
  expect(validPresentationImportRecord(missingSuffix)).toBe(false)
})
