import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import {
  createPresentationProductionDeliverySkill,
  validPresentationImportRecord,
} from '../src/skills/powerpoint/presentation-page-delivery'
import { expect, it, vi } from 'vitest'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import {
  createPresentationDeliverySkill,
  type PresentationImportRecord,
  type PresentationDeliveryOptions,
} from '../src/skills/powerpoint/presentation-delivery.js'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document.js'
function fixture() {
  const host = ['original'],
    receipts = new Map<string, PresentationImportRecord>(),
    proposals = createStructuredProposalController()
  const artifact = {
    documentId: 'doc',
    projectId: 'project',
    requestId: 'request',
    pptxBase64: 'UEsDBAAAAAA=',
    slideCount: 3,
    pages: [0, 1, 2].map((i) => ({
      id: `page${i}`,
      title: `Page ${i}`,
      sourceSlideId: `${256 + i}#`,
    })),
  }
  const adapter = {
    available: () => true,
    snapshot: vi.fn(async () => ({ slideIds: [...host], fingerprint: JSON.stringify(host) })),
    insert: vi.fn(),
    insertPage: vi.fn(async (_bytes: string, id: string) => {
      const slideId = `host-${id}`
      host.push(slideId)
      return { slideIds: [slideId] }
    }),
    verify: vi.fn(async () => true),
  }
  const documentId = vi.fn(async () => 'doc'),
    readReceipt = (key: string) => receipts.get(key),
    writeReceipt = vi.fn(async (key: string, record: PresentationImportRecord | undefined) => {
      if (record) receipts.set(key, structuredClone(record))
      else receipts.delete(key)
    })
  const options: PresentationDeliveryOptions = {
    adapter,
    proposals,
    available: () => true,
    artifact: () => artifact,
    documentId,
    readReceipt,
    writeReceipt,
  }
  const skill = createPresentationDeliverySkill(options),
    call = { id: 'import', name: 'import_generated_presentation', input: { project_id: 'project' } }
  const confirm = async () => {
    await skill.executeTool(call)
    return proposals.confirm(proposals.pending()!.id)
  }
  return {
    host,
    receipts,
    proposals,
    artifact,
    adapter,
    documentId,
    writeReceipt,
    options,
    skill,
    call,
    confirm,
  }
}
it('checkpoints each verified page and never reimports a completed deck', async () => {
  const f = fixture()
  await f.confirm()
  expect(f.adapter.insert).not.toHaveBeenCalled()
  expect(f.adapter.insertPage).toHaveBeenCalledTimes(3)
  expect(f.receipts.get('project/request')).toMatchObject({
    state: 'complete',
    slideIds: ['host-256#', 'host-257#', 'host-258#'],
  })
  const completed = f.receipts.get('project/request')!.checkpoint!.completed
  expect(completed.map((page) => page.completedAt)).toEqual([
    expect.any(String),
    expect.any(String),
    expect.any(String),
  ])
  expect(
    completed.every(
      (page, index) => index === 0 || page.completedAt! >= completed[index - 1]!.completedAt!,
    ),
  ).toBe(true)
  expect(await f.skill.executeTool(f.call)).toMatchObject({
    output: expect.stringContaining('already_imported'),
  })
  const status = await f.skill.executeTool({
    id: 'status',
    name: 'read_presentation_import_status',
    input: {},
  })
  expect(JSON.parse(status.output)).toMatchObject({ status: 'complete', completed: 3, total: 3 })
})
it('resumes after a safe failure with a fresh confirmation and preserves earlier pages', async () => {
  const f = fixture(),
    original = f.adapter.insertPage.getMockImplementation()!
  f.adapter.insertPage
    .mockImplementationOnce(original)
    .mockRejectedValueOnce(new Error('cancelled'))
  await expect(f.confirm()).rejects.toThrow('cancelled')
  expect(f.receipts.get('project/request')?.checkpoint).toMatchObject({
    completed: [{ sourceSlideId: '256#', slideId: 'host-256#' }],
  })
  expect(f.receipts.get('project/request')?.checkpoint?.inFlight).toBeUndefined()
  const restored = createPresentationDeliverySkill(f.options)
  expect(await restored.executeTool(f.call)).toMatchObject({
    output: expect.stringContaining('awaiting_confirmation'),
  })
  expect(f.adapter.insertPage).toHaveBeenCalledTimes(2)
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(f.host).toEqual(['original', 'host-256#', 'host-257#', 'host-258#'])
})
it('preserves uncertainty across restart and refuses changed bytes or source IDs', async () => {
  const f = fixture()
  f.adapter.verify.mockResolvedValue(false)
  await expect(f.confirm()).rejects.toThrow('office_state_uncertain')
  expect(f.receipts.get('project/request')?.checkpoint?.inFlight).toMatchObject({
    sourceSlideId: '256#',
    startedAt: expect.any(String),
  })
  const progress = await f.skill.executeTool({
    id: 'status',
    name: 'read_presentation_import_status',
    input: {},
  })
  expect(JSON.parse(progress.output).pages[0]).toMatchObject({
    state: 'uncertain',
    startedAt: expect.any(String),
  })
  expect(await createPresentationDeliverySkill(f.options).executeTool(f.call)).toMatchObject({
    isError: true,
    output: 'presentation_import_uncertain',
  })
  f.artifact.pptxBase64 += 'AA'
  expect(await createPresentationDeliverySkill(f.options).executeTool(f.call)).toMatchObject({
    isError: true,
    output: 'presentation_import_state_invalid',
  })
  expect(f.adapter.insertPage).toHaveBeenCalledOnce()
})
it('rejects a changed host or foreign document when resuming', async () => {
  const f = fixture()
  f.adapter.insertPage.mockRejectedValueOnce(new Error('cancelled'))
  await expect(f.confirm()).rejects.toThrow('cancelled')
  f.host.push('manual')
  expect(await f.skill.executeTool(f.call)).toMatchObject({
    isError: true,
    output: 'presentation_import_host_changed',
  })
  f.documentId.mockResolvedValue('other')
  expect(await f.skill.executeTool(f.call)).toMatchObject({
    isError: true,
    output: 'presentation_document_changed',
  })
})
it('restores the previous in-flight settings after a failed completed checkpoint save', async () => {
  const f = fixture(),
    settings = new Map<string, unknown>(),
    save = vi.fn(async () => {})
  const binding = createPresentationDocumentBinding({
    get: (key) => settings.get(key),
    set: (key, value) => {
      settings.set(key, value)
    },
    save,
    location: () => '',
  })
  f.options.readReceipt = binding.readReceipt
  f.options.writeReceipt = binding.writeReceipt
  const skill = createPresentationDeliverySkill(f.options)
  save
    .mockResolvedValueOnce()
    .mockResolvedValueOnce()
    .mockRejectedValueOnce(new Error('save_failed'))
  await skill.executeTool(f.call)
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow()
  expect(binding.readReceipt('project/request')?.checkpoint?.inFlight).toMatchObject({
    sourceSlideId: '256#',
    startedAt: expect.any(String),
  })
  expect(await skill.executeTool(f.call)).toMatchObject({
    isError: true,
    output: 'presentation_import_uncertain',
  })
  expect(f.adapter.insertPage).toHaveBeenCalledOnce()
})
it('rejects source changes after proposal and stops real cancellation after recording the committed page', async () => {
  const f = fixture()
  await f.skill.executeTool(f.call)
  f.artifact.pages[0]!.sourceSlideId = '999#'
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('proposal_stale')
  expect(f.adapter.insertPage).not.toHaveBeenCalled()
  const g = fixture(),
    original = g.adapter.insertPage.getMockImplementation()!
  g.adapter.insertPage.mockImplementationOnce(async (...args) => {
    const receipt = await original(...args)
    g.proposals.newTurn()
    return receipt
  })
  await expect(g.confirm()).rejects.toThrow('cancelled')
  expect(g.receipts.get('project/request')?.checkpoint).toMatchObject({
    completed: [{ sourceSlideId: '256#', slideId: 'host-256#' }],
  })
  expect(g.receipts.get('project/request')?.checkpoint?.inFlight).toBeUndefined()
})
it('validates persisted checkpoints and restores no reservation when its first save fails', async () => {
  const f = fixture(),
    settings = new Map<string, unknown>(),
    save = vi.fn(async () => {})
  const binding = createPresentationDocumentBinding({
    get: (key) => settings.get(key),
    set: (key, value) => {
      settings.set(key, value)
    },
    save,
    location: () => '',
  })
  await expect(
    binding.writeReceipt('project/request', {
      state: 'pending',
      documentId: 'doc',
      checkpoint: null,
    } as unknown as PresentationImportRecord),
  ).rejects.toThrow('presentation_import_state_invalid')
  const skill = createPresentationDeliverySkill({
    ...f.options,
    readReceipt: binding.readReceipt,
    writeReceipt: binding.writeReceipt,
  })
  save.mockRejectedValueOnce(new Error('save_failed'))
  await skill.executeTool(f.call)
  await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow()
  expect(binding.readReceipt('project/request')).toBeUndefined()
  expect(f.adapter.insertPage).not.toHaveBeenCalled()
})
it('blocks malformed checkpoint prefixes, duplicate host IDs and invalid source identifiers', async () => {
  const settings = new Map<string, unknown>()
  const binding = createPresentationDocumentBinding({
    get: (key) => settings.get(key),
    set: (key, value) => {
      settings.set(key, value)
    },
    save: async () => {},
    location: () => '',
  })
  const checkpoint = {
    version: 1,
    artifactDigest: 'a'.repeat(64),
    sourceSlideIds: ['256#', '257#'],
    baselineSlideIds: ['original'],
    completed: [],
  }
  for (const bad of [
    { ...checkpoint, inFlight: null },
    { ...checkpoint, sourceSlideIds: ['255#'] },
    { ...checkpoint, sourceSlideIds: ['4294967296#'] },
    { ...checkpoint, completed: [{ sourceSlideId: '257#', slideId: 'new' }] },
    { ...checkpoint, completed: [{ sourceSlideId: '256#', slideId: 'original' }] },
    {
      ...checkpoint,
      completed: [{ sourceSlideId: '256#', slideId: 'new', completedAt: 'tomorrow' }],
    },
    { ...checkpoint, inFlight: { sourceSlideId: '257#' } },
    { ...checkpoint, inFlight: { sourceSlideId: '256#', startedAt: 'tomorrow' } },
  ]) {
    await expect(
      binding.writeReceipt('project/request', {
        state: 'pending',
        documentId: 'doc',
        checkpoint: bad,
      } as unknown as PresentationImportRecord),
    ).rejects.toThrow('presentation_import_state_invalid')
  }
  expect(settings.size).toBe(0)
})
it('preserves each actual reserved start on normal completion and after reopen, even if the clock rolls back', async () => {
  const f = fixture()
  let clock = Date.parse('2099-01-01T00:00:00.000Z')
  const spy = vi.spyOn(Date, 'now').mockImplementation(() => clock)
  const original = f.adapter.insertPage.getMockImplementation()!
  f.adapter.insertPage.mockImplementation(async (...args) => {
    clock = Date.parse('2020-01-01T00:00:00.000Z')
    return original(...args)
  })
  try {
    await f.confirm()
    const starts = f.writeReceipt.mock.calls
      .map(([, record]) => record?.checkpoint?.inFlight?.startedAt)
      .filter(Boolean)
    const completed = f.receipts.get('project/request')!.checkpoint!.completed
    expect(completed.map((page) => page.startedAt)).toEqual(starts)
    expect(starts).toHaveLength(3)
    expect(completed.every((page) => page.startedAt! <= page.completedAt!)).toBe(true)
    const reopened = createPresentationDeliverySkill(f.options)
    const result = await reopened.executeTool({
      id: 'read',
      name: 'read_presentation_import_status',
      input: {},
    })
    expect(
      JSON.parse(result.output).pages.map((page: { startedAt: string; completedAt: string }) => [
        page.startedAt,
        page.completedAt,
      ]),
    ).toEqual(completed.map((page) => [page.startedAt, page.completedAt]))
  } finally {
    spy.mockRestore()
  }
})
it('retains the original start when read-only reconciliation claims a lost host ACK; legacy starts remain absent', async () => {
  for (const legacy of [false, true]) {
    const f = fixture(),
      deck = benchmarkDeck()
    deck.slides = [deck.slides[0]!]
    const source = Buffer.from((await compilePresentationDeck(deck)).bytes).toString('base64')
    const artifact = {
      ...f.artifact,
      pptxBase64: '',
      planRevision: 1,
      pagePptxBase64: [source, source, source],
      pages: f.artifact.pages.map((page) => ({ ...page, sourceSlideId: '256#' })),
    }
    const options = {
      ...f.options,
      artifact: () => artifact,
      adapter: { ...f.adapter, exportPage: async () => source },
    }
    const skill = createPresentationProductionDeliverySkill(options)
    f.adapter.insertPage.mockImplementationOnce(async () => {
      f.host.push('lost-host')
      throw Error('office_state_uncertain')
    })
    const started = await skill.executeTool({
      id: 'import',
      name: 'import_presentation_production',
      input: { project_id: 'project' },
    })
    expect(started.isError).not.toBe(true)
    await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow(
      'office_state_uncertain',
    )
    const pending = f.receipts.get('production/project/request')!
    const originalStart = pending.checkpoint!.inFlight!.startedAt
    if (legacy) delete pending.checkpoint!.inFlight!.startedAt
    const reopened = createPresentationProductionDeliverySkill({
      ...options,
      adapter: { ...f.adapter, exportPage: async () => source },
    })
    const result = await reopened.executeTool({
      id: 'reconcile',
      name: 'reconcile_presentation_production_import',
      input: {},
    })
    expect(result.isError, result.output).not.toBe(true)
    const completed = f.receipts.get('production/project/request')!.checkpoint!.completed[0]!
    if (legacy) {
      expect(completed).not.toHaveProperty('startedAt')
      expect(JSON.parse(result.output).pages[0]).not.toHaveProperty('startedAt')
    } else {
      expect(completed.startedAt).toBe(originalStart)
      expect(JSON.parse(result.output).pages[0].startedAt).toBe(originalStart)
    }
    expect(completed.completedAt).toEqual(expect.any(String))
    expect(f.adapter.insertPage).toHaveBeenCalledOnce()
  }
})
it('strictly validates dual event timestamps and ordering without rejecting legacy completion records', () => {
  const time = '2026-09-29T00:00:00.000Z',
    later = '2026-09-29T00:00:01.000Z'
  const page = { sourceSlideId: '256#', slideId: 'new', startedAt: time, completedAt: later }
  const record = {
    state: 'pending',
    documentId: 'doc',
    checkpoint: {
      version: 1,
      artifactDigest: 'a'.repeat(64),
      sourceSlideIds: ['256#', '257#'],
      baselineSlideIds: ['old'],
      completed: [page],
    },
  }
  expect(validPresentationImportRecord(record)).toBe(true)
  for (const patch of [
    { startedAt: '2026-09-29T00:00:00Z' },
    { startedAt: '2026-02-30T00:00:00.000Z' },
    { startedAt: undefined },
    { startedAt: later, completedAt: time },
    { completedAt: undefined },
    { extra: true },
  ])
    expect(
      validPresentationImportRecord({
        ...record,
        checkpoint: { ...record.checkpoint, completed: [{ ...page, ...patch }] },
      }),
    ).toBe(false)
  expect(
    validPresentationImportRecord({
      ...record,
      checkpoint: { ...record.checkpoint, completed: [{ sourceSlideId: '256#', slideId: 'new' }] },
    }),
  ).toBe(true)
  expect(
    validPresentationImportRecord({
      ...record,
      checkpoint: {
        ...record.checkpoint,
        completed: [
          page,
          { sourceSlideId: '257#', slideId: 'next', startedAt: time, completedAt: later },
        ],
      },
    }),
  ).toBe(false)
})
