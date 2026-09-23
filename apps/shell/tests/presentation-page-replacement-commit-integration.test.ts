import { createHash } from 'node:crypto'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it, vi } from 'vitest'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'
import { createOfficeHostRuntime } from '../../office-addin/src/agent/host-runtime'
import { createPresentationDocumentBinding } from '../../office-addin/src/skills/powerpoint/presentation-document'
import { presentationArtifactContent } from '../../office-addin/src/skills/powerpoint/presentation-page-delivery'
import type {
  CompiledPresentationArtifact,
  PresentationImportRecord,
} from '../../office-addin/src/skills/powerpoint/presentation-delivery'

it('commits and undoes through real adapters and durable maps, recovering terminal save failures without duplicate host writes', async () => {
  const root = mkdtempSync(join(tmpdir(), 'wiswork-commit-undo-'))
  const settings = new Map<string, string>([['wiswork.presentation.document.v1', 'docid']])
  let failState: string | undefined
  const binding = createPresentationDocumentBinding({
    get: (key) => settings.get(key),
    set: (key, value) => {
      settings.set(key, value)
    },
    location: () => 'file://deck.pptx',
    save: async () => {
      const raw = settings.get('wiswork.presentation.page-replacement.v1')
      const saved = raw ? JSON.parse(raw) : undefined
      if (failState && (saved?.change?.state ?? saved?.state) === failState) {
        failState = undefined
        throw new Error('simulated_settings_save_failure')
      }
    },
  })
  const documentId = await binding.documentId(),
    deck = benchmarkPlannedDeck(),
    signal = new AbortController().signal
  let service = createPresentationService({ userDataPath: root })
  const call = async (operation: string, extra = {}) =>
    JSON.parse(
      Buffer.from(
        await service({ operation, documentId, projectId: deck.id, ...extra }, signal),
      ).toString(),
    )
  let runtime: ReturnType<typeof createOfficeHostRuntime> | undefined
  const host = new Map<string, string>()
  let ids: string[] = [],
    sequence = 0
  const queued: (() => void)[] = []
  const removed: string[] = [],
    inserted: string[] = []
  const slides = {
    load: vi.fn(),
    get items() {
      return ids.map((id) => ({ id }))
    },
    getItem: (id: string) => ({
      id,
      load: vi.fn(),
      exportAsBase64: () => {
        if (!host.has(id)) throw new Error('ItemNotFound')
        return { value: host.get(id)! }
      },
      delete: () =>
        queued.push(() => {
          removed.push(id)
          ids = ids.filter((x) => x !== id)
          host.delete(id)
        }),
    }),
  }
  const context = {
    presentation: {
      slides,
      insertSlidesFromBase64: (base64: string, options: { targetSlideId: string }) =>
        queued.push(() => {
          const index = ids.indexOf(options.targetSlideId)
          if (index < 0) throw new Error('ItemNotFound')
          const id = `inserted-${++sequence}`
          inserted.push(id)
          host.set(id, base64)
          ids.splice(index + 1, 0, id)
        }),
    },
    sync: async () => {
      while (queued.length) queued.shift()!()
    },
  }
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', requirements: { isSetSupported: () => true } },
  })
  vi.stubGlobal('PowerPoint', {
    run: async (fn: (c: typeof context) => Promise<unknown>) => fn(context),
  })
  try {
    await call('save_plan', { expectedRevision: 0, plan: benchmarkPlan() })
    await call('production_begin', { requestId: 'parent', planRevision: 1, deck })
    expect(await call('production_run', { requestId: 'parent' })).toMatchObject({
      status: 'compiled',
    })
    const pages = await Promise.all(
      deck.slides.map((p) => call('production_page', { requestId: 'parent', pageId: p.id })),
    )
    const target = deck.slides[1]!
    await call('production_rebuild_page', {
      parentRequestId: 'parent',
      requestId: 'child',
      pageId: target.id,
      slide: { ...target, notes: 'Rebuilt page' },
    })
    await call('production_run', { requestId: 'child' })
    const artifact: CompiledPresentationArtifact = {
      documentId,
      projectId: deck.id,
      requestId: 'parent',
      planRevision: 1,
      pptxBase64: '',
      slideCount: deck.slides.length,
      pages: deck.slides.map((p, i) => ({
        id: p.id,
        title: p.title,
        sourceSlideId: pages[i].sourceSlideId,
      })),
      pagePptxBase64: pages.map((p) => p.pptxBase64),
    }
    const hostIds = deck.slides.map((_, i) => `host-${i}`)
    ids = ['existing', ...hostIds]
    for (const [i, id] of hostIds.entries()) host.set(id, pages[i].pptxBase64)
    host.set('existing', pages[0].pptxBase64)
    const manual = await compilePresentationDeck({
      ...deck,
      slides: [{ ...target, notes: 'User correction to preserve on undo' }],
    })
    const original = Buffer.from(manual.bytes).toString('base64')
    host.set(hostIds[1]!, original)
    const receipt: PresentationImportRecord = {
      state: 'complete',
      documentId,
      slideIds: hostIds,
      checkpoint: {
        version: 2,
        pageIds: deck.slides.map((p) => p.id),
        sourceSlideIds: pages.map((p) => p.sourceSlideId),
        baselineSlideIds: ['existing'],
        artifactDigest: createHash('sha256')
          .update(presentationArtifactContent(artifact))
          .digest('hex'),
        completed: hostIds.map((slideId, i) => ({
          slideId,
          sourceSlideId: pages[i].sourceSlideId,
        })),
      },
    }
    const parentKey = `production/${deck.id}/parent`,
      childKey = `production/${deck.id}/child`
    await binding.writeReceipt(parentKey, receipt)
    const create = () =>
      createOfficeHostRuntime('powerpoint', {
        presentation: {
          ...binding,
          available: () => true,
          request: async (body, s) => new Response(Buffer.from(await service(body, s ?? signal))),
        },
      })
    const tool = async (name: string, input: Record<string, unknown>) =>
      runtime!.skill.executeTool({ id: name, name, input: { project_id: deck.id, ...input } })
    const prepare = async (request_id: string) => {
      const r = await tool('prepare_presentation_production_import', { request_id })
      expect(r.isError, r.output).not.toBe(true)
    }
    const propose = async (name: string) => {
      const r = await tool(name, { change_id: 'change' })
      expect(r.isError, r.output).not.toBe(true)
      return runtime!.proposals.pending()!.id
    }
    runtime = create()
    await prepare('parent')
    const saved = await tool('save_presentation_page_backup', {
      request_id: 'child',
      page_id: target.id,
      backup_id: 'backup',
    })
    expect(saved.isError, saved.output).not.toBe(true)
    const staged = await tool('stage_presentation_page_replacement', {
      request_id: 'child',
      page_id: target.id,
      backup_id: 'backup',
      change_id: 'change',
    })
    expect(staged.isError, staged.output).not.toBe(true)
    await runtime.proposals.confirm(runtime.proposals.pending()!.id)
    expect(binding.readPageReplacement()?.state).toBe('staged')
    expect(inserted).toEqual(['inserted-1'])
    const commitId = await propose('commit_presentation_page_replacement')
    failState = 'applied'
    await expect(runtime.proposals.confirm(commitId)).rejects.toThrow(
      'simulated_settings_save_failure',
    )
    expect(binding.readPageReplacement()?.state).toBe('commit_pending')
    expect(removed).toEqual(['host-1'])
    expect(ids).toHaveLength(9)
    expect(runtime.importProgress!.read()).toBeUndefined()
    runtime.dispose()
    service = createPresentationService({ userDataPath: root })
    runtime = create()
    await prepare('parent')
    await runtime.proposals.confirm(await propose('commit_presentation_page_replacement'))
    expect(binding.readPageReplacement()?.state).toBe('applied')
    expect(removed).toEqual(['host-1'])
    expect(() => binding.readReceipt(parentKey)).toThrow('presentation_import_superseded')
    expect(binding.readReceipt(childKey)?.slideIds?.[1]).toBe('inserted-1')
    expect(runtime.importProgress!.read()).toBeUndefined()
    await prepare('child')
    expect(runtime.importProgress!.read()?.pages[1]?.slideId).toBe('inserted-1')
    const parentAgain = await tool('prepare_presentation_production_import', {
      request_id: 'parent',
    })
    expect(parentAgain.isError).toBe(true)
    expect(runtime.importProgress!.read()?.pages[1]?.slideId).toBe('inserted-1')
    const undoId = await propose('undo_presentation_page_replacement')
    failState = 'undone'
    await expect(runtime.proposals.confirm(undoId)).rejects.toThrow(
      'simulated_settings_save_failure',
    )
    expect(binding.readPageReplacement()?.state).toBe('restore_inserted')
    expect(inserted).toEqual(['inserted-1', 'inserted-2'])
    expect(removed).toEqual(['host-1', 'inserted-1'])
    expect(host.get('inserted-2')).toBe(original)
    expect(ids).toEqual(['existing', 'host-0', 'inserted-2', ...hostIds.slice(2)])
    runtime.dispose()
    runtime = create()
    await prepare('child')
    await runtime.proposals.confirm(await propose('undo_presentation_page_replacement'))
    expect(binding.readPageReplacement()?.state).toBe('undone')
    expect(inserted).toHaveLength(2)
    expect(removed).toHaveLength(2)
    expect(() => binding.readReceipt(childKey)).toThrow('presentation_import_superseded')
    await prepare('parent')
    expect(runtime.importProgress!.read()?.pages[1]?.slideId).toBe('inserted-2')
    expect(binding.readReceipt(parentKey)?.checkpoint?.artifactDigest).toBe(
      receipt.checkpoint!.artifactDigest,
    )
    for (const [i, id] of hostIds.entries())
      if (i !== 1) expect(host.get(id)).toBe(pages[i].pptxBase64)
  } finally {
    runtime?.dispose()
    vi.unstubAllGlobals()
    rmSync(root, { recursive: true, force: true })
  }
})
