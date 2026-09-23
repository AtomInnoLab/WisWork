import { createPresentationDocumentBinding } from '../../office-addin/src/skills/powerpoint/presentation-document'
import { BrowserPresentationPageReplacementAdapter } from '../../office-addin/src/skills/powerpoint/browser-presentation-page-replacement-adapter'
import { createHash } from 'node:crypto'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it, vi } from 'vitest'
import { openPptx } from '@wiswork/pptx-engine'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'
import { createOfficeHostRuntime } from '../../office-addin/src/agent/host-runtime'
import { BrowserPowerPointAdapter } from '../../office-addin/src/skills/powerpoint/browser-powerpoint-adapter'
import { presentationArtifactContent } from '../../office-addin/src/skills/powerpoint/presentation-page-delivery'
import type {
  CompiledPresentationArtifact,
  PresentationImportRecord,
} from '../../office-addin/src/skills/powerpoint/presentation-delivery'

it('backs up, stages and discards a page revision across PC and Taskpane restart while preserving the original', async () => {
  const root = mkdtempSync(join(tmpdir(), 'wiswork-page-backup-integration-'))
  const settings = new Map<string, string>([['wiswork.presentation.document.v1', 'docid']])
  const binding = createPresentationDocumentBinding({
    get: (key) => settings.get(key),
    set: (key, value) => {
      settings.set(key, value)
    },
    save: async () => {},
    location: () => 'file://deck.pptx',
  })
  const documentId = await binding.documentId()
  let staged = false
  const inspectStage = vi
    .spyOn(BrowserPresentationPageReplacementAdapter.prototype, 'inspect')
    .mockImplementation(async (record) => ({
      status: staged ? 'staged' : 'baseline',
      slideIds: staged
        ? [
            ...record.beforeSlideIds.slice(0, record.beforeSlideIds.indexOf(record.oldSlideId) + 1),
            'staged-host',
            ...record.beforeSlideIds.slice(record.beforeSlideIds.indexOf(record.oldSlideId) + 1),
          ]
        : record.beforeSlideIds,
    }))
  const insertStage = vi
    .spyOn(BrowserPresentationPageReplacementAdapter.prototype, 'stage')
    .mockImplementation(async (record, base64, onInserted, assertCurrent) => {
      expect(binding.readPageReplacement()?.state).toBe('pending')
      expect((await openPptx(Buffer.from(base64, 'base64'))).deck.slides).toHaveLength(1)
      await assertCurrent()
      staged = true
      await onInserted('staged-host')
      expect(binding.readPageReplacement()?.state).toBe('inserted')
      expect(record.oldSlideId).toBe('host-1')
    })
  const discardStage = vi
    .spyOn(BrowserPresentationPageReplacementAdapter.prototype, 'discard')
    .mockImplementation(async (_record, assertCurrent) => {
      expect(binding.readPageReplacement()?.state).toBe('discard_pending')
      await assertCurrent()
      staged = false
    })
  let service = createPresentationService({ userDataPath: root })
  const deck = benchmarkPlannedDeck(),
    signal = new AbortController().signal
  const call = async (operation: string, extra = {}) =>
    JSON.parse(
      Buffer.from(
        await service(
          {
            operation,
            documentId,
            projectId: deck.id,
            ...extra,
          },
          signal,
        ),
      ).toString(),
    )
  let runtime: ReturnType<typeof createOfficeHostRuntime> | undefined
  let exported: ReturnType<typeof vi.spyOn> | undefined
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
      slide: { ...target, notes: 'Revision' },
    })
    expect(await call('production_run', { requestId: 'child' })).toMatchObject({
      status: 'compiled',
    })
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
    const hostPackage = await compilePresentationDeck({
      ...deck,
      slides: [{ ...target, notes: 'Manual user correction before backup' }],
    })
    const hostBytes = hostPackage.bytes
    expect(Buffer.from(hostBytes).equals(Buffer.from(pages[1].pptxBase64, 'base64'))).toBe(false)
    exported = vi
      .spyOn(BrowserPowerPointAdapter.prototype, 'exportPresentationPagePackage')
      .mockResolvedValue({
        slideId: hostIds[1]!,
        slideIds: ['existing', ...hostIds],
        base64: Buffer.from(hostBytes).toString('base64'),
      })
    const create = () =>
      createOfficeHostRuntime('powerpoint', {
        presentation: {
          ...binding,
          available: () => true,
          documentId: binding.documentId,
          lastProject: () => deck.id,
          rememberProject: async () => {},
          readReceipt: (key) => (key === `production/${deck.id}/parent` ? receipt : undefined),
          request: async (body, s) => new Response(Buffer.from(await service(body, s ?? signal))),
        },
      })
    runtime = create()
    expect(
      (
        await runtime.skill.executeTool({
          id: 'prepare',
          name: 'prepare_presentation_production_import',
          input: { project_id: deck.id, request_id: 'parent' },
        })
      ).isError,
    ).not.toBe(true)
    const saved = await runtime.skill.executeTool({
      id: 'save',
      name: 'save_presentation_page_backup',
      input: {
        project_id: deck.id,
        request_id: 'child',
        page_id: target.id,
        backup_id: 'backup-1',
      },
    })
    expect(saved.isError, saved.output).not.toBe(true)
    expect(saved.mutated).not.toBe(true)
    expect(exported).toHaveBeenCalledWith(hostIds[1], undefined)
    expect(runtime.proposals.pending()).toBeUndefined()
    const metadata = await call('page_backup_status', { backupId: 'backup-1' })
    expect(metadata).toMatchObject({
      status: 'ready',
      requestId: 'child',
      parentRequestId: 'parent',
      hostSlideId: hostIds[1],
      slideIds: ['existing', ...hostIds],
    })
    const parentReceipt = JSON.stringify(receipt)
    const stage = await runtime.skill.executeTool({
      id: 'stage',
      name: 'stage_presentation_page_replacement',
      input: {
        project_id: deck.id,
        request_id: 'child',
        page_id: target.id,
        backup_id: 'backup-1',
        change_id: 'stage-1',
      },
    })
    expect(stage.isError, stage.output).not.toBe(true)
    expect(staged).toBe(false)
    expect(binding.readPageReplacement()).toBeUndefined()
    await runtime.proposals.confirm(runtime.proposals.pending()!.id)
    expect(binding.readPageReplacement()).toMatchObject({
      state: 'staged',
      newSlideId: 'staged-host',
      oldSlideId: hostIds[1],
    })
    expect(staged).toBe(true)
    expect(JSON.stringify(receipt)).toBe(parentReceipt)
    runtime.dispose()
    service = createPresentationService({ userDataPath: root })
    runtime = create()
    expect(
      (
        await runtime.skill.executeTool({
          id: 'prepare-again',
          name: 'prepare_presentation_production_import',
          input: { project_id: deck.id, request_id: 'parent' },
        })
      ).isError,
    ).not.toBe(true)
    const inspected = await runtime.skill.executeTool({
      id: 'inspect',
      name: 'inspect_presentation_page_replacement',
      input: { project_id: deck.id, change_id: 'stage-1' },
    })
    expect(inspected.isError, inspected.output).not.toBe(true)
    expect(inspected.output).toContain('staged')
    expect(insertStage).toHaveBeenCalledTimes(1)
    const discard = await runtime.skill.executeTool({
      id: 'discard',
      name: 'discard_presentation_page_replacement',
      input: { project_id: deck.id, change_id: 'stage-1' },
    })
    expect(discard.isError, discard.output).not.toBe(true)
    await runtime.proposals.confirm(runtime.proposals.pending()!.id)
    expect(staged).toBe(false)
    expect(discardStage).toHaveBeenCalledTimes(1)
    expect(binding.readPageReplacement()?.state).toBe('discarded')
    expect(JSON.stringify(receipt)).toBe(parentReceipt)
    const downloaded = await runtime.skill.executeTool({
      id: 'read',
      name: 'read_presentation_page_backup',
      input: { project_id: deck.id, backup_id: 'backup-1' },
    })
    expect(downloaded.isError, downloaded.output).not.toBe(true)
    const output = JSON.parse(downloaded.output)
    const path = output.path
    expect(typeof path).toBe('string')
    const bytes = runtime.vfs.readBytes(path)
    expect(Buffer.from(bytes)).toEqual(Buffer.from(hostBytes))
    expect((await openPptx(bytes)).deck.slides).toHaveLength(1)
    expect(await call('page_backup_status', { backupId: 'backup-1', documentId: 'other' })).toEqual(
      { error: 'document_mismatch' },
    )
    expect(await call('production_page', { requestId: 'parent', pageId: target.id })).toEqual(
      pages[1],
    )
    const childPrepare = await runtime.skill.executeTool({
      id: 'child-import',
      name: 'prepare_presentation_production_import',
      input: { project_id: deck.id, request_id: 'child' },
    })
    expect(childPrepare.isError).toBe(true)
    expect(childPrepare.output).toContain('presentation_page_replacement_required')
  } finally {
    runtime?.dispose()
    exported?.mockRestore()
    inspectStage.mockRestore()
    insertStage.mockRestore()
    discardStage.mockRestore()
    rmSync(root, { recursive: true, force: true })
  }
})
