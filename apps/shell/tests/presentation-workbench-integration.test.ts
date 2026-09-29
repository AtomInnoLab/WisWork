import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it, vi } from 'vitest'
import { parsePresentationDeck } from '@wiswork/pptx-engine/presentation'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { PresentationStore } from '@wiswork/project-store'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'
import { createOfficeHostRuntime } from '../../office-addin/src/agent/host-runtime'

it('reopens a paused eight-page workbench, downloads saved progress, retries only failure and prepares import', async () => {
  const root = mkdtempSync(join(tmpdir(), 'wiswork-workbench-'))
  const plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  const documentId = 'workbench-doc',
    requestId = 'production-1'
  let releaseFirst!: () => void
  const firstGate = new Promise<void>((resolve) => {
    releaseFirst = resolve
  })
  const attempts = new Map<string, number>()
  const compile: typeof compilePresentationDeck = vi.fn(async (value) => {
    const id = parsePresentationDeck(value).slides[0]!.id
    const attempt = (attempts.get(id) ?? 0) + 1
    attempts.set(id, attempt)
    if (id === deck.slides[0]!.id && attempt === 1) await firstGate
    if (id === deck.slides[1]!.id && attempt === 1) throw new Error('injected compiler failure')
    return compilePresentationDeck(value)
  })
  let service = createPresentationService({ userDataPath: root, compile })
  const signal = new AbortController().signal
  const call = async (operation: string, extra: Record<string, unknown> = {}) =>
    JSON.parse(
      Buffer.from(
        await service({ operation, documentId, projectId: deck.id, ...extra }, signal),
      ).toString(),
    )
  const writeQa = vi.fn(async () => {}),
    writeReceipt = vi.fn(async () => {})
  const create = () =>
    createOfficeHostRuntime('powerpoint', {
      presentation: {
        available: () => true,
        documentId: async () => documentId,
        lastProject: () => deck.id,
        rememberProject: async () => {},
        request: async (body, s) => {
          if ((body as { operation?: string }).operation === 'production_job_resume')
            resumeOperations.push(body)
          return new Response(Buffer.from(await service(body, s ?? signal)))
        },
        readReceipt: () => undefined,
        writeReceipt,
        readQa: () => undefined,
        writeQa,
      },
    })
  const resumeOperations: unknown[] = []
  let runtime = create()
  const status = () => call('production_job_status', { requestId })
  try {
    expect(await call('save_plan', { expectedRevision: 0, plan })).toMatchObject({ revision: 1 })
    expect(await call('production_begin', { requestId, planRevision: 1, deck })).toMatchObject({
      requestId,
    })
    await runtime.presentation!.refresh()
    const disconnected = new AbortController()
    const started = await runtime.skill.executeTool(
      {
        id: 'start',
        name: 'start_presentation_production_job',
        input: { project_id: deck.id, request_id: requestId },
      },
      disconnected.signal,
    )
    expect(started.isError, started.output).not.toBe(true)
    await vi.waitFor(() => expect(attempts.get(deck.slides[0]!.id)).toBe(1))
    await call('production_begin', { requestId: 'newer-production', planRevision: 1, deck })
    await runtime.presentation!.refresh()
    expect(runtime.presentation!.snapshot().project?.production?.requestId).toBe(requestId)
    expect(
      runtime.presentation!.snapshot().project?.productionTasks?.map((task) => task.requestId),
    ).toEqual(['newer-production', requestId])
    await runtime.presentation!.pauseProductionJob(requestId)
    expect((await status()).job.state).toBe('pausing')
    disconnected.abort()
    runtime.dispose()
    releaseFirst()
    await vi.waitFor(async () => expect((await status()).job.state).toBe('paused'), {
      timeout: 10000,
    })
    // A new Taskpane/controller and service instance share saved progress, not chat memory.
    service = createPresentationService({ userDataPath: root, compile })
    runtime = create()
    await runtime.presentation!.refresh()
    expect(runtime.presentation!.snapshot().project?.production?.requestId).toBe('newer-production')
    await runtime.presentation!.selectProduction(requestId)
    expect(runtime.presentation!.snapshot().project).toMatchObject({
      production: { compiledCount: 1 },
      productionJob: { state: 'paused' },
    })
    await runtime.presentation!.downloadProductionPage(deck.slides[0]!.id)
    expect(runtime.presentation!.snapshot().error).toBeUndefined()
    expect(runtime.vfs.list('/home/user').some((path) => path.endsWith('.pptx'))).toBe(true)
    await runtime.presentation!.resumeProductionJob(requestId)
    await vi.waitFor(async () => expect((await status()).job.state).toBe('failed'), {
      timeout: 15000,
    })
    const failed = await status()
    expect(failed.production.compiledCount).toBe(7)
    expect(failed.job.events).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ type: 'run.paused' }),
        expect.objectContaining({ type: 'page.failed', pageId: deck.slides[1]!.id }),
      ]),
    )
    // Match the real button precondition: automatic polling disables actions until idle.
    await vi.waitFor(() => expect(runtime.presentation!.snapshot().phase).toBe('idle'))
    await runtime.presentation!.refresh()
    await runtime.presentation!.resumeProductionJob(requestId)
    expect(resumeOperations.length, JSON.stringify(runtime.presentation!.snapshot())).toBe(2)
    await vi.waitFor(
      async () => {
        const current = await status()
        expect(current.job.state, JSON.stringify(current)).toBe('completed')
      },
      {
        timeout: 15000,
      },
    )
    await runtime.presentation!.refresh()
    for (const [index, page] of deck.slides.entries())
      expect(attempts.get(page.id)).toBe(index === 1 ? 2 : 1)
    await runtime.presentation!.prepareProduction()
    expect(runtime.presentation!.snapshot().error).toBeUndefined()
    expect(runtime.importProgress?.read()).toMatchObject({
      source: 'production',
      total: 8,
      completed: 0,
      status: 'not_started',
    })
    const persisted = new PresentationStore(root).productionJob(deck.id, documentId, requestId)!
    expect(persisted.state).toBe('completed')
    expect(persisted.events.at(-1)?.type).toBe('run.completed')
    expect(writeQa).not.toHaveBeenCalled()
    expect(writeReceipt).not.toHaveBeenCalled()
  } finally {
    releaseFirst()
    runtime.dispose()
    await call('production_job_cancel', { requestId })
    await vi.waitFor(
      async () =>
        expect(['completed', 'cancelled', 'failed', 'paused', 'interrupted', undefined]).toContain(
          (await status()).job?.state,
        ),
      { timeout: 10000 },
    )
    rmSync(root, { recursive: true, force: true })
  }
}, 45000)
