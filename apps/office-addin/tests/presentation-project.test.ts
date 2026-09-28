import { describe, expect, it, vi } from 'vitest'
import { createPresentationProjectController } from '../src/skills/powerpoint/presentation-project.js'

const project = {
  projectId: 'project-1',
  title: '研究汇报',
  status: 'pending',
  latestRequestId: 'request-2',
  latestCompiledRequestId: 'request-1',
  slideCount: 1,
  slides: [{ id: 'slide-1', title: '研究结论' }],
  history: [
    { requestId: 'request-2', sequence: 2, status: 'pending', slideCount: 1 },
    { requestId: 'request-1', sequence: 1, status: 'compiled', slideCount: 1 },
  ],
}
function fixture() {
  const request = vi.fn(
    async (_body: unknown, _signal?: AbortSignal) => new Response(JSON.stringify(project)),
  )
  const executeTool = vi.fn(async () => ({ output: '{}', mutated: false, summary: '已恢复' }))
  const documentId = vi.fn(async () => 'document-1')
  const available = vi.fn(() => true)
  const pdfAvailable = vi.fn(() => true)
  const productionPdfAvailable = vi.fn(() => true)
  const lastProject = vi.fn((): string | undefined => 'project-1')
  let saved: { projectId: string; documentId: string; requestId: string } | undefined
  const selectedProduction = vi.fn((projectId: string, boundDocumentId: string) =>
    saved?.projectId === projectId && saved.documentId === boundDocumentId
      ? saved.requestId
      : undefined,
  )
  const rememberSelectedProduction = vi.fn(
    async (projectId: string, boundDocumentId: string, requestId: string) => {
      saved = { projectId, documentId: boundDocumentId, requestId }
    },
  )
  const createController = () =>
    createPresentationProjectController({
      request,
      executeTool,
      documentId,
      available,
      pdfAvailable,
      productionPdfAvailable,
      lastProject,
      selectedProduction,
      rememberSelectedProduction,
    })
  const controller = createController()
  return {
    controller,
    createController,
    request,
    executeTool,
    documentId,
    available,
    pdfAvailable,
    productionPdfAvailable,
    lastProject,
    selectedProduction,
    rememberSelectedProduction,
  }
}
describe('presentation project controls', () => {
  it('exports only a completed selected production task as a PDF preview', async () => {
    const f = fixture()
    const production = {
      projectId: 'project-1',
      requestId: 'pages',
      planRevision: 1,
      status: 'compiled',
      compiledCount: 1,
      total: 1,
      pages: [{ id: 'slide-1', title: '研究结论', state: 'compiled', attempt: 1 }],
    }
    f.request.mockImplementation(
      async (body) =>
        new Response(
          JSON.stringify(
            (body as { operation: string }).operation === 'status'
              ? { ...project, production }
              : { error: 'invalid_request' },
          ),
        ),
    )
    await f.controller.refresh()
    expect(f.controller.snapshot().error).toBeUndefined()
    expect(f.controller.snapshot().project?.production).toMatchObject({ status: 'compiled' })
    await f.controller.exportProductionPdf?.()
    expect(f.executeTool).toHaveBeenCalledWith(
      expect.objectContaining({
        name: 'export_presentation_pdf',
        input: { project_id: 'project-1', request_id: 'pages', source: 'production' },
      }),
      expect.any(AbortSignal),
    )
    expect(f.controller.snapshot().deliveryNotice).toContain('PDF 预览已保存')
    f.executeTool.mockClear()
    f.productionPdfAvailable.mockReturnValue(false)
    await f.controller.exportProductionPdf?.()
    expect(f.executeTool).not.toHaveBeenCalled()
  })

  it('keeps bounded review comment summaries separate from QA status', async () => {
    const f = fixture()
    f.request.mockResolvedValue(
      new Response(
        JSON.stringify({
          ...project,
          reviewComments: {
            revision: 2,
            openCount: 1,
            resolvedCount: 1,
            recent: [
              {
                id: 'comment-1',
                targetKind: 'source',
                targetId: 'source-1',
                authorLabel: '审阅人甲',
                text: '请复核来源',
                state: 'open',
                planRevision: 1,
                createdAt: '2026-09-28T00:00:00.000Z',
              },
            ],
          },
        }),
      ),
    )
    await f.controller.refresh()
    expect(f.controller.snapshot().project?.reviewComments).toMatchObject({
      openCount: 1,
      recent: [{ id: 'comment-1' }],
    })
    expect(f.controller.snapshot().project?.checks).toBeUndefined()
  })
  it('loads persisted page inventory/history and restores or resumes through the generation skill', async () => {
    const f = fixture()
    await f.controller.refresh()
    expect(f.controller.snapshot()).toMatchObject({ phase: 'idle', project })
    await f.controller.restore()
    expect(f.executeTool).toHaveBeenCalledWith(
      expect.objectContaining({
        name: 'restore_presentation_project',
        input: { project_id: 'project-1' },
      }),
      expect.any(AbortSignal),
    )
    await f.controller.resume('request-2')
    expect(f.executeTool).toHaveBeenCalledWith(
      expect.objectContaining({
        name: 'resume_presentation_project',
        input: { project_id: 'project-1', request_id: 'request-2' },
      }),
      expect.any(AbortSignal),
    )
  })
  it('does not create a document identity or call the PC when no project is known', async () => {
    const f = fixture()
    f.lastProject.mockReturnValue(undefined)
    await f.controller.refresh()
    expect(f.controller.snapshot()).toEqual({ phase: 'idle' })
    expect(f.request).not.toHaveBeenCalled()
    expect(f.documentId).not.toHaveBeenCalled()
  })
  it('does not expose a previous project while recovering a different saved project', async () => {
    const f = fixture()
    await f.controller.refresh()
    f.lastProject.mockReturnValue('project-2')
    let resolve!: (response: Response) => void
    f.request.mockImplementationOnce(
      () =>
        new Promise((done) => {
          resolve = done
        }),
    )
    const pending = f.controller.refresh()
    expect(f.controller.snapshot()).toEqual({ phase: 'loading' })
    await vi.waitFor(() => expect(f.request).toHaveBeenCalledTimes(2))
    resolve(new Response(JSON.stringify({ error: 'unavailable' })))
    await pending
    expect(f.controller.snapshot().project).toBeUndefined()
    expect(f.controller.snapshot().error).toBeTruthy()
  })
  it.each(['clear', 'cancel'] as const)(
    'invalidates a late status response after %s',
    async (action) => {
      const f = fixture()
      let done!: (response: Response) => void
      f.request.mockImplementation(
        () =>
          new Promise((resolve) => {
            done = resolve
          }),
      )
      const pending = f.controller.refresh()
      await vi.waitFor(() => expect(f.request).toHaveBeenCalled())
      const signal = f.request.mock.calls[0]![1]!
      f.controller[action]()
      expect(signal.aborted).toBe(true)
      done(new Response(JSON.stringify(project)))
      await pending
      expect(f.controller.snapshot().project).toBeUndefined()
      expect(f.controller.snapshot().phase).toBe('idle')
    },
  )
  it('rejects foreign, malformed, and falsely reviewed project status', async () => {
    for (const value of [
      { ...project, projectId: 'other' },
      { ...project, slideCount: 2 },
      { ...project, checks: { render: 'passed' } },
      { ...project, reviewComments: null },
    ]) {
      const f = fixture()
      f.request.mockResolvedValue(new Response(JSON.stringify(value)))
      await f.controller.refresh()
      expect(f.controller.snapshot().project).toBeUndefined()
      expect(f.controller.snapshot().error).toBeTruthy()
    }
  })
  it('does not publish a status after switching documents or losing capability', async () => {
    const f = fixture()
    f.documentId.mockResolvedValueOnce('document-1').mockResolvedValue('document-2')
    await f.controller.refresh()
    expect(f.controller.snapshot().project).toBeUndefined()
    expect(f.controller.snapshot().error).toContain('文档')
    const g = fixture()
    g.request.mockImplementation(async () => {
      g.available.mockReturnValue(false)
      return new Response(JSON.stringify(project))
    })
    await g.controller.refresh()
    expect(g.controller.snapshot().project).toBeUndefined()
  })
  it('explains old-PC compatibility without exposing a raw code', async () => {
    const f = fixture()
    f.request.mockResolvedValue(new Response(JSON.stringify({ error: 'invalid_request' })))
    await f.controller.refresh()
    expect(f.controller.snapshot().error).toContain('升级')
    expect(f.controller.snapshot().error).not.toContain('invalid_request')
  })
  it('does not duplicate an active action and aborts it on clear', async () => {
    const f = fixture()
    await f.controller.refresh()
    let done!: () => void
    f.executeTool.mockImplementation(async () => {
      await new Promise<void>((resolve) => {
        done = resolve
      })
      return { output: '{}', mutated: false, summary: '已恢复' }
    })
    const pending = f.controller.resume('request-2')
    await vi.waitFor(() => expect(f.executeTool).toHaveBeenCalledOnce())
    await f.controller.resume('request-2')
    expect(f.executeTool).toHaveBeenCalledOnce()
    f.controller.clear()
    done()
    await pending
    expect(f.controller.snapshot()).toEqual({ phase: 'idle' })
  })
  it('rejects an arbitrary stale request instead of resuming a different revision', async () => {
    const f = fixture()
    await f.controller.refresh()
    await f.controller.resume('other-request')
    expect(f.executeTool).not.toHaveBeenCalled()
  })
})

describe('plan-only project status', () => {
  it('runs a document-bound literal source audit and drops it after plan refresh', async () => {
    const { benchmarkPlan } =
      await import('../../../packages/pptx-engine/tests/fixtures/presentation-plan.js')
    const plan = benchmarkPlan()
    const attachmentId = 'a'.repeat(64)
    plan.sources[0]!.uri = `attachment:${attachmentId}`
    const status = {
      projectId: plan.projectId,
      title: plan.title,
      status: 'planned',
      slideCount: plan.slides.length,
      slides: plan.slides.map(({ id, title }) => ({ id, title })),
      history: [],
      plan: { revision: 1, value: plan },
    }
    const f = fixture()
    f.lastProject.mockReturnValue(plan.projectId)
    f.request.mockResolvedValue(new Response(JSON.stringify(status)))
    await f.controller.refresh()
    f.executeTool.mockResolvedValue({
      output: JSON.stringify({
        projectId: plan.projectId,
        planRevision: 1,
        sources: [{ sourceId: plan.sources[0]!.id, attachmentId, status: 'found', offset: 5 }],
        checks: {
          support: 'not_verified',
          sourceAuthority: 'not_verified',
          timeliness: 'not_verified',
        },
      }),
      mutated: false,
      summary: '已核对',
    })
    await f.controller.auditSources()
    expect(f.executeTool).toHaveBeenCalledWith(
      expect.objectContaining({
        name: 'audit_presentation_sources',
        input: { project_id: plan.projectId },
      }),
      expect.any(AbortSignal),
    )
    expect(f.controller.snapshot().sourceAudit?.sources[0]).toMatchObject({
      status: 'found',
      offset: 5,
    })
    await f.controller.refresh()
    expect(f.controller.snapshot().sourceAudit).toBeUndefined()
    f.executeTool.mockResolvedValue({
      output: JSON.stringify({
        projectId: plan.projectId,
        planRevision: 0,
        sources: [{ sourceId: plan.sources[0]!.id, attachmentId, status: 'found', offset: 5 }],
        checks: {
          support: 'not_verified',
          sourceAuthority: 'not_verified',
          timeliness: 'not_verified',
        },
      }),
      mutated: false,
      summary: '已核对',
    })
    await f.controller.auditSources()
    expect(f.controller.snapshot().sourceAudit).toBeUndefined()
    expect(f.controller.snapshot().error).toBeTruthy()
  })
  it('accepts document source preparation and rejects forged attachment mappings', async () => {
    const { benchmarkPlan } =
      await import('../../../packages/pptx-engine/tests/fixtures/presentation-plan.js')
    const attachmentId = 'a'.repeat(64)
    const plan = benchmarkPlan()
    plan.sources[0]!.uri = `attachment:${attachmentId}`
    const value = {
      projectId: plan.projectId,
      title: plan.title,
      status: 'planned',
      slideCount: plan.slides.length,
      slides: plan.slides.map(({ id, title }) => ({ id, title })),
      history: [],
      plan: { revision: 1, value: plan },
      sourcePreparation: [{ sourceId: plan.sources[0]!.id, attachmentId, status: 'ready' }],
    }
    const f = fixture()
    f.lastProject.mockReturnValue(plan.projectId)
    f.request.mockResolvedValueOnce(new Response(JSON.stringify(value)))
    await f.controller.refresh()
    expect(f.controller.snapshot().project?.sourcePreparation).toEqual(value.sourcePreparation)
    f.request.mockResolvedValueOnce(
      new Response(
        JSON.stringify({
          ...value,
          sourcePreparation: [{ ...value.sourcePreparation[0], attachmentId: 'b'.repeat(64) }],
        }),
      ),
    )
    await f.controller.refresh()
    expect(f.controller.snapshot().error).toBeTruthy()
  })
  it('accepts a saved plan before compilation without inventing request history or QA', async () => {
    const { benchmarkPlan } =
      await import('../../../packages/pptx-engine/tests/fixtures/presentation-plan.js')
    const plan = benchmarkPlan()
    const value = {
      projectId: plan.projectId,
      title: plan.title,
      status: 'planned',
      slideCount: plan.slides.length,
      slides: plan.slides.map(({ id, title }) => ({ id, title })),
      history: [],
      plan: { revision: 1, value: plan },
    }
    const f = fixture()
    f.lastProject.mockReturnValue(plan.projectId)
    f.request.mockResolvedValue(new Response(JSON.stringify(value)))
    await f.controller.refresh()
    expect(f.controller.snapshot().error).toBeUndefined()
    expect(f.controller.snapshot().project).toEqual(value)
    await f.controller.resume('made-up')
    expect(f.executeTool).not.toHaveBeenCalled()
  })
  it('accepts bounded saved plan revisions and rejects malformed replay history', async () => {
    const { benchmarkPlan } =
      await import('../../../packages/pptx-engine/tests/fixtures/presentation-plan.js')
    const plan = benchmarkPlan()
    const snapshot = {
      sourceCount: plan.sources.length,
      claimCount: plan.claims.length,
      slideCount: plan.slides.length,
      sourcesDigest: 'a'.repeat(64),
      claimsDigest: 'b'.repeat(64),
      slidesDigest: 'c'.repeat(64),
      styleDigest: 'd'.repeat(64),
    }
    const event = {
      revision: 1,
      inputDigest: 'a'.repeat(64),
      createdAt: '2026-09-24T00:00:00.000Z',
      snapshot,
    }
    const value = {
      projectId: plan.projectId,
      title: plan.title,
      status: 'planned',
      slideCount: plan.slides.length,
      slides: plan.slides.map(({ id, title }) => ({ id, title })),
      history: [],
      plan: { revision: 1, value: plan, revisions: [event] },
    }
    const f = fixture()
    f.lastProject.mockReturnValue(plan.projectId)
    f.request.mockResolvedValueOnce(new Response(JSON.stringify(value)))
    await f.controller.refresh()
    expect(f.controller.snapshot().project?.plan?.revisions).toEqual([event])
    f.request.mockResolvedValueOnce(
      new Response(
        JSON.stringify({
          ...value,
          plan: { ...value.plan, revisions: [{ ...event, revision: 2 }] },
        }),
      ),
    )
    await f.controller.refresh()
    expect(f.controller.snapshot().error).toBeTruthy()
    f.request.mockResolvedValueOnce(
      new Response(
        JSON.stringify({
          ...value,
          plan: {
            ...value.plan,
            revisions: [{ ...event, snapshot: { ...snapshot, sourceCount: 999 } }],
          },
        }),
      ),
    )
    await f.controller.refresh()
    expect(f.controller.snapshot().error).toBeTruthy()
  })
})
it('projects page production and resumes its explicit request through the page tool', async () => {
  const f = fixture(),
    production = {
      projectId: 'project-1',
      requestId: 'pages-1',
      planRevision: 1,
      status: 'partial',
      compiledCount: 1,
      total: 2,
      pages: [
        { id: 'a', title: 'A', state: 'compiled', attempt: 1 },
        { id: 'b', title: 'B', state: 'failed', attempt: 1, error: 'compile_failed' },
      ],
    }
  f.request.mockImplementation(
    async (body) =>
      new Response(
        JSON.stringify(
          (body as { operation: string }).operation === 'production_job_status'
            ? { error: 'invalid_request' }
            : { ...project, production },
        ),
      ),
  )
  await f.controller.refresh()
  expect(f.controller.snapshot().project?.production).toEqual(production)
  await f.controller.runProduction('wrong')
  expect(f.executeTool).not.toHaveBeenCalled()
  await f.controller.runProduction('pages-1')
  expect(f.executeTool).toHaveBeenCalledWith(
    expect.objectContaining({
      name: 'run_presentation_production',
      input: { project_id: 'project-1', request_id: 'pages-1' },
    }),
    expect.any(AbortSignal),
  )
})

it('polls active background work and clears polling without cancelling PC work', async () => {
  vi.useFakeTimers()
  try {
    const f = fixture()
    const production = {
      projectId: 'project-1',
      requestId: 'pages',
      planRevision: 1,
      status: 'pending',
      compiledCount: 0,
      total: 1,
      pages: [{ id: 'a', title: 'A', state: 'pending', attempt: 0 }],
    }
    const job = {
      version: 1,
      projectId: 'project-1',
      documentId: 'document-1',
      requestId: 'pages',
      inputDigest: 'a'.repeat(64),
      planDigest: 'b'.repeat(64),
      planRevision: 1,
      revision: 1,
      state: 'running',
      events: [{ sequence: 1, createdAt: '2026-09-24T00:00:00.000Z', type: 'run.started' }],
    }
    f.request.mockImplementation(
      async (body) =>
        new Response(
          JSON.stringify(
            (body as { operation: string }).operation === 'status'
              ? { ...project, production }
              : {
                  job,
                  production: {
                    ...production,
                    inputDigest: job.inputDigest,
                    planDigest: job.planDigest,
                  },
                },
          ),
        ),
    )
    await f.controller.refresh()
    expect(f.controller.snapshot().project?.productionJob?.state).toBe('running')
    await vi.advanceTimersByTimeAsync(1500)
    expect(f.request).toHaveBeenCalledTimes(4)
    f.controller.cancel()
    await vi.advanceTimersByTimeAsync(5000)
    expect(f.request).toHaveBeenCalledTimes(4)
    expect(f.executeTool).not.toHaveBeenCalled()
  } finally {
    vi.useRealTimers()
  }
})
it('automatically resumes only an interrupted frozen page job after reconnect', async () => {
  const f = fixture()
  const production = {
    projectId: 'project-1',
    requestId: 'pages',
    planRevision: 1,
    status: 'pending',
    compiledCount: 0,
    total: 1,
    pages: [{ id: 'a', title: 'A', state: 'pending', attempt: 0 }],
  }
  let state: 'interrupted' | 'running' = 'interrupted'
  const job = () => ({
    version: 1,
    projectId: 'project-1',
    documentId: 'document-1',
    requestId: 'pages',
    inputDigest: 'a'.repeat(64),
    planDigest: 'b'.repeat(64),
    planRevision: 1,
    revision: state === 'running' ? 3 : 2,
    state,
    events: [
      { sequence: 1, createdAt: '2026-09-24T00:00:00.000Z', type: 'run.started' },
      { sequence: 2, createdAt: '2026-09-24T00:00:01.000Z', type: 'run.interrupted' },
      ...(state === 'running'
        ? [{ sequence: 3, createdAt: '2026-09-24T00:00:02.000Z', type: 'run.started' }]
        : []),
    ],
  })
  f.request.mockImplementation(
    async (body) =>
      new Response(
        JSON.stringify(
          (body as { operation: string }).operation === 'status'
            ? { ...project, production }
            : {
                job: job(),
                production: {
                  ...production,
                  inputDigest: 'a'.repeat(64),
                  planDigest: 'b'.repeat(64),
                },
              },
        ),
      ),
  )
  let attempts = 0
  f.executeTool.mockImplementation(async () => {
    attempts++
    if (attempts === 1)
      return {
        output: 'presentation_service_unavailable',
        isError: true,
        mutated: false,
        summary: '连接中断',
      }
    state = 'running'
    return { output: '{}', mutated: false, summary: '已继续' }
  })
  await f.controller.refresh()
  expect(f.executeTool).not.toHaveBeenCalled()
  f.controller.prepareReconnect()
  await f.controller.refresh()
  expect(f.controller.snapshot().error).toBeTruthy()
  await f.controller.refresh()
  expect(f.executeTool).toHaveBeenCalledTimes(2)
  expect(f.executeTool).toHaveBeenCalledWith(
    expect.objectContaining({
      name: 'resume_presentation_production_job',
      input: { project_id: 'project-1', request_id: 'pages' },
    }),
    expect.any(AbortSignal),
  )
  expect(f.controller.snapshot().project?.productionJob?.state).toBe('running')
  f.controller.cancel()
})
it('does not automatically resume a paused job or one in a different document', async () => {
  const f = fixture()
  const production = {
    projectId: 'project-1',
    requestId: 'pages',
    planRevision: 1,
    status: 'pending',
    compiledCount: 0,
    total: 1,
    pages: [{ id: 'a', title: 'A', state: 'pending', attempt: 0 }],
  }
  const job = {
    version: 1,
    projectId: 'project-1',
    documentId: 'document-1',
    requestId: 'pages',
    inputDigest: 'a'.repeat(64),
    planDigest: 'b'.repeat(64),
    planRevision: 1,
    revision: 2,
    state: 'paused',
    events: [
      { sequence: 1, createdAt: '2026-09-24T00:00:00.000Z', type: 'run.started' },
      { sequence: 2, createdAt: '2026-09-24T00:00:01.000Z', type: 'run.paused' },
    ],
  }
  f.request.mockImplementation(
    async (body) =>
      new Response(
        JSON.stringify(
          (body as { operation: string }).operation === 'status'
            ? { ...project, production }
            : {
                job,
                production: {
                  ...production,
                  inputDigest: job.inputDigest,
                  planDigest: job.planDigest,
                },
              },
        ),
      ),
  )
  f.controller.prepareReconnect()
  await f.controller.refresh()
  expect(f.executeTool).not.toHaveBeenCalled()
  f.controller.clear()
  f.controller.prepareReconnect()
  job.state = 'interrupted'
  job.events[1]!.type = 'run.interrupted'
  f.documentId.mockResolvedValueOnce('document-1').mockResolvedValue('document-2')
  await f.controller.refresh()
  expect(f.executeTool).not.toHaveBeenCalled()
})
it('downloads only completed pages and prepares only a fully compiled production', async () => {
  const f = fixture()
  const production = {
    projectId: 'project-1',
    requestId: 'pages',
    planRevision: 1,
    status: 'partial',
    compiledCount: 1,
    total: 2,
    pages: [
      { id: 'a', title: 'A', state: 'compiled', attempt: 1 },
      { id: 'b', title: 'B', state: 'pending', attempt: 0 },
    ],
  }
  f.request.mockImplementation(
    async (body) =>
      new Response(
        JSON.stringify(
          (body as { operation: string }).operation === 'status'
            ? { ...project, production }
            : { error: 'invalid_request' },
        ),
      ),
  )
  await f.controller.refresh()
  await f.controller.downloadProductionPage('b')
  await f.controller.prepareProduction()
  expect(f.executeTool).not.toHaveBeenCalled()
  await f.controller.downloadProductionPage('a')
  expect(f.executeTool).toHaveBeenCalledWith(
    expect.objectContaining({
      name: 'read_presentation_page_artifact',
      input: { project_id: 'project-1', request_id: 'pages', page_id: 'a' },
    }),
    expect.any(AbortSignal),
  )
})
it('retains the loaded project on transient same-document failure but clears after a document switch', async () => {
  const f = fixture()
  await f.controller.refresh()
  f.request.mockImplementation(async () => new Response('', { status: 503 }))
  await f.controller.refresh()
  expect(f.controller.snapshot().project).toEqual(project)
  expect(f.controller.snapshot().error).toBeTruthy()
  f.documentId.mockResolvedValue('other')
  await f.controller.refresh()
  expect(f.controller.snapshot().project).toBeUndefined()
})

describe('saved production task selection', () => {
  function tasksFixture() {
    const f = fixture()
    const production = {
      projectId: 'project-1',
      requestId: 'new',
      planRevision: 1,
      status: 'pending',
      compiledCount: 0,
      total: 1,
      pages: [{ id: 'a', title: 'A', state: 'pending', attempt: 0 }],
    }
    const productionTasks = [
      {
        requestId: 'new',
        sequence: 2,
        planRevision: 1,
        status: 'pending',
        compiledCount: 0,
        total: 1,
      },
      {
        requestId: 'old',
        sequence: 1,
        planRevision: 1,
        status: 'pending',
        compiledCount: 0,
        total: 1,
        jobState: 'paused',
      },
    ]
    f.request.mockImplementation(async (body) => {
      const request = body as { operation: string; requestId?: string; documentId: string }
      if (request.operation === 'status')
        return new Response(JSON.stringify({ ...project, production, productionTasks }))
      const job =
        request.requestId === 'old'
          ? {
              version: 1,
              projectId: 'project-1',
              documentId: request.documentId,
              requestId: 'old',
              inputDigest: 'a'.repeat(64),
              planDigest: 'b'.repeat(64),
              planRevision: 1,
              revision: 3,
              state: 'paused',
              events: [
                { sequence: 1, createdAt: '2026-09-24T00:00:00.000Z', type: 'run.started' },
                { sequence: 2, createdAt: '2026-09-24T00:00:01.000Z', type: 'run.pause_requested' },
                { sequence: 3, createdAt: '2026-09-24T00:00:02.000Z', type: 'run.paused' },
              ],
            }
          : null
      return new Response(
        JSON.stringify({
          job,
          production: {
            ...production,
            requestId: request.requestId,
            inputDigest: 'a'.repeat(64),
            planDigest: 'b'.repeat(64),
          },
        }),
      )
    })
    return { ...f, productionTasks, production }
  }
  it('selects older paused work and retains explicit selection across refresh', async () => {
    const f = tasksFixture()
    await f.controller.refresh()
    await f.controller.selectProduction('old')
    expect(f.controller.snapshot().project?.production?.requestId).toBe('old')
    expect(f.controller.snapshot().project?.productionJob?.state).toBe('paused')
    await f.controller.refresh()
    expect(f.controller.snapshot().project?.production?.requestId).toBe('old')
    await f.controller.resumeProductionJob('old')
    expect(f.executeTool).toHaveBeenCalledWith(
      expect.objectContaining({
        name: 'resume_presentation_production_job',
        input: { project_id: 'project-1', request_id: 'old' },
      }),
      expect.any(AbortSignal),
    )
    const calls = f.request.mock.calls.length
    await f.controller.selectProduction('invented')
    expect(f.request).toHaveBeenCalledTimes(calls)
  })
  it('restores an explicit older task after controller recreation and ignores stale or copied selection', async () => {
    const f = tasksFixture()
    await f.controller.refresh()
    await f.controller.selectProduction('old')
    expect(f.rememberSelectedProduction).toHaveBeenCalledWith('project-1', 'document-1', 'old')
    const reopened = f.createController()
    await reopened.refresh()
    expect(reopened.snapshot().project?.production?.requestId).toBe('old')
    f.documentId.mockResolvedValue('copied-document')
    const copy = f.createController()
    await copy.refresh()
    expect(copy.snapshot().project?.production?.requestId).toBe('new')
    f.documentId.mockResolvedValue('document-1')
    f.productionTasks.splice(1, 1)
    const removed = f.createController()
    await removed.refresh()
    expect(removed.snapshot().project?.production?.requestId).toBe('new')
  })
  it('keeps the previous task when saving a new task choice fails', async () => {
    const f = tasksFixture()
    await f.controller.refresh()
    f.rememberSelectedProduction.mockRejectedValueOnce(new Error('save_failed'))
    await f.controller.selectProduction('old')
    expect(f.controller.snapshot().project?.production?.requestId).toBe('new')
    expect(f.controller.snapshot().error).toContain('无法保存所选页任务')
  })
  it('keeps the document task choice after cancelling a wait but resets this panel after new conversation', async () => {
    const f = tasksFixture()
    await f.controller.refresh()
    await f.controller.selectProduction('old')
    f.controller.cancel()
    await f.controller.refresh()
    expect(f.controller.snapshot().project?.production?.requestId).toBe('old')
    f.controller.clear()
    await f.controller.refresh()
    expect(f.controller.snapshot().project?.production?.requestId).toBe('new')
  })
  it('restores the persisted task choice after a bridge reconnection', async () => {
    const f = tasksFixture()
    await f.controller.refresh()
    await f.controller.selectProduction('old')
    f.controller.clear()
    f.controller.prepareReconnect()
    await f.controller.refresh()
    expect(f.controller.snapshot().project?.production?.requestId).toBe('old')
  })
  it.each(['clear', 'document'] as const)('resets selected task after %s', async (action) => {
    const f = tasksFixture()
    await f.controller.refresh()
    await f.controller.selectProduction('old')
    if (action === 'clear') f.controller.clear()
    else f.documentId.mockResolvedValue('other-document')
    await f.controller.refresh()
    expect(f.controller.snapshot().project?.production?.requestId).toBe('new')
  })
  it('rejects invalid or duplicated task identities and unordered sequences', async () => {
    for (const invalid of [
      [
        {
          requestId: '../bad',
          sequence: 1,
          planRevision: 1,
          status: 'pending',
          compiledCount: 0,
          total: 1,
        },
      ],
      [
        {
          requestId: 'a',
          sequence: 1,
          planRevision: 1,
          status: 'compiled',
          compiledCount: 0,
          total: 1,
        },
      ],
      [
        {
          requestId: 'a',
          sequence: 1,
          planRevision: 1,
          status: 'pending',
          compiledCount: 0,
          total: 1,
          jobState: 'unknown',
        },
      ],
    ]) {
      const f = fixture()
      f.request.mockResolvedValue(
        new Response(JSON.stringify({ ...project, productionTasks: invalid })),
      )
      await f.controller.refresh()
      expect(f.controller.snapshot().project).toBeUndefined()
    }
  })
  it('accepts bounded older job events and rejects malformed event summaries', async () => {
    const f = tasksFixture()
    Object.assign(f.productionTasks[1]!, {
      lastEvent: {
        type: 'run.paused',
        createdAt: '2026-09-24T00:00:02.000Z',
      },
    })
    await f.controller.refresh()
    expect(f.controller.snapshot().project?.productionTasks?.[1]?.lastEvent?.type).toBe(
      'run.paused',
    )
    for (const lastEvent of [
      null,
      { type: 'unknown', createdAt: '2026-09-24T00:00:02.000Z' },
      { type: 'page.failed', createdAt: '2026-09-24T00:00:02.000Z' },
      { type: 'run.paused', createdAt: 'bad' },
      { type: 'run.paused', createdAt: '2026-09-24T00:00:02.000Z', pageId: 'wrong' },
    ]) {
      const bad = tasksFixture()
      Object.assign(bad.productionTasks[1]!, { lastEvent })
      await bad.controller.refresh()
      expect(bad.controller.snapshot().project).toBeUndefined()
    }
  })
  it('does not prepare derived revisions from the workbench even when every page compiled', async () => {
    const f = tasksFixture()
    f.production.status = 'compiled'
    f.production.compiledCount = 1
    f.production.pages[0]!.state = 'compiled'
    f.production.pages[0]!.attempt = 1
    Object.assign(f.production, {
      revision: { parentRequestId: 'parent', pageId: 'a', parentInputDigest: 'c'.repeat(64) },
    })
    f.productionTasks[0]!.status = 'compiled'
    f.productionTasks[0]!.compiledCount = 1
    await f.controller.refresh()
    await f.controller.prepareProduction()
    expect(f.executeTool).not.toHaveBeenCalled()
  })
})

it('binds report actions to the selected request, CAS revision, and clears late responses', async () => {
  const { deliveryReportFixture } = await import('./presentation-delivery-fixture.js')
  const report = await deliveryReportFixture()
  const documentId = vi.fn(async () => 'd')
  const production = {
    projectId: report.projectId,
    requestId: 'r',
    planRevision: 1,
    status: 'pending',
    compiledCount: 0,
    total: report.pages.length,
    pages: report.pages.map((page) => ({
      id: page.pageId,
      title: page.title,
      state: 'pending',
      attempt: 0,
    })),
  }
  const executeTool = vi.fn(async () => ({
    output: JSON.stringify(report),
    mutated: false,
    summary: 'report',
  }))
  const controller = createPresentationProjectController({
    documentId,
    available: () => true,
    lastProject: () => report.projectId,
    executeTool,
    request: async (body) =>
      new Response(
        JSON.stringify(
          (body as { operation: string }).operation === 'status'
            ? { ...project, projectId: report.projectId, production }
            : { error: 'invalid_request' },
        ),
      ),
  })
  await controller.refresh()
  await controller.readDeliveryReport()
  expect(controller.snapshot().deliveryReport).toEqual(report)
  const issue = report.pages.flatMap((page) => page.issues)[0]!
  const action = {
    actionId: 'act',
    issueId: issue.id,
    issueDigest: issue.digest,
    state: 'deferred' as const,
    note: 'Needs source',
  }
  await controller.recordIssueAction(action)
  expect(executeTool).toHaveBeenLastCalledWith(
    expect.objectContaining({
      name: 'record_presentation_issue_action',
      input: { project_id: report.projectId, request_id: 'r', expected_revision: 0, action },
    }),
    expect.any(AbortSignal),
  )
  let finish!: (value: Awaited<ReturnType<typeof executeTool>>) => void
  executeTool.mockImplementation(
    () =>
      new Promise((resolve) => {
        finish = resolve
      }),
  )
  const pending = controller.readDeliveryReport()
  await vi.waitFor(() => expect(finish).toBeTypeOf('function'))
  controller.clear()
  finish({ output: JSON.stringify(report), mutated: false, summary: 'late' })
  await pending
  expect(controller.snapshot()).toEqual({ phase: 'idle' })
  executeTool.mockResolvedValue({
    output: JSON.stringify(report),
    mutated: false,
    summary: 'report',
  })
  await controller.refresh()
  documentId.mockResolvedValue('different')
  await controller.readDeliveryReport()
  expect(controller.snapshot().deliveryReport).toBeUndefined()
  expect(controller.snapshot().project).toBeUndefined()
})
