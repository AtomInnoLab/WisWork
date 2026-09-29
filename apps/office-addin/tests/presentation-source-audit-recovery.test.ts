import { createHash } from 'node:crypto'
import { expect, it, vi } from 'vitest'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { canonicalPresentationValue } from '@wiswork/project-store/presentation-canonical'
import { presentationSourceAuditHistory } from '@wiswork/project-store/presentation-source-audit'
import { createPresentationProjectController } from '../src/skills/powerpoint/presentation-project.js'

function fixture() {
  const plan = benchmarkPlan()
  plan.sources[0]!.uri = `attachment:${'a'.repeat(64)}`
  const planDigest = createHash('sha256').update(canonicalPresentationValue(plan)).digest('hex')
  const run = {
    id: 'audit',
    sequence: 1,
    scope: 'source_excerpt_audit' as const,
    planRevision: 1,
    planDigest,
    sourceRefs: [{ sourceId: plan.sources[0]!.id, attachmentId: 'a'.repeat(64) }],
    state: 'completed' as const,
    startedAt: '2026-09-29T00:00:00.000Z',
    finishedAt: '2026-09-29T00:01:00.000Z',
    sources: [
      {
        sourceId: plan.sources[0]!.id,
        attachmentId: 'a'.repeat(64),
        status: 'found' as const,
        offset: 0,
      },
    ],
  }
  const history = presentationSourceAuditHistory({
    version: 1,
    projectId: plan.projectId,
    documentId: 'doc',
    revision: 2,
    runs: [run],
  })
  let documentId = 'doc'
  const options = {
    request: vi.fn(
      async (body: unknown) =>
        new Response(
          JSON.stringify(
            (body as { operation: string }).operation === 'read_source_audit'
              ? { projectId: plan.projectId, documentId: 'doc', audit: run }
              : (body as { operation: string }).operation === 'get_plan'
                ? { projectId: plan.projectId, revision: 1, plan }
                : {
                    projectId: plan.projectId,
                    title: plan.title,
                    status: 'planned',
                    slideCount: plan.slides.length,
                    slides: plan.slides.map(({ id, title }) => ({ id, title })),
                    history: [],
                    plan: { revision: 1, value: plan },
                    sourceAuditHistory: history,
                  },
          ),
        ),
    ),
    available: () => true,
    lastProject: () => plan.projectId,
    documentId: async () => documentId,
    executeTool: vi.fn(async () => ({ output: '', summary: '' })),
    rememberProject: async () => {},
  }
  return {
    options,
    plan,
    run,
    history,
    create: () => createPresentationProjectController(options),
    changeDocument: () => {
      documentId = 'foreign'
    },
  }
}
it('restores matching durable audit results after controller recreation while preserving historical event identity', async () => {
  const f = fixture(),
    controller = f.create()
  await controller.refresh()
  expect(controller.snapshot().project?.sourceAuditHistory?.runs[0]).toMatchObject({
    id: 'audit',
    sourceCount: 1,
    foundCount: 1,
  })
  expect(controller.snapshot().sourceAudit).toMatchObject({
    planRevision: 1,
    sources: f.run.sources,
  })
  const expected = controller.snapshot().sourceAudit
  controller.clear()
  const reopened = f.create()
  await reopened.refresh()
  expect(reopened.snapshot().sourceAudit).toEqual(expected)
})
it('refreshes durable research history immediately after a new workbench audit without losing its current project', async () => {
  const f = fixture(),
    controller = f.create()
  await controller.refresh()
  const next = {
    ...f.run,
    id: 'second',
    sequence: 3,
    startedAt: '2026-09-29T00:02:00.000Z',
    finishedAt: '2026-09-29T00:03:00.000Z',
  }
  const request = f.options.request
  f.options.request = vi.fn(async (body) =>
    (body as { operation: string; auditId?: string }).operation === 'read_source_audit' &&
    (body as { auditId?: string }).auditId === 'second'
      ? new Response(
          JSON.stringify({ projectId: f.plan.projectId, documentId: 'doc', audit: next }),
        )
      : request(body),
  )
  f.options.executeTool.mockImplementation(async () => {
    f.history.runs.push(
      presentationSourceAuditHistory({
        version: 1,
        projectId: f.plan.projectId,
        documentId: 'doc',
        revision: 4,
        runs: [next],
      }).runs[0]!,
    )
    f.history.revision = 4
    return {
      output: JSON.stringify({
        projectId: f.plan.projectId,
        planRevision: 1,
        sources: next.sources,
        checks: {
          support: 'not_verified',
          sourceAuthority: 'not_verified',
          timeliness: 'not_verified',
        },
      }),
      summary: '',
    }
  })
  await controller.auditSources()
  expect(controller.snapshot().sourceAudit).toMatchObject({
    auditId: 'second',
    finishedAt: next.finishedAt,
  })
  expect(controller.snapshot().project?.sourceAuditHistory?.runs).toHaveLength(2)
})
it('never adopts older or mismatched-plan audit results as current and retains a valid project when history is unreadable', async () => {
  const f = fixture()
  f.history.runs[0]!.planDigest = 'b'.repeat(64)
  const controller = f.create()
  await controller.refresh()
  expect(controller.snapshot().sourceAudit).toBeUndefined()
  expect(
    f.options.request.mock.calls.every(
      ([body]) => (body as { operation: string }).operation !== 'read_source_audit',
    ),
  ).toBe(true)
  f.history.runs[0]!.sourceCount = 999
  await controller.refresh()
  expect(controller.snapshot().project?.plan).toBeDefined()
  expect(controller.snapshot().project?.sourceAuditHistoryUnavailable).toBe(true)
})
it.each(['document', 'cancel', 'foreign_result'] as const)(
  'isolates %s while a durable audit read is pending',
  async (kind) => {
    const f = fixture()
    const original = f.options.request
    let release!: () => void
    const waiting = new Promise<void>((resolve) => {
      release = resolve
    })
    f.options.request = vi.fn(async (body) => {
      if ((body as { operation: string }).operation === 'read_source_audit') {
        await waiting
        if (kind === 'foreign_result')
          return new Response(
            JSON.stringify({ projectId: f.plan.projectId, documentId: 'foreign', audit: f.run }),
          )
      }
      return original(body)
    })
    const controller = f.create(),
      pending = controller.refresh()
    await vi.waitFor(() =>
      expect(
        f.options.request.mock.calls.some(
          ([body]) => (body as { operation: string }).operation === 'read_source_audit',
        ),
      ).toBe(true),
    )
    if (kind === 'document') f.changeDocument()
    if (kind === 'cancel') controller.cancel()
    release()
    await pending
    expect(controller.snapshot().sourceAudit).toBeUndefined()
    if (kind === 'foreign_result')
      expect(controller.snapshot().project?.sourceAuditHistoryUnavailable).toBe(true)
  },
)
