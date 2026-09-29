import { expect, it, vi } from 'vitest'
import { createOfficeHostRuntime } from '../src/agent/host-runtime.js'
import { researchRecord, researchSummary } from './presentation-research-fixture.js'
it('registers actual research tools and refreshes a standalone project after build before any plan', async () => {
  const record = researchRecord()
  const summary = researchSummary()
  const request = vi.fn(
    async (body: unknown) =>
      new Response(
        JSON.stringify(
          (body as { operation: string }).operation === 'research_capabilities'
            ? { version: 1, available: true }
            : (body as { operation: string }).operation === 'research_build'
              ? { history: summary, record }
              : (body as { operation: string }).operation === 'research_list'
                ? summary
                : (body as { operation: string }).operation === 'research_read'
                  ? record
                  : { error: 'not_found' },
        ),
      ),
  )
  let lastProject: string | undefined
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      available: () => true,
      request,
      documentId: async () => 'doc',
      lastProject: () => lastProject,
      rememberProject: async (id) => {
        lastProject = id
      },
    },
  })
  try {
    expect(runtime.research).toBeDefined()
    expect(runtime.skill.tools.map((tool) => tool.name)).toEqual(
      expect.arrayContaining([
        'build_research_ledger',
        'read_research_ledger',
        'list_research_ledgers',
        'export_research_ledger',
      ]),
    )
    const result = await runtime.skill.executeTool({
      id: 'build',
      name: 'build_research_ledger',
      input: {
        project_id: 'research',
        ledger_id: 'ledger1',
        expected_revision: 0,
        draft: record.draft,
      },
    })
    expect(result.isError).not.toBe(true)
    await vi.waitFor(() => expect(runtime.research?.snapshot().summary).toEqual(summary))
    expect(runtime.presentation?.snapshot().project).toBeUndefined()
    await runtime.research?.read('ledger1')
    expect(runtime.research?.snapshot().record).toEqual(record)
    runtime.clearSession()
    expect(runtime.research?.snapshot().record).toBeUndefined()
    await runtime.research?.refresh()
    expect(runtime.research?.snapshot().summary).toEqual(summary)
  } finally {
    runtime.dispose()
  }
})

it('routes the current-plan binding to its exact archived ledger even when a newer research record exists', async () => {
  const { benchmarkPlan } =
    await import('../../../packages/pptx-engine/tests/fixtures/presentation-plan.js')
  const record = researchRecord(),
    summary = researchSummary(),
    plan = benchmarkPlan()
  plan.projectId = 'research'
  plan.research = {
    ledgerId: record.id,
    sequence: record.sequence,
    draftDigest: record.draftDigest,
    sources: [{ sourceId: 'source', researchSourceId: 'source1' }],
    claims: [{ claimId: 'source-1', researchClaimId: 'claim1' }],
  }
  const latest = {
    ...summary.records[0]!,
    id: 'ledger2',
    sequence: 2,
    draftDigest: 'c'.repeat(64),
    startedAt: '2026-09-29T00:00:02.000Z',
    finishedAt: '2026-09-29T00:00:03.000Z',
  }
  const currentSummary = {
    ...summary,
    revision: 4,
    totalRecords: 2,
    records: [...summary.records, latest],
  }
  const request = vi.fn(async (body: unknown) => {
    const operation = (body as { operation: string }).operation
    return new Response(
      JSON.stringify(
        operation === 'status'
          ? {
              projectId: 'research',
              title: plan.title,
              status: 'planned',
              plan: { revision: 1, value: plan },
              slideCount: plan.slides.length,
              slides: plan.slides.map(({ id, title }) => ({ id, title })),
              history: [],
            }
          : operation === 'research_capabilities'
            ? { version: 1, available: true }
            : operation === 'research_list'
              ? currentSummary
              : operation === 'research_read'
                ? record
                : { error: 'invalid_request' },
      ),
    )
  })
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      available: () => true,
      request,
      documentId: async () => 'doc',
      lastProject: () => 'research',
      rememberProject: async () => {},
    },
  })
  try {
    await runtime.presentation?.refresh()
    await runtime.presentation?.readBoundResearch?.()
    expect(runtime.presentation?.snapshot().error).toBeUndefined()
    expect(runtime.research?.snapshot().record?.id).toBe('ledger1')
    expect(runtime.research?.snapshot().record?.draftDigest).toBe(record.draftDigest)
    expect(
      request.mock.calls
        .filter(([body]) => (body as { operation: string }).operation === 'research_read')
        .map(([body]) => (body as { ledgerId: string }).ledgerId),
    ).toEqual(['ledger1'])
    expect(runtime.presentation?.snapshot().project?.plan?.value.research).toEqual(plan.research)
  } finally {
    runtime.dispose()
  }
})
