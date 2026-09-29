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
