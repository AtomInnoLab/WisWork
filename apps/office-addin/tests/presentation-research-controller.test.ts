import { expect, it, vi } from 'vitest'
import type { AgentSkill } from '@wiswork/agent-core'
import { createPresentationResearchController } from '../src/agent/presentation-research.js'
import { researchRecord, researchSummary } from './presentation-research-fixture.js'
function fixture() {
  const request = vi.fn(
    async (_body: unknown, _signal?: AbortSignal) =>
      new Response(JSON.stringify({ version: 1, available: true })),
  )
  const executeTool = vi.fn<AgentSkill['executeTool']>(async (call) => ({
    output: JSON.stringify(
      call.name === 'list_research_ledgers'
        ? researchSummary()
        : call.name === 'read_research_ledger'
          ? researchRecord()
          : { paths: ['research.json', 'research.md'] },
    ),
    mutated: false,
    summary: '研究已整理，未核验',
  }))
  const documentId = vi.fn(async () => 'doc')
  const options = {
    request,
    executeTool,
    documentId,
    available: () => true,
    lastProject: () => 'research',
  }
  return {
    request,
    executeTool,
    documentId,
    options,
    controller: createPresentationResearchController(options),
  }
}
it('probes and restores independent research before any plan or production exists', async () => {
  const f = fixture()
  await f.controller.refresh()
  expect(f.controller.snapshot().available).toBe(true)
  expect(f.controller.snapshot().summary).toEqual(researchSummary())
  await f.controller.read('ledger1')
  expect(f.controller.snapshot().record).toEqual(researchRecord())
  await f.controller.export('ledger1')
  expect(f.executeTool).toHaveBeenLastCalledWith(
    expect.objectContaining({
      name: 'export_research_ledger',
      input: { project_id: 'research', ledger_id: 'ledger1' },
    }),
    expect.any(AbortSignal),
  )
  const reopened = createPresentationResearchController(f.options)
  await reopened.refresh()
  expect(reopened.snapshot().summary).toEqual(researchSummary())
  expect(f.request).toHaveBeenCalledWith(
    { operation: 'research_capabilities', documentId: 'doc', includeCleanup: true },
    expect.any(AbortSignal),
  )
})
it('hides unsupported old PCs and does not interpret malformed capabilities as support', async () => {
  const f = fixture()
  for (const error of ['invalid_request', 'upgrade_required']) {
    f.request.mockImplementation(async () => new Response(JSON.stringify({ error })))
    await f.controller.refresh()
    expect(f.controller.snapshot().available).toBe(false)
  }
  expect(f.executeTool).not.toHaveBeenCalled()
  f.request.mockResolvedValueOnce(
    new Response(JSON.stringify({ version: 1, available: true, raw: 'private' })),
  )
  await f.controller.refresh()
  expect(f.controller.snapshot().available).not.toBe(true)
  expect(f.controller.snapshot().error).not.toContain('private')
})
it('rejects cross-document and late records after clear or document switching', async () => {
  for (const clear of [true, false]) {
    const f = fixture()
    await f.controller.refresh()
    let finish!: (value: Awaited<ReturnType<AgentSkill['executeTool']>>) => void
    f.executeTool.mockImplementationOnce(
      () =>
        new Promise((done) => {
          finish = done
        }),
    )
    const pending = f.controller.read('ledger1')
    await vi.waitFor(() => expect(finish).toBeTypeOf('function'))
    if (clear) f.controller.clear()
    else f.documentId.mockResolvedValue('other')
    finish({ output: JSON.stringify(researchRecord()), mutated: false, summary: 'late' })
    await pending
    expect(f.controller.snapshot().record).toBeUndefined()
  }
  const f = fixture()
  await f.controller.refresh()
  f.executeTool.mockResolvedValueOnce({
    output: JSON.stringify({ ...researchRecord(), documentId: 'other' }),
    mutated: false,
    summary: 'invalid',
  })
  await f.controller.read('ledger1')
  expect(f.controller.snapshot().record).toBeUndefined()
})

it('selects an independent project explicitly and restores after a lost response without rebuilding research', async () => {
  const f = fixture()
  const controller = createPresentationResearchController({
    ...f.options,
    lastProject: () => undefined,
  })
  await controller.refresh()
  expect(controller.snapshot().available).toBe(true)
  expect(controller.snapshot().summary).toBeUndefined()
  await controller.selectProject('research')
  expect(controller.snapshot().summary).toEqual(researchSummary())
  f.executeTool.mockRejectedValueOnce(new Error('private-http-body'))
  await controller.read('ledger1')
  expect(controller.snapshot().error).not.toContain('private-http-body')
  await controller.refresh()
  expect(controller.snapshot().summary).toEqual(researchSummary())
  expect(f.executeTool.mock.calls.every(([call]) => call.name !== 'build_research_ledger')).toBe(
    true,
  )
})
it('stops capability waits before dispatching a tool and rejects wrong-project summaries', async () => {
  const f = fixture()
  let finish!: (value: Response) => void
  f.request.mockImplementationOnce(
    () =>
      new Promise((done) => {
        finish = done
      }),
  )
  const pending = f.controller.refresh()
  await vi.waitFor(() => expect(finish).toBeTypeOf('function'))
  f.controller.cancel()
  finish(new Response(JSON.stringify({ version: 1, available: true })))
  await pending
  expect(f.executeTool).not.toHaveBeenCalled()
  f.executeTool.mockResolvedValueOnce({
    output: JSON.stringify({ ...researchSummary(), projectId: 'other' }),
    mutated: false,
    summary: 'invalid',
  })
  await f.controller.refresh()
  expect(f.controller.snapshot().summary).toBeUndefined()
})

it('keeps confirmed research capability visible and refreshable after the initial list response is lost', async () => {
  const f = fixture()
  f.executeTool.mockRejectedValueOnce(new Error('private-network-detail'))
  await f.controller.refresh()
  expect(f.controller.snapshot().available).toBe(true)
  expect(f.controller.snapshot().error).toContain('刷新本机记录')
  expect(f.controller.snapshot().error).not.toContain('private-network-detail')
  await f.controller.refresh()
  expect(f.controller.snapshot().summary).toEqual(researchSummary())
})
