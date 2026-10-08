import { expect, it, vi } from 'vitest'
import { createPresentationResearchController } from '../src/agent/presentation-research.js'
import { researchRecord, researchSummary } from './presentation-research-fixture.js'
function fixture() {
  const running = researchRecord()
  running.state = 'running'
  delete running.finishedAt
  delete running.sources
  const terminal = { ...researchRecord(), state: 'failed' as const, error: 'aborted' as const }
  delete terminal.sources
  const summary = researchSummary()
  summary.revision = 1
  summary.records[0]!.state = 'running'
  delete summary.records[0]!.finishedAt
  let saved: unknown
  const writeAbandonAttempt = vi.fn((_doc: string, value: unknown) => {
    saved = value
  })
  const summaryReply = vi.fn(async () => summary)
  const request = vi.fn(async (body: unknown) => {
    const op = (body as { operation: string }).operation
    if (op === 'research_list') return Response.json(await summaryReply())
    return new Response(
      JSON.stringify(
        op === 'research_capabilities'
          ? {
              version: 1,
              available: true,
              cleanupAvailable: true,
              recoveryAvailable: true,
              historyVersions: [1, 2],
            }
          : terminal,
      ),
    )
  })
  const documentId = vi.fn(async () => 'doc')
  const options = {
    available: () => true,
    documentId,
    lastProject: () => 'research',
    request,
    readAbandonAttempt: () => saved,
    writeAbandonAttempt,
    executeTool: vi.fn(async () => ({
      output: JSON.stringify(summary),
      mutated: false,
      summary: '研究',
    })),
  }
  return {
    options,
    summary,
    summaryReply,
    running,
    terminal,
    request,
    writeAbandonAttempt,
    documentId,
    controller: createPresentationResearchController(options),
  }
}
it('persists exact running identity, ends only explicitly and publishes actual terminal record without ownership claim', async () => {
  const f = fixture()
  await f.controller.refresh()
  expect(f.controller.snapshot().recoveryAvailable).toBe(true)
  await f.controller.abandonRecord('ledger1', f.running.draftDigest)
  expect(f.request).toHaveBeenCalledWith(
    {
      operation: 'research_abandon',
      documentId: 'doc',
      projectId: 'research',
      ledgerId: 'ledger1',
      expectedDraftDigest: f.running.draftDigest,
      expectedRevision: 1,
    },
    expect.any(AbortSignal),
  )
  expect(f.writeAbandonAttempt.mock.calls[0]![1]).toEqual({
    documentId: 'doc',
    projectId: 'research',
    ledgerId: 'ledger1',
    sequence: 1,
    draftDigest: f.running.draftDigest,
    expectedRevision: 1,
  })
  expect(f.controller.snapshot().abandonRecord).toEqual(f.terminal)
  expect(f.controller.snapshot().notice).toContain('不证明由本次操作结束')
  expect(f.controller.snapshot().abandonAttempt).toBeUndefined()
})
it('lost ACK and reopening recover only by exact readonly read, retaining original CAS for explicit retry', async () => {
  const f = fixture()
  await f.controller.refresh()
  f.request.mockImplementation(async (body) => {
    const op = (body as { operation: string }).operation
    if (op === 'research_list') return Response.json(await f.summaryReply())
    if (op === 'research_abandon') throw new Error('lost')
    return new Response(
      JSON.stringify(
        op === 'research_read'
          ? f.running
          : {
              version: 1,
              available: true,
              cleanupAvailable: true,
              recoveryAvailable: true,
              historyVersions: [1, 2],
            },
      ),
    )
  })
  await f.controller.abandonRecord('ledger1', f.running.draftDigest)
  expect(f.controller.snapshot().abandonAttempt?.expectedRevision).toBe(1)
  const reopened = createPresentationResearchController(f.options)
  const before = f.request.mock.calls.length
  await reopened.refresh()
  expect(
    f.request.mock.calls.slice(before).map(([body]) => (body as { operation: string }).operation),
  ).toEqual(['research_capabilities', 'research_list', 'research_read'])
  expect(reopened.snapshot().abandonAttempt?.ledgerId).toBe('ledger1')
  Object.assign(f.summary, { version: 2, lastSequence: 2, revision: 4 })
  await reopened.retryAbandon()
  expect(f.request).toHaveBeenCalledWith(
    expect.objectContaining({ operation: 'research_abandon', expectedRevision: 1 }),
    expect.any(AbortSignal),
  )
})
it('refuses to send when storage fails and safely retains identity when terminal clear fails', async () => {
  const f = fixture()
  await f.controller.refresh()
  f.writeAbandonAttempt.mockImplementation(() => {
    throw new Error('disk')
  })
  await expect(
    f.controller.abandonRecord('ledger1', f.running.draftDigest),
  ).resolves.toBeUndefined()
  expect(
    f.request.mock.calls.filter(
      ([body]) => (body as { operation: string }).operation === 'research_abandon',
    ),
  ).toHaveLength(0)
  f.writeAbandonAttempt.mockImplementation((_doc, value) => {
    if (!value) throw new Error('remove')
  })
  await expect(f.controller.retryAbandon()).resolves.toBeUndefined()
  expect(f.controller.snapshot().abandonAttempt).toBeDefined()
  expect(f.controller.snapshot().error).not.toContain('remove')
})
it('keeps ambiguous records and rejects forged SHA, terminal identity, cancel and document switches', async () => {
  for (const variant of ['sha', 'identity', 'cancel', 'document']) {
    const f = fixture()
    await f.controller.refresh()
    f.request.mockImplementation(async () => {
      if (variant === 'cancel') f.controller.cancel()
      if (variant === 'document') f.documentId.mockResolvedValue('other')
      return new Response(
        JSON.stringify(
          variant === 'sha'
            ? { ...f.terminal, draft: { ...f.terminal.draft, scope: 'forged' } }
            : variant === 'identity'
              ? { ...f.terminal, id: 'foreign' }
              : f.terminal,
        ),
      )
    })
    await f.controller.abandonRecord('ledger1', f.running.draftDigest)
    expect(f.controller.snapshot().abandonRecord).toBeUndefined()
  }
})
it('safe no-commit rejection releases metadata so fresh explicit confirmation can use refreshed revision', async () => {
  const f = fixture()
  await f.controller.refresh()
  f.request.mockResolvedValueOnce(new Response(JSON.stringify({ error: 'revision_conflict' })))
  await f.controller.abandonRecord('ledger1', f.running.draftDigest)
  expect(f.controller.snapshot().abandonAttempt).toBeUndefined()
  Object.assign(f.summary, { version: 2, lastSequence: 2, revision: 4 })
  await f.controller.refresh()
  await f.controller.abandonRecord('ledger1', f.running.draftDigest)
  const writes = f.request.mock.calls.filter(
    ([body]) => (body as { operation: string }).operation === 'research_abandon',
  )
  expect(writes.at(-1)).toEqual([
    expect.objectContaining({ operation: 'research_abandon', expectedRevision: 4 }),
    expect.any(AbortSignal),
  ])
  expect(f.request).toHaveBeenLastCalledWith(
    expect.objectContaining({ operation: 'research_list' }),
    expect.any(AbortSignal),
  )
})
it('does not accept unrelated terminal ACK as proof and checks actual record read-only', async () => {
  const f = fixture()
  await f.controller.refresh()
  f.request.mockImplementation(
    async (body) =>
      new Response(
        JSON.stringify(
          (body as { operation: string }).operation === 'research_abandon'
            ? researchRecord()
            : f.running,
        ),
      ),
  )
  await f.controller.abandonRecord('ledger1', f.running.draftDigest)
  expect(f.request).toHaveBeenLastCalledWith(
    expect.objectContaining({ operation: 'research_read', ledgerId: 'ledger1' }),
    expect.any(AbortSignal),
  )
  expect(f.controller.snapshot().abandonRecord).toBeUndefined()
  expect(f.controller.snapshot().abandonAttempt).toBeDefined()
})
it('explicit forgetting only clears local recovery identity, never mutates PC', async () => {
  const f = fixture()
  await f.controller.refresh()
  f.request.mockRejectedValue(new Error('lost'))
  await f.controller.abandonRecord('ledger1', f.running.draftDigest)
  const calls = f.request.mock.calls.length
  await f.controller.forgetAbandon()
  expect(f.request.mock.calls).toHaveLength(calls)
  expect(f.controller.snapshot().abandonAttempt).toBeUndefined()
  expect(f.controller.snapshot().notice).toContain('仅清除本机恢复身份')
})
it('refreshes real summary revision after terminal read and clears summary if readonly refresh fails', async () => {
  for (const fail of [false, true]) {
    const f = fixture()
    await f.controller.refresh()
    const latest = researchSummary()
    f.summaryReply.mockImplementation(async () => {
      if (fail) throw new Error('list unavailable')
      return latest
    })
    await f.controller.abandonRecord('ledger1', f.running.draftDigest)
    expect(f.controller.snapshot().abandonRecord).toEqual(f.terminal)
    expect(f.controller.snapshot().summary).toEqual(fail ? undefined : latest)
  }
})
