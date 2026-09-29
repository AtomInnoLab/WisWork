import { expect, it, vi } from 'vitest'
import type { AgentSkill } from '@wiswork/agent-core'
import { createPresentationResearchController } from '../src/agent/presentation-research.js'
import { researchSummary } from './presentation-research-fixture.js'
function fixture() {
  let summary = researchSummary()
  const documentId = vi.fn(async () => 'doc')
  const receipt = (body: Record<string, unknown>) => ({
    version: 1,
    documentId: 'doc',
    projectId: 'research',
    ledgerId: 'ledger1',
    sequence: 1,
    draftDigest: summary.records[0]?.draftDigest ?? researchSummary().records[0]!.draftDigest,
    deleteId: body.deleteId,
    deletedAt: '2026-09-29T00:02:00.000Z',
    revision: 3,
  })
  const request = vi.fn(async (raw: unknown, _signal?: AbortSignal) => {
    const body = raw as Record<string, unknown>
    if (body.operation === 'research_capabilities')
      return new Response(
        JSON.stringify({
          version: 1,
          available: true,
          cleanupAvailable: true,
          historyVersions: [1, 2],
        }),
      )
    if (body.operation === 'research_delete') {
      const reply = receipt(body)
      summary = {
        version: 2,
        documentId: 'doc',
        projectId: 'research',
        revision: 3,
        totalRecords: 0,
        lastSequence: 1,
        records: [],
      }
      return new Response(JSON.stringify(reply))
    }
    return new Response(JSON.stringify({ error: 'not_found' }))
  })
  const executeTool = vi.fn<AgentSkill['executeTool']>(async () => ({
    output: JSON.stringify(summary),
    mutated: false,
    summary: '只读',
  }))
  const options = {
    available: () => true,
    request,
    documentId,
    lastProject: () => 'research',
    executeTool,
  }
  const controller = createPresentationResearchController(options)
  return {
    options,
    controller,
    request,
    documentId,
    executeTool,
    receipt,
    setSummary: (value: typeof summary) => {
      summary = value
    },
  }
}
it('deletes only an explicitly chosen finished record with fixed identity and CAS then restores V2 on reopen', async () => {
  const f = fixture()
  await f.controller.refresh()
  expect(f.controller.snapshot().cleanupAvailable).toBe(true)
  const digest = researchSummary().records[0]!.draftDigest
  await f.controller.deleteRecord('ledger1', digest)
  const body = f.request.mock.calls.find(
    ([raw]) => (raw as Record<string, unknown>).operation === 'research_delete',
  )![0]
  expect(body).toMatchObject({
    operation: 'research_delete',
    documentId: 'doc',
    projectId: 'research',
    ledgerId: 'ledger1',
    expectedDraftDigest: digest,
    expectedRevision: 2,
    deleteId: expect.any(String),
  })
  expect(f.controller.snapshot().summary).toMatchObject({
    version: 2,
    lastSequence: 1,
    totalRecords: 0,
  })
  expect(f.controller.snapshot().deleteReceipt?.ledgerId).toBe('ledger1')
  expect(f.controller.snapshot().notice).toContain(
    '原附件、PowerPoint 文稿、交付包与导出副本仍保留',
  )
  await f.controller.refresh()
  expect(
    f.request.mock.calls.filter(
      ([raw]) => (raw as Record<string, unknown>).operation === 'research_delete',
    ),
  ).toHaveLength(1)
})
it('recovers a lost ACK by status only and retains an explicit fixed retry when status is absent', async () => {
  const f = fixture()
  await f.controller.refresh()
  let attempt!: Record<string, unknown>
  f.request.mockImplementation(async (raw) => {
    const body = raw as Record<string, unknown>
    if (body.operation === 'research_delete') {
      attempt = body
      throw new Error('private-response')
    }
    if (body.operation === 'research_delete_status')
      return new Response(JSON.stringify(f.receipt(attempt)))
    return new Response(
      JSON.stringify({
        version: 1,
        available: true,
        cleanupAvailable: true,
        historyVersions: [1, 2],
      }),
    )
  })
  await f.controller.deleteRecord('ledger1', researchSummary().records[0]!.draftDigest)
  expect(f.controller.snapshot().deleteReceipt?.deleteId).toBe(attempt.deleteId)
  expect(
    f.request.mock.calls.filter(
      ([raw]) => (raw as Record<string, unknown>).operation === 'research_delete',
    ),
  ).toHaveLength(1)
  const g = fixture()
  await g.controller.refresh()
  const bodies: unknown[] = []
  g.request.mockImplementation(async (raw) => {
    const body = raw as Record<string, unknown>
    if (body.operation === 'research_delete') {
      bodies.push(body)
      throw new Error('lost')
    }
    if (body.operation === 'research_delete_status')
      return new Response(JSON.stringify({ error: 'not_found' }))
    return new Response(
      JSON.stringify({
        version: 1,
        available: true,
        cleanupAvailable: true,
        historyVersions: [1, 2],
      }),
    )
  })
  await g.controller.deleteRecord('ledger1', researchSummary().records[0]!.draftDigest)
  await g.controller.refresh()
  await g.controller.checkDeleteStatus()
  expect(bodies).toHaveLength(1)
  await g.controller.retryDelete()
  expect(bodies).toHaveLength(2)
  expect(bodies[1]).toEqual(bodies[0])
  expect(g.controller.snapshot().error).not.toContain('lost')
})
it('blocks running, stale digest and old capability without a delete request', async () => {
  for (const mode of ['running', 'digest', 'old']) {
    const f = fixture()
    if (mode === 'running') {
      const { finishedAt: _end, ...record } = researchSummary().records[0]!
      f.setSummary({
        ...researchSummary(),
        revision: 1,
        records: [{ ...record, state: 'running' }],
      })
    }
    if (mode === 'old')
      f.request.mockResolvedValue(new Response(JSON.stringify({ version: 1, available: true })))
    await f.controller.refresh()
    await f.controller.deleteRecord(
      'ledger1',
      mode === 'digest' ? 'f'.repeat(64) : researchSummary().records[0]!.draftDigest,
    )
    expect(
      f.request.mock.calls.some(
        ([raw]) => (raw as Record<string, unknown>).operation === 'research_delete',
      ),
    ).toBe(false)
  }
})
it('rejects forged receipts and keeps deletion failures safe and actionable', async () => {
  for (const field of [
    'documentId',
    'projectId',
    'ledgerId',
    'deleteId',
    'draftDigest',
    'sequence',
  ]) {
    const f = fixture()
    await f.controller.refresh()
    f.request.mockImplementation(async (raw) => {
      const body = raw as Record<string, unknown>
      if (body.operation === 'research_capabilities')
        return new Response(
          JSON.stringify({
            version: 1,
            available: true,
            cleanupAvailable: true,
            historyVersions: [1, 2],
          }),
        )
      return new Response(
        JSON.stringify({ ...f.receipt(body), [field]: field === 'sequence' ? 2 : 'wrong-private' }),
      )
    })
    await f.controller.deleteRecord('ledger1', researchSummary().records[0]!.draftDigest)
    expect(f.controller.snapshot().deleteReceipt).toBeUndefined()
    expect(f.controller.snapshot().error).not.toContain('wrong-private')
  }
  for (const error of [
    'record_protected',
    'record_running',
    'cleanup_quota_exceeded',
    'revision_conflict',
  ]) {
    const f = fixture()
    await f.controller.refresh()
    f.request.mockImplementation(async () => new Response(JSON.stringify({ error })))
    await f.controller.deleteRecord('ledger1', researchSummary().records[0]!.draftDigest)
    expect(f.controller.snapshot().error).toBeTruthy()
  }
})
it('ignores late delete ACK after cancellation or document change and never automatically resends', async () => {
  for (const change of ['cancel', 'document', 'project']) {
    const f = fixture()
    await f.controller.refresh()
    let finish!: (value: Response) => void
    let body!: Record<string, unknown>
    f.request.mockImplementation(async (raw) => {
      body = raw as Record<string, unknown>
      return new Promise<Response>((done) => {
        finish = done
      })
    })
    const pending = f.controller.deleteRecord('ledger1', researchSummary().records[0]!.draftDigest)
    await vi.waitFor(() => expect(finish).toBeTypeOf('function'))
    if (change === 'cancel') f.controller.cancel()
    else if (change === 'document') f.documentId.mockResolvedValue('other')
    else f.controller.clear()
    finish(new Response(JSON.stringify(f.receipt(body))))
    await pending
    expect(f.controller.snapshot().deleteReceipt).toBeUndefined()
    expect(f.controller.snapshot().phase).toBe('idle')
  }
})

it('rejects a changed document before sending cleanup from an old summary', async () => {
  const f = fixture()
  await f.controller.refresh()
  f.documentId.mockResolvedValue('other-doc')
  await f.controller.deleteRecord('ledger1', researchSummary().records[0]!.draftDigest)
  expect(
    f.request.mock.calls.some(
      ([raw]) => (raw as Record<string, unknown>).operation === 'research_delete',
    ),
  ).toBe(false)
  expect(f.controller.snapshot().error).toContain('文档已改变')
})
it('persists only bounded identity and recovers after taskpane reopen by receipt reads without deleting again', async () => {
  const { createPresentationResearchDeletePersistence } =
    await import('../src/agent/presentation-research-cleanup-storage.js')
  const values = new Map<string, string>()
  const persistence = createPresentationResearchDeletePersistence('doc', {
    getItem: (key) => values.get(key) ?? null,
    setItem: (key, value) => {
      values.set(key, value)
    },
    removeItem: (key) => {
      values.delete(key)
    },
  })
  const f = fixture()
  let attempt!: Record<string, unknown>
  let acknowledged = false
  f.request.mockImplementation(async (raw) => {
    const body = raw as Record<string, unknown>
    if (body.operation === 'research_capabilities')
      return new Response(
        JSON.stringify({
          version: 1,
          available: true,
          cleanupAvailable: true,
          historyVersions: [1, 2],
        }),
      )
    if (body.operation === 'research_delete') {
      attempt = body
      throw new Error('ACK lost')
    }
    return new Response(JSON.stringify(acknowledged ? f.receipt(attempt) : { error: 'not_found' }))
  })
  const options = {
    ...f.options,
    readDeleteAttempt: persistence.read,
    writeDeleteAttempt: persistence.write,
  }
  const first = createPresentationResearchController(options)
  await first.refresh()
  await first.deleteRecord('ledger1', researchSummary().records[0]!.draftDigest)
  expect(values.size).toBe(1)
  expect(Object.keys(JSON.parse(Array.from(values.values())[0]!)).sort()).toEqual([
    'deleteId',
    'documentId',
    'draftDigest',
    'expectedRevision',
    'ledgerId',
    'projectId',
    'sequence',
  ])
  first.cancel()
  acknowledged = true
  f.setSummary({
    version: 2,
    documentId: 'doc',
    projectId: 'research',
    revision: 3,
    totalRecords: 0,
    lastSequence: 1,
    records: [],
  })
  const reopened = createPresentationResearchController(options)
  await reopened.refresh()
  expect(reopened.snapshot().deleteReceipt?.deleteId).toBe(attempt.deleteId)
  expect(reopened.snapshot().summary?.version).toBe(2)
  expect(values.size).toBe(0)
  expect(
    f.request.mock.calls.filter(
      ([raw]) => (raw as Record<string, unknown>).operation === 'research_delete',
    ),
  ).toHaveLength(1)
  expect(() => persistence.read('another-doc')).toThrow('presentation_document_changed')
})
it('cannot bypass unavailable identity storage by explicitly retrying', async () => {
  const f = fixture()
  const controller = createPresentationResearchController({
    ...f.options,
    writeDeleteAttempt: () => {
      throw new Error('storage-private')
    },
  })
  await controller.refresh()
  await controller.deleteRecord('ledger1', researchSummary().records[0]!.draftDigest)
  await controller.retryDelete()
  expect(
    f.request.mock.calls.some(
      ([raw]) => (raw as Record<string, unknown>).operation === 'research_delete',
    ),
  ).toBe(false)
})

it('clears proven no-commit rejects so refreshing permits a newly confirmed attempt with current CAS', async () => {
  for (const error of ['revision_conflict', 'record_protected']) {
    const f = fixture()
    let first = true
    const original = f.request.getMockImplementation()!
    const writes = vi.fn()
    f.request.mockImplementation(async (raw, signal) => {
      const body = raw as Record<string, unknown>
      if (body.operation !== 'research_delete') return original(raw, signal)
      if (first) {
        first = false
        return new Response(JSON.stringify({ error }))
      }
      f.setSummary({
        version: 2,
        documentId: 'doc',
        projectId: 'research',
        revision: 4,
        totalRecords: 1,
        lastSequence: 2,
        records: [
          {
            ...researchSummary().records[0]!,
            id: 'ledger2',
            sequence: 2,
            state: 'running',
            finishedAt: undefined,
          },
        ],
      })
      return new Response(JSON.stringify({ ...f.receipt(body), revision: 4 }))
    })
    const controller = createPresentationResearchController({
      ...f.options,
      writeDeleteAttempt: writes,
    })
    await controller.refresh()
    await controller.deleteRecord('ledger1', researchSummary().records[0]!.draftDigest)
    expect(controller.snapshot().deleteAttempt).toBeUndefined()
    expect(writes).toHaveBeenLastCalledWith('doc', undefined)
    const { finishedAt: _end, ...record } = researchSummary().records[0]!
    f.setSummary({
      ...researchSummary(),
      revision: 3,
      totalRecords: 2,
      records: [
        researchSummary().records[0]!,
        { ...record, id: 'ledger2', sequence: 2, state: 'running' },
      ],
    })
    await controller.refresh()
    await controller.deleteRecord('ledger1', researchSummary().records[0]!.draftDigest)
    const bodies = f.request.mock.calls
      .map(([raw]) => raw as Record<string, unknown>)
      .filter((body) => body.operation === 'research_delete')
    expect(bodies).toHaveLength(2)
    expect(bodies[0]!.expectedRevision).toBe(2)
    expect(bodies[1]!.expectedRevision).toBe(3)
    expect(bodies[1]!.deleteId).not.toBe(bodies[0]!.deleteId)
    expect(controller.snapshot().deleteReceipt?.revision).toBe(4)
  }
})
it('reports identity storage clear failures safely without rejecting the controller promise', async () => {
  const f = fixture()
  const controller = createPresentationResearchController({
    ...f.options,
    writeDeleteAttempt: (_document, attempt) => {
      if (!attempt) throw new Error('private-storage-failed')
    },
  })
  await controller.refresh()
  f.request.mockResolvedValue(new Response(JSON.stringify({ error: 'record_protected' })))
  await expect(
    controller.deleteRecord('ledger1', researchSummary().records[0]!.draftDigest),
  ).resolves.toBeUndefined()
  expect(controller.snapshot().phase).toBe('idle')
  expect(controller.snapshot().deleteAttempt).toBeDefined()
  expect(controller.snapshot().error).not.toContain('private-storage-failed')
})
