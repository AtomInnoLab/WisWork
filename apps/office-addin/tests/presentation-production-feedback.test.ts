import {
  MAX_PRESENTATION_PRODUCTION_FEEDBACK_BYTES,
  parsePresentationProductionFeedbackLedger,
} from '@wiswork/project-store/presentation-feedback'
import { expect, it, vi } from 'vitest'
import { createPresentationProjectController } from '../src/skills/powerpoint/presentation-project.js'
const production = {
  projectId: 'p',
  requestId: 'r',
  planRevision: 1,
  status: 'compiled',
  compiledCount: 2,
  total: 2,
  pages: [
    { id: 'a', title: '甲', state: 'compiled', attempt: 1 },
    { id: 'b', title: '乙', state: 'compiled', attempt: 1 },
  ],
}
const project = {
  projectId: 'p',
  title: '反馈项目',
  status: 'compiled',
  latestRequestId: 'r',
  latestCompiledRequestId: 'r',
  slideCount: 2,
  slides: [
    { id: 'a', title: '甲' },
    { id: 'b', title: '乙' },
  ],
  history: [{ requestId: 'r', sequence: 1, status: 'compiled', slideCount: 2 }],
  production,
  checks: {
    structure: 'passed',
    geometry: 'passed',
    render: 'not_run',
    sources: 'not_verified',
    roundTrip: 'not_run',
  },
}
const ledger = {
  version: 1,
  source: 'user_reported',
  projectId: 'p',
  documentId: 'doc',
  requestId: 'r',
  inputDigest: 'a'.repeat(64),
  planDigest: 'b'.repeat(64),
  planRevision: 1,
  pageIds: ['a', 'b'],
  revision: 1,
  snapshots: [
    {
      revision: 1,
      recordedAt: '2026-09-29T00:00:00.000Z',
      pages: [
        { pageId: 'a', status: 'needs_correction', note: '人工判断' },
        { pageId: 'b', status: 'not_evaluated' },
      ],
    },
  ],
}
function fixture() {
  let feedback: unknown = null
  const request = vi.fn(
    async (body: unknown) =>
      new Response(
        JSON.stringify(
          (body as { operation: string }).operation === 'status'
            ? project
            : (body as { operation: string }).operation === 'production_job_status'
              ? { error: 'invalid_request' }
              : { feedback },
        ),
      ),
  )
  const documentId = vi.fn(async () => 'doc'),
    available = vi.fn(() => true),
    executeTool = vi.fn(async () => ({ output: '{}', summary: '', mutated: false }))
  const controller = createPresentationProjectController({
    request,
    documentId,
    available,
    executeTool,
    lastProject: () => 'p',
  })
  return {
    controller,
    request,
    documentId,
    available,
    executeTool,
    setFeedback: (value: unknown) => {
      feedback = value
    },
  }
}
it('explicitly reads and records a page patch on the exact completed frozen task without agent tools', async () => {
  const f = fixture()
  await f.controller.refresh()
  expect(f.request).toHaveBeenCalledTimes(2)
  await f.controller.recordProductionFeedback?.([{ pageId: 'a', status: 'needs_correction' }])
  expect(f.request).toHaveBeenCalledTimes(2)
  await f.controller.readProductionFeedback!()
  expect(f.controller.snapshot().productionFeedback).toBeNull()
  f.setFeedback(ledger)
  await f.controller.recordProductionFeedback!([
    { pageId: 'a', status: 'needs_correction', note: '人工判断' },
  ])
  expect(f.request).toHaveBeenLastCalledWith(
    {
      operation: 'production_feedback_record',
      documentId: 'doc',
      projectId: 'p',
      requestId: 'r',
      expectedRevision: 0,
      pages: [{ pageId: 'a', status: 'needs_correction', note: '人工判断' }],
    },
    expect.any(AbortSignal),
  )
  expect(f.controller.snapshot().productionFeedback).toEqual(ledger)
  expect(f.controller.snapshot().project?.production).toEqual(production)
  expect(f.executeTool).not.toHaveBeenCalled()
  const reopened = fixture()
  reopened.setFeedback(ledger)
  await reopened.controller.refresh()
  await reopened.controller.readProductionFeedback!()
  expect(reopened.controller.snapshot().productionFeedback).toEqual(ledger)
})
it('rejects foreign identity and clears unreadable feedback instead of retaining historical success', async () => {
  const f = fixture()
  await f.controller.refresh()
  f.setFeedback(ledger)
  await f.controller.readProductionFeedback!()
  f.setFeedback({ ...ledger, requestId: 'other' })
  await f.controller.readProductionFeedback!()
  expect(f.controller.snapshot().productionFeedback).toBeUndefined()
  expect(f.controller.snapshot().productionFeedbackUnavailable).toBe(true)
  expect(f.controller.snapshot().project?.projectId).toBe('p')
})
it('does not write after document changes and ignores a cancelled late read', async () => {
  const f = fixture()
  await f.controller.refresh()
  await f.controller.readProductionFeedback!()
  f.documentId.mockResolvedValue('other')
  await f.controller.recordProductionFeedback!([{ pageId: 'a', status: 'needs_correction' }])
  expect(
    f.request.mock.calls.some(
      ([body]) => (body as { operation: string }).operation === 'production_feedback_record',
    ),
  ).toBe(false)
  const g = fixture()
  await g.controller.refresh()
  let finish!: (value: Response) => void
  g.request.mockImplementation(
    async () =>
      new Promise((resolve) => {
        finish = resolve
      }),
  )
  const reading = g.controller.readProductionFeedback!()
  await vi.waitFor(() => expect(finish).toBeDefined())
  g.controller.clear()
  finish(new Response(JSON.stringify({ feedback: ledger })))
  await reading
  expect(g.controller.snapshot()).toEqual({ phase: 'idle' })
})

it('accepts a valid near-limit ledger plus its exact envelope and rejects an oversized response', async () => {
  const pageIds = Array.from({ length: 32 }, (_, index) => `p${index}`)
  const large = {
    ...ledger,
    pageIds,
    revision: 14,
    snapshots: Array.from({ length: 14 }, (_, index) => ({
      revision: index + 1,
      recordedAt: '2026-09-29T00:00:00.000Z',
      pages: pageIds.map((pageId) => ({ pageId, status: 'not_evaluated', note: 'a'.repeat(2000) })),
    })),
  }
  const bytes = (value: unknown) => new TextEncoder().encode(JSON.stringify(value)).byteLength
  let extra = MAX_PRESENTATION_PRODUCTION_FEEDBACK_BYTES - 5 - bytes(large)
  for (const snapshot of large.snapshots)
    for (const page of snapshot.pages) {
      const controls = Math.min(2000, Math.floor(extra / 5))
      page.note = '\0'.repeat(controls) + 'a'.repeat(2000 - controls)
      extra -= controls * 5
      if (extra > 0 && extra < 5 && controls < 2000) {
        page.note = page.note.slice(0, -extra) + '"'.repeat(extra)
        extra = 0
      }
    }
  expect(extra).toBe(0)
  expect(bytes(large)).toBe(MAX_PRESENTATION_PRODUCTION_FEEDBACK_BYTES - 5)
  expect(parsePresentationProductionFeedbackLedger(large).revision).toBe(14)
  let text = JSON.stringify({ feedback: large })
  expect(new TextEncoder().encode(text).byteLength).toBe(
    MAX_PRESENTATION_PRODUCTION_FEEDBACK_BYTES + 8,
  )
  const fullProduction = {
    ...production,
    compiledCount: 32,
    total: 32,
    pages: pageIds.map((id) => ({ id, title: id, state: 'compiled', attempt: 1 })),
  }
  const fullProject = {
    ...project,
    slideCount: 32,
    slides: pageIds.map((id) => ({ id, title: id })),
    history: [{ requestId: 'r', sequence: 1, status: 'compiled', slideCount: 32 }],
    production: fullProduction,
  }
  const f = fixture()
  f.request.mockImplementation(
    async (body) =>
      new Response(
        (body as { operation: string }).operation === 'status'
          ? JSON.stringify(fullProject)
          : (body as { operation: string }).operation === 'production_job_status'
            ? JSON.stringify({ error: 'invalid_request' })
            : text,
      ),
  )
  await f.controller.refresh()
  await f.controller.readProductionFeedback!()
  expect(f.controller.snapshot().productionFeedback?.revision).toBe(14)
  text += ' '.repeat(6)
  await f.controller.readProductionFeedback!()
  expect(f.controller.snapshot().productionFeedback).toBeUndefined()
  expect(f.controller.snapshot().productionFeedbackUnavailable).toBe(true)
})

it('rejects disappearance, rollback, changed digests or rewritten known history while accepting a valid append', async () => {
  const second = {
    ...ledger,
    revision: 2,
    snapshots: [
      ...ledger.snapshots,
      { ...ledger.snapshots[0]!, revision: 2, recordedAt: '2026-09-29T00:00:01.000Z' },
    ],
  }
  for (const next of [
    null,
    ledger,
    { ...second, inputDigest: 'c'.repeat(64) },
    {
      ...second,
      snapshots: [
        {
          ...second.snapshots[0]!,
          pages: [
            { pageId: 'a', status: 'no_correction' },
            { pageId: 'b', status: 'not_evaluated' },
          ],
        },
        second.snapshots[1]!,
      ],
    },
  ]) {
    const f = fixture()
    await f.controller.refresh()
    f.setFeedback(second)
    await f.controller.readProductionFeedback!()
    expect(f.controller.snapshot().productionFeedback?.revision).toBe(2)
    f.setFeedback(next)
    await f.controller.readProductionFeedback!()
    expect(f.controller.snapshot().productionFeedback).toBeUndefined()
    expect(f.controller.snapshot().productionFeedbackUnavailable).toBe(true)
    f.setFeedback(null)
    await f.controller.readProductionFeedback!()
    expect(f.controller.snapshot().productionFeedback).toBeUndefined()
  }
  const f = fixture()
  await f.controller.refresh()
  f.setFeedback(ledger)
  await f.controller.readProductionFeedback!()
  f.setFeedback(second)
  await f.controller.readProductionFeedback!()
  expect(f.controller.snapshot().productionFeedback).toEqual(second)
})
