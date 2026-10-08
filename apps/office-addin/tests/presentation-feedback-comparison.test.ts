import { expect, it, vi } from 'vitest'
import {
  comparisonFixture as fixture,
  comparisonReport,
  comparisonProject,
} from './presentation-feedback-comparison-fixture.js'
it('explicitly selects only a different compiled baseline without any comparison or agent mutation', async () => {
  const f = fixture()
  await f.controller.refresh()
  const count = f.request.mock.calls.length
  f.controller.selectFeedbackComparisonBaseline!('candidate')
  expect(f.controller.snapshot().feedbackComparisonBaselineRequestId).toBeUndefined()
  f.controller.selectFeedbackComparisonBaseline!('unfinished')
  expect(f.controller.snapshot().feedbackComparisonBaselineRequestId).toBeUndefined()
  f.controller.selectFeedbackComparisonBaseline!('baseline')
  expect(f.controller.snapshot().feedbackComparisonBaselineRequestId).toBe('baseline')
  expect(f.request).toHaveBeenCalledTimes(count)
  expect(f.executeTool).not.toHaveBeenCalled()
  await f.controller.readFeedbackComparison!()
  expect(f.controller.snapshot().feedbackComparisonUnavailable).toBe(true)
  expect(f.request).toHaveBeenLastCalledWith(
    {
      operation: 'production_feedback_compare',
      documentId: 'doc',
      projectId: 'p',
      requestId: 'candidate',
      baselineRequestId: 'baseline',
    },
    expect.any(AbortSignal),
  )
})
it('changing baseline or clearing invalidates a pending comparison and prevents late results', async () => {
  const f = fixture()
  await f.controller.refresh()
  f.controller.selectFeedbackComparisonBaseline!('baseline')
  let finish!: (value: Response) => void
  f.request.mockImplementation(
    async () =>
      new Promise((resolve) => {
        finish = resolve
      }),
  )
  const reading = f.controller.readFeedbackComparison!()
  await vi.waitFor(() => expect(finish).toBeDefined())
  f.controller.selectFeedbackComparisonBaseline!()
  finish(new Response(JSON.stringify({ comparison: {} })))
  await reading
  expect(f.controller.snapshot().feedbackComparison).toBeUndefined()
  expect(f.controller.snapshot().feedbackComparisonBaselineRequestId).toBeUndefined()
  expect(f.controller.snapshot().phase).toBe('idle')
})

it('reads a strictly parsed frozen comparison without agent writes and clears it when baseline changes', async () => {
  const f = fixture(),
    report = comparisonReport()
  await f.controller.refresh()
  f.setComparison(report)
  f.controller.selectFeedbackComparisonBaseline!('baseline')
  await f.controller.readFeedbackComparison!()
  expect(f.controller.snapshot().feedbackComparison).toEqual(report)
  expect(f.executeTool).not.toHaveBeenCalled()
  expect(f.controller.snapshot().project?.production?.status).toBe('compiled')
  f.controller.selectFeedbackComparisonBaseline!()
  expect(f.controller.snapshot().feedbackComparison).toBeUndefined()
})
it('rejects same-pair rollback, changed frozen inputs and same-version edits but accepts a newer feedback snapshot', async () => {
  const known = comparisonReport({ baselineRevision: 2, candidateRevision: 2 })
  for (const report of [
    comparisonReport(),
    { ...known, candidate: { ...known.candidate, inputDigest: 'e'.repeat(64) } },
    comparisonReport({ baselineRevision: 2, candidateRevision: 2, candidateNeeds: 2 }),
    comparisonReport({ baselineRevision: 2, candidateRevision: 2, missing: true }),
  ]) {
    const f = fixture()
    await f.controller.refresh()
    f.controller.selectFeedbackComparisonBaseline!('baseline')
    f.setComparison(known)
    await f.controller.readFeedbackComparison!()
    expect(f.controller.snapshot().feedbackComparison).toEqual(known)
    f.setComparison(report)
    await f.controller.readFeedbackComparison!()
    expect(f.controller.snapshot().feedbackComparison).toBeUndefined()
    expect(f.controller.snapshot().feedbackComparisonUnavailable).toBe(true)
    f.setComparison(comparisonReport())
    await f.controller.readFeedbackComparison!()
    expect(f.controller.snapshot().feedbackComparison).toBeUndefined()
  }
  const f = fixture()
  await f.controller.refresh()
  f.controller.selectFeedbackComparisonBaseline!('baseline')
  f.setComparison(known)
  await f.controller.readFeedbackComparison!()
  const newer = comparisonReport({ baselineRevision: 3, candidateRevision: 3, candidateNeeds: 2 })
  f.setComparison(newer)
  await f.controller.readFeedbackComparison!()
  expect(f.controller.snapshot().feedbackComparison).toEqual(newer)
})
it('rejects foreign scope and a capability or document change before sending', async () => {
  for (const change of [{ documentId: 'other' }, { projectId: 'other' }]) {
    const f = fixture()
    await f.controller.refresh()
    f.controller.selectFeedbackComparisonBaseline!('baseline')
    f.setComparison({ ...comparisonReport(), ...change })
    await f.controller.readFeedbackComparison!()
    expect(f.controller.snapshot().feedbackComparisonUnavailable).toBe(true)
  }
  const f = fixture()
  await f.controller.refresh()
  f.controller.selectFeedbackComparisonBaseline!('baseline')
  f.documentId.mockResolvedValue('other')
  await f.controller.readFeedbackComparison!()
  expect(
    f.request.mock.calls.some(
      ([body]) => (body as { operation: string }).operation === 'production_feedback_compare',
    ),
  ).toBe(false)
})

it('requires the current known feedback snapshot at the same version and disallows a later version with earlier time', async () => {
  const known = comparisonReport({ candidateRevision: 2 }),
    side = known.candidate
  const feedback = {
    version: 1,
    source: 'user_reported',
    documentId: 'doc',
    projectId: 'p',
    requestId: 'candidate',
    inputDigest: side.inputDigest,
    planDigest: side.planDigest,
    planRevision: 2,
    pageIds: side.pages.map((page) => page.pageId),
    revision: 2,
    snapshots: [
      { revision: 1, recordedAt: '2026-09-29T00:00:00.000Z', pages: side.pages },
      { revision: 2, recordedAt: side.feedbackRecordedAt, pages: side.pages },
    ],
  }
  const cases = [
    comparisonReport({ candidateRevision: 2, candidateNeeds: 2 }),
    { ...known, candidate: { ...side, feedbackRecordedAt: '2026-09-29T00:00:02.000Z' } },
    {
      ...comparisonReport({ candidateRevision: 3 }),
      candidate: {
        ...comparisonReport({ candidateRevision: 3 }).candidate,
        feedbackRecordedAt: '2026-09-29T00:00:00.000Z',
      },
    },
  ]
  for (const comparison of [
    ...cases,
    known,
    comparisonReport({ candidateRevision: 3, candidateNeeds: 2 }),
  ]) {
    const f = fixture()
    f.request.mockImplementation(async (body) => {
      const operation = (body as { operation: string }).operation
      return new Response(
        JSON.stringify(
          operation === 'status'
            ? comparisonProject
            : operation === 'production_feedback_read'
              ? { feedback }
              : operation === 'production_feedback_compare'
                ? { comparison }
                : { error: 'invalid_request' },
        ),
      )
    })
    await f.controller.refresh()
    await f.controller.readProductionFeedback!()
    expect(f.controller.snapshot().productionFeedback?.revision).toBe(2)
    f.controller.selectFeedbackComparisonBaseline!('baseline')
    await f.controller.readFeedbackComparison!()
    if (cases.includes(comparison)) {
      expect(f.controller.snapshot().feedbackComparison).toBeUndefined()
      expect(f.controller.snapshot().feedbackComparisonUnavailable).toBe(true)
    } else expect(f.controller.snapshot().feedbackComparison).toEqual(comparison)
  }
})
