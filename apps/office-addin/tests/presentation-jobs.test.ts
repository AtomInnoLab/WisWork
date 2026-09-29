import { describe, expect, it, vi } from 'vitest'
import { createPresentationJobsSkill } from '../src/skills/powerpoint/presentation-jobs.js'
const production = {
  inputDigest: 'a'.repeat(64),
  planDigest: 'b'.repeat(64),
  projectId: 'p',
  requestId: 'r',
  planRevision: 1,
  status: 'pending',
  compiledCount: 0,
  total: 1,
  pages: [{ id: 'a', title: 'A', state: 'pending', attempt: 0 }],
}
const job = {
  version: 1,
  projectId: 'p',
  documentId: 'd',
  requestId: 'r',
  inputDigest: 'a'.repeat(64),
  planDigest: 'b'.repeat(64),
  planRevision: 1,
  revision: 1,
  state: 'running',
  events: [{ sequence: 1, createdAt: '2026-09-24T00:00:00.000Z', type: 'run.started' }],
}
function fixture(value: unknown = { job, production }) {
  const request = vi.fn(async () => new Response(JSON.stringify(value)))
  const documentId = vi.fn(async () => 'd')
  const skill = createPresentationJobsSkill({ request, documentId, available: () => true })
  const call = (
    name = 'read_presentation_production_job',
    input = { project_id: 'p', request_id: 'r' },
  ) => skill.executeTool({ id: 'x', name, input })
  return { request, documentId, skill, call }
}
describe('background production job skill', () => {
  it.each(['start', 'read', 'pause', 'resume', 'cancel'])(
    'maps %s to a bounded PC operation without host mutation',
    async (action) => {
      const f = fixture()
      const result = await f.call(`${action}_presentation_production_job`)
      expect(result.isError).toBeFalsy()
      expect(result.mutated).toBe(false)
      expect(f.request).toHaveBeenCalledWith(
        {
          operation: `production_job_${action === 'read' ? 'status' : action}`,
          documentId: 'd',
          projectId: 'p',
          requestId: 'r',
        },
        undefined,
      )
    },
  )
  it('rejects foreign identity and mismatched plan revision', async () => {
    for (const change of [
      { documentId: 'other' },
      { requestId: 'other' },
      { planRevision: 2 },
      { inputDigest: 'c'.repeat(64) },
      {
        state: 'completed',
        events: [{ sequence: 1, createdAt: '2026-09-24T00:00:00.000Z', type: 'run.completed' }],
      },
    ])
      expect((await fixture({ job: { ...job, ...change }, production }).call()).isError).toBe(true)
  })
  it('invalidates late responses after clear or document change', async () => {
    const f = fixture()
    f.request.mockImplementation(async () => {
      f.skill.clear()
      return new Response(JSON.stringify({ job, production }))
    })
    expect((await f.call()).output).toBe('cancelled')
    const g = fixture()
    g.documentId.mockResolvedValueOnce('d').mockResolvedValue('other')
    expect((await g.call()).isError).toBe(true)
  })
  it('accepts no persisted job and rejects unrecognized input', async () => {
    expect((await fixture({ job: null, production }).call()).isError).toBeFalsy()
    expect((await fixture().call('unknown')).isError).toBe(true)
  })
})

it('reads durable font failure on its real page and retains the exact resume request', async () => {
  const value = {
    production: {
      ...production,
      status: 'partial',
      pages: [{ id: 'a', title: 'A', state: 'failed', attempt: 1, error: 'font_unavailable' }],
    },
    job: {
      ...job,
      revision: 4,
      state: 'failed',
      events: [
        ...job.events,
        {
          sequence: 2,
          createdAt: '2026-09-24T00:00:00.500Z',
          type: 'page.started',
          pageId: 'a',
          attempt: 1,
        },
        {
          sequence: 3,
          createdAt: '2026-09-24T00:00:01.000Z',
          type: 'page.failed',
          pageId: 'a',
          attempt: 1,
          error: 'font_unavailable',
        },
        {
          sequence: 4,
          createdAt: '2026-09-24T00:00:02.000Z',
          type: 'run.failed',
          error: 'font_unavailable',
        },
      ],
    },
  }
  const f = fixture(value),
    read = await f.call()
  expect(read.isError).toBeFalsy()
  const {
    inputDigest: _inputDigest,
    planDigest: _planDigest,
    ...parsedProduction
  } = value.production
  expect(JSON.parse(read.output)).toEqual({ job: value.job, production: parsedProduction })
  expect(read.mutated).toBe(false)
  await f.call('resume_presentation_production_job')
  expect(f.request).toHaveBeenLastCalledWith(
    { operation: 'production_job_resume', documentId: 'd', projectId: 'p', requestId: 'r' },
    undefined,
  )
  const bad = fixture({
    ...value,
    job: {
      ...value.job,
      events: value.job.events.map((event) =>
        event.type === 'page.failed' ? { ...event, error: 'arbitrary_failure' } : event,
      ),
    },
  })
  expect((await bad.call()).isError).toBe(true)
})
