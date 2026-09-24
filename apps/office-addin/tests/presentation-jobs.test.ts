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
