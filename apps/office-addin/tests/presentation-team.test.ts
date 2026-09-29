import { expect, it, vi } from 'vitest'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { presentationTeamId } from '@wiswork/pptx-engine/presentation-team'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import { createPresentationTeamSkill } from '../src/skills/powerpoint/presentation-team.js'
async function fixture(role = 'owner') {
  const owner = 'a'.repeat(64),
    actor = role === 'owner' ? owner : 'b'.repeat(64),
    plan = benchmarkPlan()
  let identity = { version: 1, actorSubject: actor, pcSubject: owner },
    doc = 'doc'
  let ledger = {
    version: 1,
    teamId: await presentationTeamId(owner, doc, plan.projectId),
    documentId: doc,
    projectId: plan.projectId,
    ownerSubject: owner,
    revision: 1,
    createdAt: '2026-09-29T00:00:00.000Z',
    updatedAt: '2026-09-29T00:00:00.000Z',
    publishedPlan: { revision: 1, plan },
    members: role === 'owner' || role === 'deny' ? [] : [{ subject: actor, role }],
    comments: [] as any[],
  }
  const teamRequest = vi.fn(async (body: any) => {
    if (body.operation === 'team_identity') return Response.json({ identity })
    if (body.operation === 'team_project_read') return Response.json({ team: ledger })
    if (body.operation === 'team_plan_read')
      return Response.json({ projectId: plan.projectId, revision: 1, plan })
    ledger = {
      ...ledger,
      revision: body.operation === 'team_project_create' ? 1 : ledger.revision + 1,
    }
    if (body.operation === 'team_comment_add')
      ledger.comments.push({
        ...body.comment,
        authorSubject: actor,
        planRevision: 1,
        state: 'open',
        createdAt: ledger.createdAt,
        updatedAt: ledger.updatedAt,
      })
    if (body.operation === 'team_member_set')
      ledger.members.push({ subject: body.memberSubject, role: body.role })
    return Response.json({ team: ledger })
  })
  const request = vi.fn(async () => Response.json({ projectId: plan.projectId, revision: 1, plan }))
  const proposals = createStructuredProposalController()
  const skill = createPresentationTeamSkill({
    teamAvailable: () => true,
    teamRequest,
    request,
    documentId: async () => doc,
    proposals,
  })
  return {
    skill,
    teamRequest,
    request,
    proposals,
    plan,
    get ledger() {
      return ledger
    },
    changeAccount() {
      identity = { ...identity, actorSubject: 'c'.repeat(64) }
    },
    changeDoc() {
      doc = 'other'
    },
    changePlan() {
      plan.title = 'Changed'
    },
  }
}
async function confirm(f: Awaited<ReturnType<typeof fixture>>) {
  const p = f.proposals.pending()!
  const decision = f.proposals.waitForDecision(p.id)
  await f.proposals.confirm(p.id)
  return decision
}
it('hides tools without authenticated team capability', async () => {
  const f = await fixture()
  expect(
    createPresentationTeamSkill({
      request: f.request,
      documentId: async () => 'doc',
      proposals: f.proposals,
    }).tools,
  ).toEqual([])
})
it('owner previews exact private plan and writes only after confirmation', async () => {
  const f = await fixture()
  const r = await f.skill.executeTool({
    id: 'create',
    name: 'create_presentation_team',
    input: { project_id: f.plan.projectId, plan_revision: 1 },
  })
  expect(r.isError, r.output).not.toBe(true)
  expect(f.request).not.toHaveBeenCalled()
  expect(f.teamRequest.mock.calls.some(([body]) => body.operation === 'team_plan_read')).toBe(true)
  expect(f.proposals.pending()?.preview).toMatchObject({
    planSummary: { title: f.plan.title, sourceIds: f.plan.sources.map((source) => source.id) },
  })
  expect(f.teamRequest.mock.calls.some(([b]) => b.operation === 'team_project_create')).toBe(false)
  await confirm(f)
  expect(f.teamRequest.mock.calls.some(([b]) => b.operation === 'team_project_create')).toBe(true)
})
it.each(['viewer', 'deny'])('refuses %s comment before proposal or write', async (role) => {
  const f = await fixture(role)
  const r = await f.skill.executeTool({
    id: 'comment',
    name: 'add_presentation_team_comment',
    input: {
      team_id: f.ledger.teamId,
      expected_revision: 1,
      comment: {
        id: 'comment',
        targetKind: 'slide',
        targetId: f.plan.slides[0]!.id,
        text: 'Review this',
      },
    },
  })
  expect(r.isError).toBe(true)
  expect(f.proposals.pending()).toBeUndefined()
  expect(
    f.teamRequest.mock.calls.every(([b]) =>
      ['team_identity', 'team_project_read'].includes(b.operation),
    ),
  ).toBe(true)
})
it('reviewer comments with authenticated author and cannot manage members', async () => {
  const f = await fixture('reviewer')
  const r = await f.skill.executeTool({
    id: 'comment',
    name: 'add_presentation_team_comment',
    input: {
      team_id: f.ledger.teamId,
      expected_revision: 1,
      comment: {
        id: 'comment',
        targetKind: 'slide',
        targetId: f.plan.slides[0]!.id,
        text: 'Review this',
      },
    },
  })
  expect(r.isError, r.output).not.toBe(true)
  await confirm(f)
  expect(f.ledger.comments[0].authorSubject).toBe('b'.repeat(64))
  expect(
    (
      await f.skill.executeTool({
        id: 'member',
        name: 'set_presentation_team_member',
        input: {
          team_id: f.ledger.teamId,
          expected_revision: 2,
          member_subject: 'c'.repeat(64),
          role: 'viewer',
        },
      })
    ).isError,
  ).toBe(true)
})
it.each(['account', 'doc', 'plan', 'clear'])(
  'refuses %s changes before confirmed write',
  async (kind) => {
    const f = await fixture()
    await f.skill.executeTool({
      id: 'create',
      name: 'create_presentation_team',
      input: { project_id: f.plan.projectId, plan_revision: 1 },
    })
    if (kind === 'account') f.changeAccount()
    if (kind === 'doc') f.changeDoc()
    if (kind === 'plan') f.changePlan()
    if (kind === 'clear') f.skill.clear()
    await expect(confirm(f)).rejects.toThrow()
    expect(f.teamRequest.mock.calls.some(([b]) => b.operation === 'team_project_create')).toBe(
      false,
    )
  },
)
it('rejects supplied author and stale revisions', async () => {
  const f = await fixture()
  for (const input of [
    {
      team_id: f.ledger.teamId,
      expected_revision: 2,
      comment: { id: 'comment', targetKind: 'slide', targetId: f.plan.slides[0]!.id, text: 'Text' },
    },
    {
      team_id: f.ledger.teamId,
      expected_revision: 1,
      comment: {
        id: 'comment',
        targetKind: 'slide',
        targetId: f.plan.slides[0]!.id,
        text: 'Text',
        authorSubject: 'c'.repeat(64),
      },
    },
  ])
    expect(
      (await f.skill.executeTool({ id: 'bad', name: 'add_presentation_team_comment', input }))
        .isError,
    ).toBe(true)
  expect(f.proposals.pending()).toBeUndefined()
})
it('viewer reads only published snapshot without a private plan request', async () => {
  const f = await fixture('viewer')
  const r = await f.skill.executeTool({
    id: 'read',
    name: 'read_presentation_team',
    input: { team_id: f.ledger.teamId },
  })
  expect(r.isError, r.output).not.toBe(true)
  expect(JSON.parse(r.output).team.publishedPlan.plan).toEqual(f.plan)
  expect(f.request).not.toHaveBeenCalled()
})
it('pre-cancellation performs no identity read or write', async () => {
  const f = await fixture()
  const c = new AbortController()
  c.abort()
  expect(
    await f.skill.executeTool(
      { id: 'read', name: 'read_presentation_team_identity', input: {} },
      c.signal,
    ),
  ).toMatchObject({ isError: true, output: 'cancelled' })
  expect(f.teamRequest).not.toHaveBeenCalled()
})
it('rejects a substituted mutation receipt after confirmed write', async () => {
  const f = await fixture()
  const original = f.teamRequest.getMockImplementation()!
  f.teamRequest.mockImplementation(async (body) => {
    const response = await original(body)
    if (body.operation === 'team_project_create') {
      const value = await response.json()
      value.team.ownerSubject = 'c'.repeat(64)
      return Response.json(value)
    }
    return response
  })
  await f.skill.executeTool({
    id: 'create',
    name: 'create_presentation_team',
    input: { project_id: f.plan.projectId, plan_revision: 1 },
  })
  await expect(confirm(f)).rejects.toThrow('presentation_response_invalid')
})
it('does not expose a raw transport error from confirmed mutation', async () => {
  const f = await fixture()
  const original = f.teamRequest.getMockImplementation()!
  f.teamRequest.mockImplementation(async (body) => {
    if (body.operation === 'team_project_create') throw Error('Bearer secret-token')
    return original(body)
  })
  await f.skill.executeTool({
    id: 'create',
    name: 'create_presentation_team',
    input: { project_id: f.plan.projectId, plan_revision: 1 },
  })
  await expect(confirm(f)).rejects.toThrow('presentation_team_unavailable')
})
it('previews a bounded explicit summary while publishing the exact large legal saved plan', async () => {
  const f = await fixture()
  const source = f.plan.sources[0]!
  f.plan.sources = Array.from({ length: 7 }, (_, i) => ({
    ...source,
    id: i === 0 ? source.id : `source-${i}`,
    excerpt: 'Original'.repeat(1500),
  }))
  const r = await f.skill.executeTool({
    id: 'large',
    name: 'create_presentation_team',
    input: { project_id: f.plan.projectId, plan_revision: 1 },
  })
  expect(r.isError, r.output).not.toBe(true)
  const preview = f.proposals.pending()!.preview
  expect(preview.plan).toBeUndefined()
  expect(preview.planSummary).toMatchObject({ sourceCount: 7, excerptSummaryOnly: true })
  expect(new TextEncoder().encode(JSON.stringify(preview)).length).toBeLessThan(48 * 1024)
  await confirm(f)
  expect(f.teamRequest.mock.calls.some(([b]) => b.operation === 'team_project_create')).toBe(true)
})
it('binds a mutation to the preview account as a server precondition', async () => {
  const f = await fixture('reviewer')
  const original = f.teamRequest.getMockImplementation()!
  let writes = 0
  f.teamRequest.mockImplementation(async (body) => {
    if (body.operation === 'team_comment_add') {
      expect(body.expectedIdentity).toEqual({
        version: 1,
        actorSubject: 'b'.repeat(64),
        pcSubject: 'a'.repeat(64),
      })
      return Response.json({ error: 'access_denied' }, { status: 403 })
    }
    if (!['team_identity', 'team_project_read'].includes(body.operation)) writes++
    return original(body)
  })
  await f.skill.executeTool({
    id: 'comment',
    name: 'add_presentation_team_comment',
    input: {
      team_id: f.ledger.teamId,
      expected_revision: 1,
      comment: {
        id: 'comment',
        targetKind: 'slide',
        targetId: f.plan.slides[0]!.id,
        text: 'Review',
      },
    },
  })
  await expect(confirm(f)).rejects.toThrow('access_denied')
  expect(writes).toBe(0)
  expect(f.ledger.comments).toEqual([])
})
