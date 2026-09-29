import { afterEach, expect, it } from 'vitest'
import { mkdtempSync, rmSync, readdirSync, writeFileSync, symlinkSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createPresentationService } from '../src/main/presentation-service'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import type { PresentationTeamContext } from '@wiswork/pptx-engine/presentation-team'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const owner = 'a'.repeat(64),
  reviewer = 'b'.repeat(64),
  viewer = 'c'.repeat(64)
const context = (actor = owner, pc = owner): PresentationTeamContext => ({
  version: 1,
  actorSubject: actor,
  pcSubject: pc,
})
async function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'team-pc-'))
  roots.push(root)
  let service = createPresentationService({ userDataPath: root })
  const plan = benchmarkPlan(),
    doc = 'doc'
  const call = async (body: unknown, ctx?: PresentationTeamContext) => {
    const request =
      ctx &&
      body &&
      typeof body === 'object' &&
      !Array.isArray(body) &&
      typeof (body as { operation?: unknown }).operation === 'string' &&
      (body as { operation: string }).operation.startsWith('team_') &&
      !['team_identity', 'team_project_read', 'team_plan_read'].includes(
        (body as { operation: string }).operation,
      ) &&
      !Object.hasOwn(body, 'expectedIdentity')
        ? { ...body, expectedIdentity: ctx }
        : body
    return JSON.parse(
      Buffer.from(await service(request, new AbortController().signal, ctx)).toString('utf8'),
    )
  }
  expect(
    (
      await call({
        operation: 'save_plan',
        documentId: doc,
        projectId: plan.projectId,
        expectedRevision: 0,
        plan,
      })
    ).revision,
  ).toBe(1)
  const create = {
    operation: 'team_project_create',
    documentId: doc,
    projectId: plan.projectId,
    planRevision: 1,
  }
  return {
    root,
    plan,
    doc,
    call,
    create,
    restart: () => {
      service = createPresentationService({ userDataPath: root })
    },
  }
}
it('requires trusted identity, owner-only publication and keeps private changes out of member reads', async () => {
  const f = await fixture()
  expect(await f.call({ operation: 'team_identity' })).toEqual({ error: 'access_denied' })
  expect(await f.call({ operation: 'team_identity' }, context(reviewer))).toEqual({
    identity: context(reviewer),
  })
  expect(await f.call(f.create, context(reviewer))).toEqual({ error: 'access_denied' })
  const created = await f.call(f.create, context())
  expect(created.team).toMatchObject({
    ownerSubject: owner,
    revision: 1,
    publishedPlan: { revision: 1, plan: f.plan },
    members: [],
    comments: [],
  })
  expect(await f.call(f.create, context())).toEqual(created)
  const teamId = created.team.teamId
  expect(await f.call({ operation: 'team_project_read', teamId }, context(reviewer))).toEqual({
    error: 'access_denied',
  })
  expect(
    (
      await f.call(
        {
          operation: 'team_member_set',
          teamId,
          expectedRevision: 1,
          memberSubject: reviewer,
          role: 'reviewer',
        },
        context(),
      )
    ).team.revision,
  ).toBe(2)
  const changed = structuredClone(f.plan)
  changed.brief.objective = 'Private changed objective'
  expect(
    (
      await f.call({
        operation: 'save_plan',
        documentId: f.doc,
        projectId: f.plan.projectId,
        expectedRevision: 1,
        plan: changed,
      })
    ).revision,
  ).toBe(2)
  f.restart()
  expect(
    (await f.call({ operation: 'team_project_read', teamId }, context(reviewer))).team.publishedPlan
      .plan,
  ).toEqual(f.plan)
  expect(
    await f.call({ operation: 'team_project_read', teamId }, context(reviewer, 'd'.repeat(64))),
  ).toEqual({ error: 'access_denied' })
  expect(
    (
      await f.call(
        { operation: 'team_plan_publish', teamId, expectedRevision: 2, planRevision: 2 },
        context(),
      )
    ).team.publishedPlan.plan,
  ).toEqual(changed)
})
it('enforces roles and immutable authenticated authors, preserves comment ACK retries and revokes reads', async () => {
  const f = await fixture()
  const { team } = await f.call(f.create, context()),
    teamId = team.teamId
  await f.call(
    {
      operation: 'team_member_set',
      teamId,
      expectedRevision: 1,
      memberSubject: reviewer,
      role: 'reviewer',
    },
    context(),
  )
  await f.call(
    {
      operation: 'team_member_set',
      teamId,
      expectedRevision: 2,
      memberSubject: viewer,
      role: 'viewer',
    },
    context(),
  )
  const comment = {
    id: 'comment',
    targetKind: 'slide',
    targetId: f.plan.slides[0]!.id,
    text: 'Review this',
  }
  const add = {
    operation: 'team_comment_add',
    teamId,
    expectedRevision: 3,
    planRevision: 1,
    comment,
  }
  expect(await f.call(add, context(viewer))).toEqual({ error: 'access_denied' })
  const result = await f.call(add, context(reviewer))
  expect(result.team.comments[0]).toMatchObject({
    ...comment,
    authorSubject: reviewer,
    planRevision: 1,
    state: 'open',
  })
  f.restart()
  expect(await f.call(add, context(reviewer))).toEqual(result)
  expect(
    await f.call({ ...add, comment: { ...comment, text: 'Different' } }, context(reviewer)),
  ).toEqual({ error: 'revision_conflict' })
  expect(
    await f.call(
      { ...add, comment: { ...comment, id: 'forged', authorSubject: owner } },
      context(reviewer),
    ),
  ).toEqual({ error: 'invalid_request' })
  expect(
    await f.call(
      {
        operation: 'team_member_set',
        teamId,
        expectedRevision: 4,
        memberSubject: owner,
        role: 'viewer',
      },
      context(),
    ),
  ).toEqual({ error: 'invalid_request' })
  expect(
    (
      await f.call(
        { operation: 'team_comment_resolve', teamId, expectedRevision: 4, commentId: 'comment' },
        context(reviewer),
      )
    ).team.comments[0].state,
  ).toBe('resolved')
  await f.call(
    { operation: 'team_member_revoke', teamId, expectedRevision: 5, memberSubject: reviewer },
    context(),
  )
  expect(await f.call({ operation: 'team_project_read', teamId }, context(reviewer))).toEqual({
    error: 'access_denied',
  })
  expect(await f.call(add, context(reviewer))).toEqual({ error: 'access_denied' })
})
it('rejects forged context in Agent body, stale membership CAS, missing targets, corrupt/symlink stores', async () => {
  const f = await fixture()
  expect(await f.call({ ...f.create, context: context() }, context())).toEqual({
    error: 'invalid_request',
  })
  const { team } = await f.call(f.create, context()),
    teamId = team.teamId
  expect(
    await f.call(
      {
        operation: 'team_member_set',
        teamId,
        expectedRevision: 0,
        memberSubject: reviewer,
        role: 'reviewer',
      },
      context(),
    ),
  ).toEqual({ error: 'revision_conflict' })
  expect(
    await f.call(
      {
        operation: 'team_comment_add',
        teamId,
        expectedRevision: 1,
        planRevision: 1,
        comment: {
          id: 'bad',
          targetKind: 'source',
          targetId: 'missing',
          text: 'not a real source',
        },
      },
      context(),
    ),
  ).toEqual({ error: 'invalid_request' })
  const dir = join(f.root, 'presentation-teams'),
    file = join(dir, readdirSync(dir)[0]!)
  const raw = readFileSync(file)
  writeFileSync(file, 'bad-json')
  expect(await f.call({ operation: 'team_project_read', teamId }, context())).toEqual({
    error: 'invalid_state',
  })
  rmSync(file)
  writeFileSync(file + '.real', raw)
  symlinkSync(file + '.real', file)
  expect(await f.call({ operation: 'team_project_read', teamId }, context())).toEqual({
    error: 'invalid_state',
  })
})
it('allows reviewers to resolve only their own comments, owners any comment, and rejects expired members', async () => {
  const f = await fixture(),
    { team } = await f.call(f.create, context()),
    teamId = team.teamId,
    other = 'd'.repeat(64)
  await f.call(
    {
      operation: 'team_member_set',
      teamId,
      expectedRevision: 1,
      memberSubject: reviewer,
      role: 'reviewer',
    },
    context(),
  )
  await f.call(
    {
      operation: 'team_member_set',
      teamId,
      expectedRevision: 2,
      memberSubject: other,
      role: 'reviewer',
    },
    context(),
  )
  const comment = {
    id: 'own',
    targetKind: 'source',
    targetId: f.plan.sources[0]!.id,
    text: 'Source review is not certification',
  }
  await f.call(
    { operation: 'team_comment_add', teamId, expectedRevision: 3, planRevision: 1, comment },
    context(reviewer),
  )
  expect(
    await f.call(
      { operation: 'team_comment_resolve', teamId, expectedRevision: 4, commentId: 'own' },
      context(other),
    ),
  ).toEqual({ error: 'access_denied' })
  expect(
    await f.call(
      {
        operation: 'team_member_set',
        teamId,
        expectedRevision: 4,
        memberSubject: viewer,
        role: 'viewer',
      },
      context(other),
    ),
  ).toEqual({ error: 'access_denied' })
  expect(
    (
      await f.call(
        { operation: 'team_comment_resolve', teamId, expectedRevision: 4, commentId: 'own' },
        context(),
      )
    ).team.comments[0].state,
  ).toBe('resolved')
  expect(
    await f.call(
      { operation: 'team_member_revoke', teamId, expectedRevision: 5, memberSubject: owner },
      context(),
    ),
  ).toEqual({ error: 'invalid_request' })
})
it('retains old comment plan revision and publishes only exact saved sources after a separate owner action', async () => {
  const f = await fixture(),
    { team } = await f.call(f.create, context()),
    teamId = team.teamId
  await f.call(
    {
      operation: 'team_comment_add',
      teamId,
      expectedRevision: 1,
      planRevision: 1,
      comment: {
        id: 'historical',
        targetKind: 'source',
        targetId: f.plan.sources[0]!.id,
        text: 'Original excerpt',
      },
    },
    context(),
  )
  const changed = structuredClone(f.plan)
  changed.sources[0]!.excerpt = 'New private excerpt'
  await f.call({
    operation: 'save_plan',
    documentId: f.doc,
    projectId: f.plan.projectId,
    expectedRevision: 1,
    plan: changed,
  })
  expect(
    (await f.call({ operation: 'team_project_read', teamId }, context())).team.publishedPlan.plan
      .sources,
  ).toEqual(f.plan.sources)
  expect(
    await f.call(
      { operation: 'team_plan_publish', teamId, expectedRevision: 2, planRevision: 1 },
      context(),
    ),
  ).toEqual({ error: 'revision_conflict' })
  const published = await f.call(
    { operation: 'team_plan_publish', teamId, expectedRevision: 2, planRevision: 2 },
    context(),
  )
  expect(published.team.comments[0]).toMatchObject({ planRevision: 1, state: 'open' })
  expect(published.team.publishedPlan.plan.sources).toEqual(changed.sources)
  expect(
    await f.call(
      {
        operation: 'team_comment_add',
        teamId,
        expectedRevision: 3,
        planRevision: 1,
        comment: {
          id: 'stale',
          targetKind: 'source',
          targetId: f.plan.sources[0]!.id,
          text: 'Stale snapshot',
        },
      },
      context(),
    ),
  ).toEqual({ error: 'invalid_request' })
  expect(
    await f.call({ operation: 'team_project_read', teamId, documentId: 'other' }, context()),
  ).toEqual({ error: 'invalid_request' })
})
it('bounds membership and comment records without evicting existing review history', async () => {
  const f = await fixture(),
    { team } = await f.call(f.create, context()),
    teamId = team.teamId
  for (let i = 0; i < 32; i++)
    expect(
      (
        await f.call(
          {
            operation: 'team_member_set',
            teamId,
            expectedRevision: i + 1,
            memberSubject: (i + 1).toString(16).padStart(64, '0'),
            role: 'viewer',
          },
          context(),
        )
      ).team.members,
    ).toHaveLength(i + 1)
  expect(
    await f.call(
      {
        operation: 'team_member_set',
        teamId,
        expectedRevision: 33,
        memberSubject: 'f'.repeat(64),
        role: 'viewer',
      },
      context(),
    ),
  ).toEqual({ error: 'quota_exceeded' })
  for (let i = 0; i < 128; i++)
    expect(
      (
        await f.call(
          {
            operation: 'team_comment_add',
            teamId,
            expectedRevision: 33 + i,
            planRevision: 1,
            comment: {
              id: 'c' + i,
              targetKind: 'slide',
              targetId: f.plan.slides[0]!.id,
              text: 'Bounded history',
            },
          },
          context(),
        )
      ).team.comments,
    ).toHaveLength(i + 1)
  expect(
    await f.call(
      {
        operation: 'team_comment_add',
        teamId,
        expectedRevision: 161,
        planRevision: 1,
        comment: {
          id: 'over',
          targetKind: 'slide',
          targetId: f.plan.slides[0]!.id,
          text: 'No eviction',
        },
      },
      context(),
    ),
  ).toEqual({ error: 'quota_exceeded' })
  f.restart()
  expect(
    (await f.call({ operation: 'team_project_read', teamId }, context())).team.comments,
  ).toHaveLength(128)
})
it('enforces whole-ledger UTF8 bytes and fails malformed or cancelled requests before publication', async () => {
  const f = await fixture(),
    { team } = await f.call(f.create, context()),
    teamId = team.teamId
  let count = 0
  for (let i = 0; i < 128; i++) {
    const result = await f.call(
      {
        operation: 'team_comment_add',
        teamId,
        expectedRevision: i + 1,
        planRevision: 1,
        comment: {
          id: 'utf' + i,
          targetKind: 'slide',
          targetId: f.plan.slides[0]!.id,
          text: '界'.repeat(2000),
        },
      },
      context(),
    )
    if (result.error) {
      expect(result.error).toBe('quota_exceeded')
      break
    }
    count++
  }
  expect(count).toBeGreaterThan(0)
  expect(count).toBeLessThan(128)
  expect(
    await f.call({ operation: 'team_identity' }, { ...context(), actorSubject: 'unverified' }),
  ).toEqual({ error: 'access_denied' })
  const service = createPresentationService({ userDataPath: f.root }),
    signal = new AbortController()
  signal.abort()
  const result = JSON.parse(
    Buffer.from(await service(f.create, signal.signal, context())).toString('utf8'),
  )
  expect(result).toEqual({ error: 'aborted' })
  expect(
    (await f.call({ operation: 'team_project_read', teamId }, context())).team.comments,
  ).toHaveLength(count)
})
it('rejects missing or mismatched negative identity preconditions before any private read or publication', async () => {
  const f = await fixture()
  expect(await f.call({ ...f.create, expectedIdentity: undefined }, context())).toEqual({
    error: 'access_denied',
  })
  expect(await f.call({ ...f.create, expectedIdentity: context(reviewer) }, context())).toEqual({
    error: 'access_denied',
  })
  expect(
    await f.call(
      { ...f.create, expectedIdentity: { ...context(), authorLabel: 'fake' } },
      context(),
    ),
  ).toEqual({ error: 'access_denied' })
  expect(
    await f.call({ operation: 'team_identity', expectedIdentity: context() }, context()),
  ).toEqual({ error: 'invalid_request' })
  const dir = join(f.root, 'presentation-teams')
  expect(() => readdirSync(dir)).toThrow()
  const { team } = await f.call(f.create, context())
  expect(
    await f.call(
      { operation: 'team_project_read', teamId: team.teamId, expectedIdentity: context() },
      context(),
    ),
  ).toEqual({ error: 'invalid_request' })
})
it('denies a switched second authorized reviewer before writing even though their ACL permits comments', async () => {
  const f = await fixture(),
    { team } = await f.call(f.create, context()),
    teamId = team.teamId,
    other = 'd'.repeat(64)
  await f.call(
    {
      operation: 'team_member_set',
      teamId,
      expectedRevision: 1,
      memberSubject: reviewer,
      role: 'reviewer',
    },
    context(),
  )
  await f.call(
    {
      operation: 'team_member_set',
      teamId,
      expectedRevision: 2,
      memberSubject: other,
      role: 'reviewer',
    },
    context(),
  )
  const request = {
    operation: 'team_comment_add',
    teamId,
    expectedRevision: 3,
    planRevision: 1,
    expectedIdentity: context(reviewer),
    comment: {
      id: 'must-not-write',
      targetKind: 'slide',
      targetId: f.plan.slides[0]!.id,
      text: 'Confirmed as first reviewer',
    },
  }
  expect(await f.call(request, context(other))).toEqual({ error: 'access_denied' })
  const ledger = (await f.call({ operation: 'team_project_read', teamId }, context())).team
  expect(ledger.comments).toHaveLength(0)
  expect(ledger.revision).toBe(3)
  expect((await f.call(request, context(reviewer))).team.comments[0].authorSubject).toBe(reviewer)
})
it('allows only the authenticated paired owner to preview the exact current private plan over the isolated team route', async () => {
  const f = await fixture(),
    read = {
      operation: 'team_plan_read',
      documentId: f.doc,
      projectId: f.plan.projectId,
      planRevision: 1,
    }
  expect(await f.call(read, context())).toEqual({
    projectId: f.plan.projectId,
    revision: 1,
    plan: f.plan,
  })
  expect(await f.call(read, context(reviewer))).toEqual({ error: 'access_denied' })
  expect(await f.call(read, context(viewer))).toEqual({ error: 'access_denied' })
  expect(await f.call(read)).toEqual({ error: 'access_denied' })
  expect(await f.call({ ...read, expectedIdentity: context() }, context())).toEqual({
    error: 'invalid_request',
  })
  const changed = structuredClone(f.plan)
  changed.sources[0]!.excerpt = 'New private source not published'
  await f.call({
    operation: 'save_plan',
    documentId: f.doc,
    projectId: f.plan.projectId,
    expectedRevision: 1,
    plan: changed,
  })
  expect(await f.call(read, context())).toEqual({ error: 'revision_conflict' })
  expect(await f.call({ ...read, planRevision: 2 }, context())).toEqual({
    projectId: f.plan.projectId,
    revision: 2,
    plan: changed,
  })
  expect(
    await f.call({ ...read, documentId: 'different-doc', planRevision: 2 }, context()),
  ).toEqual({ error: 'document_mismatch' })
})
