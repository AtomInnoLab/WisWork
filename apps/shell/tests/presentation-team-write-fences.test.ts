import { afterEach, expect, it } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PresentationLifecycleStore, PresentationStore } from '@wiswork/project-store'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationTeamService, PresentationTeamStore } from '../src/main/presentation-team'
import {
  capturePresentationProjectReadLease,
  capturePresentationProjectWriteLease,
} from '../src/main/presentation-project-write-lease'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const owner = 'a'.repeat(64),
  reviewer = 'b'.repeat(64)
const context = (actor = owner) => ({ version: 1 as const, pcSubject: owner, actorSubject: actor })
async function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'team-write-fences-'))
  roots.push(root)
  const store = new PresentationStore(root),
    lifecycle = new PresentationLifecycleStore(root),
    teams = new PresentationTeamStore(root),
    plan = benchmarkPlan(),
    scope = { documentId: 'doc', projectId: plan.projectId }
  store.savePlan(scope.projectId, scope.documentId, 0, plan)
  lifecycle.initialize(scope)
  let gate: Promise<void> | undefined,
    release = () => {},
    locks = 0
  const factory = () =>
    createPresentationTeamService({
      userDataPath: root,
      readPlan: (doc, project) => store.plan(project, doc),
      acquireProjectLock: async () => {
        locks++
        if (gate) await gate
        return () => {}
      },
      captureProjectLease: (actual, mode, signal) => {
        const options = {
          store: lifecycle,
          scope: actual,
          signal,
          readExistingProject: (s: typeof scope) => store.projectScope(s.projectId, s.documentId),
        }
        if (mode === 'read') return capturePresentationProjectReadLease(options)
        const lease = capturePresentationProjectWriteLease(options)
        return { assertCurrent: lease.assertWritable }
      },
    })
  const service = factory(),
    call = (body: Record<string, unknown>, actor = owner) =>
      service(
        {
          ...body,
          ...(!['team_identity', 'team_project_read', 'team_plan_read'].includes(
            String(body.operation),
          )
            ? { expectedIdentity: context(actor) }
            : {}),
        },
        context(actor),
        new AbortController().signal,
      )
  const created = await call({ operation: 'team_project_create', ...scope, planRevision: 1 })
  if (!('team' in created)) throw Error('fixture')
  const teamId = created.team.teamId
  return {
    root,
    store,
    lifecycle,
    teams,
    plan,
    scope,
    service,
    factory,
    call,
    teamId,
    pause() {
      gate = new Promise<void>((done) => {
        release = done
      })
    },
    release() {
      release()
      gate = undefined
    },
    locks: () => locks,
    freeze() {
      lifecycle.beginDeletion(scope, 0, {
        deletionId: 'delete',
        reason: 'user',
        resources: [{ resourceId: 'team', kind: 'teams', ownership: 'project_exclusive' }],
      })
    },
  }
}
it('rejects frozen writes across service instances without changing authorized team bytes', async () => {
  const f = await fixture(),
    before = f.teams.read(f.teamId)!
  f.freeze()
  await expect(
    f.factory()(
      {
        operation: 'team_member_set',
        teamId: f.teamId,
        expectedRevision: 1,
        memberSubject: reviewer,
        role: 'reviewer',
        expectedIdentity: context(),
      },
      context(),
      new AbortController().signal,
    ),
  ).rejects.toThrow('project_deleting')
  expect(f.teams.read(f.teamId)).toEqual(before)
})
it('captures the original lifecycle revision before member writes wait on the project lock', async () => {
  const f = await fixture(),
    before = f.teams.read(f.teamId)!,
    locks = f.locks()
  f.pause()
  const pending = f.call({
    operation: 'team_member_set',
    teamId: f.teamId,
    expectedRevision: 1,
    memberSubject: reviewer,
    role: 'reviewer',
  })
  await Promise.resolve()
  const entered = f.locks() > locks
  f.freeze()
  f.release()
  await expect(pending).rejects.toThrow('revision_conflict')
  expect(entered).toBe(true)
  expect(f.teams.read(f.teamId)).toEqual(before)
})
it('reauthorizes a queued reviewer after another instance revokes the actual ledger ACL', async () => {
  const f = await fixture()
  await f.call({
    operation: 'team_member_set',
    teamId: f.teamId,
    expectedRevision: 1,
    memberSubject: reviewer,
    role: 'reviewer',
  })
  f.pause()
  const pending = f.call(
    {
      operation: 'team_comment_add',
      teamId: f.teamId,
      expectedRevision: 2,
      planRevision: 1,
      comment: {
        id: 'queued',
        targetKind: 'slide',
        targetId: f.plan.slides[0]!.id,
        text: 'review',
      },
    },
    reviewer,
  )
  await Promise.resolve()
  const before = f.teams.read(f.teamId)!
  f.teams.write({ ...before, revision: 3, members: [] }, before)
  f.release()
  await expect(pending).rejects.toThrow('access_denied')
  expect(f.teams.read(f.teamId)?.comments).toEqual([])
})
it('owns body and trusted identity before publication awaits', async () => {
  const f = await fixture(),
    ctx = context(),
    body = {
      operation: 'team_plan_publish',
      teamId: f.teamId,
      expectedRevision: 1,
      planRevision: 1,
      expectedIdentity: context(),
    }
  f.pause()
  const pending = f.service(body, ctx, new AbortController().signal)
  await Promise.resolve()
  body.planRevision = 999
  ctx.actorSubject = reviewer
  f.release()
  const result = await pending
  expect('team' in result && result.team.publishedPlan.revision).toBe(1)
})
it('checks the store guard before namespace, temporary write and final publication', async () => {
  const f = await fixture(),
    before = f.teams.read(f.teamId)!
  let checks = 0
  expect(() =>
    f.teams.write({ ...before, revision: 2 }, before, () => {
      if (++checks === 3) throw Error('project_deleting')
    }),
  ).toThrow('project_deleting')
  expect(checks).toBe(3)
  expect(f.teams.read(f.teamId)).toEqual(before)
})
it('identity reads do not ask for project ownership or initialize lifecycle metadata', async () => {
  const root = mkdtempSync(join(tmpdir(), 'team-identity-fence-'))
  roots.push(root)
  const service = createPresentationTeamService({
    userDataPath: root,
    readPlan: () => {
      throw Error('unexpected')
    },
    acquireProjectLock: async () => {
      throw Error('unexpected')
    },
    captureProjectLease: () => {
      throw Error('unexpected')
    },
  })
  expect(
    await service({ operation: 'team_identity' }, context(), new AbortController().signal),
  ).toEqual({ identity: context() })
  expect(
    new PresentationLifecycleStore(root).read({ documentId: 'doc', projectId: 'p' }),
  ).toBeUndefined()
})
it('retains the original read lease while private plan preview waits', async () => {
  const f = await fixture()
  f.pause()
  const pending = f.service(
    { operation: 'team_plan_read', ...f.scope, planRevision: 1 },
    context(),
    new AbortController().signal,
  )
  await Promise.resolve()
  f.freeze()
  f.release()
  await expect(pending).rejects.toThrow('revision_conflict')
})
it('rechecks cancellation after an acquired member lock without changing the ledger', async () => {
  const f = await fixture(),
    before = f.teams.read(f.teamId),
    controller = new AbortController()
  f.pause()
  const pending = f.service(
    {
      operation: 'team_member_set',
      teamId: f.teamId,
      expectedRevision: 1,
      memberSubject: reviewer,
      role: 'reviewer',
      expectedIdentity: context(),
    },
    context(),
    controller.signal,
  )
  await Promise.resolve()
  controller.abort()
  f.release()
  await expect(pending).rejects.toThrow('aborted')
  expect(f.teams.read(f.teamId)).toEqual(before)
})
