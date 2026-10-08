import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it, vi } from 'vitest'
import { createPresentationService } from '../../shell/src/main/presentation-service.js'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { createPresentationTeamSkill } from '../src/skills/powerpoint/presentation-team.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import type { PresentationTeamContext } from '@wiswork/pptx-engine/presentation-team'

// Synthetic verified principals exercise real PC persistence and Office confirmation;
// Relay's local OIDC tests separately verify production context construction.
it('shares only explicitly confirmed snapshots, binds reviewer identity and enforces revocation after reopening', async () => {
  const root = mkdtempSync(join(tmpdir(), 'office-team-integration-'))
  const owner = 'a'.repeat(64),
    reviewer = 'b'.repeat(64),
    viewer = 'c'.repeat(64)
  let actor = owner
  let service = createPresentationService({ userDataPath: root })
  const context = (): PresentationTeamContext => ({
    version: 1,
    actorSubject: actor,
    pcSubject: owner,
  })
  const invoke = async (body: unknown, team = false, signal?: AbortSignal) => {
    const bytes = await service(
      body,
      signal ?? new AbortController().signal,
      team ? context() : undefined,
    )
    const raw = Buffer.from(bytes).toString('utf8')
    return new Response(raw, { status: JSON.parse(raw).error ? 400 : 200 })
  }
  const raw = async (body: unknown) => (await invoke(body, true)).json()
  const proposals = createStructuredProposalController()
  const plan = benchmarkPlan()
  const skill = createPresentationTeamSkill({
    teamAvailable: () => true,
    teamRequest: (body, signal) => invoke(body, true, signal),
    request: (body, signal) => invoke(body, false, signal),
    documentId: async () => 'owner-document',
    proposals,
  })
  let nextId = 0
  const execute = (name: string, input: Record<string, unknown>) =>
    skill.executeTool({ id: String(++nextId), name, input })
  const confirm = async (result: { output: string; isError?: boolean }) => {
    expect(result.isError, result.output).not.toBe(true)
    const id = JSON.parse(result.output).proposalId
    expect(id).toBeTruthy()
    await vi.waitFor(() => expect(proposals.pending()?.lockReview?.state).not.toBe('checking'))
    const decision = proposals.waitForDecision(id)
    await proposals.confirm(id)
    expect((await decision).status).toBe('confirmed')
  }
  try {
    expect(
      (
        await (
          await invoke({
            operation: 'save_plan',
            documentId: 'owner-document',
            projectId: plan.projectId,
            expectedRevision: 0,
            plan,
          })
        ).json()
      ).revision,
    ).toBe(1)
    const created = await execute('create_presentation_team', {
      project_id: plan.projectId,
      plan_revision: 1,
    })
    // A visible proposal is not publication: no ACL/content ledger exists yet.
    const { presentationTeamId } = await import('@wiswork/pptx-engine/presentation-team')
    const teamId = await presentationTeamId(owner, 'owner-document', plan.projectId)
    expect(await raw({ operation: 'team_project_read', teamId })).toEqual({ error: 'not_found' })
    await confirm(created)
    let ledger = (await raw({ operation: 'team_project_read', teamId })).team
    expect(ledger.publishedPlan.plan).toEqual(plan)
    for (const [member, role] of [
      [reviewer, 'reviewer'],
      [viewer, 'viewer'],
    ] as const) {
      await confirm(
        await execute('set_presentation_team_member', {
          team_id: teamId,
          expected_revision: ledger.revision,
          member_subject: member,
          role,
        }),
      )
      ledger = (await raw({ operation: 'team_project_read', teamId })).team
    }
    actor = reviewer
    const comment = {
      id: 'review-source',
      targetKind: 'source',
      targetId: plan.sources[0]!.id,
      text: '核对这条来源的摘录',
    }
    const proposed = await execute('add_presentation_team_comment', {
      team_id: teamId,
      expected_revision: ledger.revision,
      comment,
    })
    expect((await raw({ operation: 'team_project_read', teamId })).team.comments).toHaveLength(0)
    await confirm(proposed)
    ledger = (await raw({ operation: 'team_project_read', teamId })).team
    expect(ledger.comments[0]).toMatchObject({
      ...comment,
      authorSubject: reviewer,
      planRevision: 1,
    })
    actor = viewer
    expect(
      (
        await execute('add_presentation_team_comment', {
          team_id: teamId,
          expected_revision: ledger.revision,
          comment: { ...comment, id: 'viewer-comment' },
        })
      ).isError,
    ).toBe(true)
    expect((await raw({ operation: 'team_project_read', teamId })).team.comments).toHaveLength(1)
    actor = owner
    const changed = structuredClone(plan)
    changed.sources[0]!.excerpt = '仅私人计划修改，尚未发布'
    expect(
      (
        await (
          await invoke({
            operation: 'save_plan',
            documentId: 'owner-document',
            projectId: plan.projectId,
            expectedRevision: 1,
            plan: changed,
          })
        ).json()
      ).revision,
    ).toBe(2)
    service = createPresentationService({ userDataPath: root })
    actor = reviewer
    const shared = await execute('read_presentation_team', { team_id: teamId })
    expect(shared.isError, shared.output).not.toBe(true)
    expect(
      (await raw({ operation: 'team_project_read', teamId })).team.publishedPlan.plan.sources[0]
        .excerpt,
    ).toBe(plan.sources[0]!.excerpt)
    actor = owner
    await confirm(
      await execute('publish_presentation_team_plan', {
        team_id: teamId,
        expected_revision: ledger.revision,
        plan_revision: 2,
      }),
    )
    ledger = (await raw({ operation: 'team_project_read', teamId })).team
    expect(ledger.publishedPlan.plan).toEqual(changed)
    expect(ledger.comments[0].planRevision).toBe(1)
    await confirm(
      await execute('revoke_presentation_team_member', {
        team_id: teamId,
        expected_revision: ledger.revision,
        member_subject: reviewer,
      }),
    )
    actor = reviewer
    expect((await execute('read_presentation_team', { team_id: teamId })).isError).toBe(true)
    expect(await raw({ operation: 'team_project_read', teamId })).toEqual({
      error: 'access_denied',
    })
  } finally {
    skill.clear()
    rmSync(root, { recursive: true, force: true })
  }
})

it('composes team proposals in the actual runtime without invalidating document QA', async () => {
  const { createOfficeHostRuntime } = await import('../src/agent/host-runtime.js')
  const root = mkdtempSync(join(tmpdir(), 'team-runtime-qa-'))
  let runtime: ReturnType<typeof createOfficeHostRuntime> | undefined
  const service = createPresentationService({ userDataPath: root })
  const owner = 'd'.repeat(64)
  const request = async (body: unknown, signal?: AbortSignal, team = false) => {
    const raw = Buffer.from(
      await service(
        body,
        signal ?? new AbortController().signal,
        team ? { version: 1, actorSubject: owner, pcSubject: owner } : undefined,
      ),
    ).toString('utf8')
    return new Response(raw, { status: JSON.parse(raw).error ? 400 : 200 })
  }
  const invalidateQa = vi.fn(async () => {})
  vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => true } } })
  try {
    const plan = benchmarkPlan()
    await request({
      operation: 'save_plan',
      documentId: 'runtime-document',
      projectId: plan.projectId,
      expectedRevision: 0,
      plan,
    })
    runtime = createOfficeHostRuntime('powerpoint', {
      presentation: {
        available: () => true,
        documentId: async () => 'runtime-document',
        request,
        teamAvailable: () => true,
        teamRequest: (body, signal) => request(body, signal, true),
        lastProject: () => plan.projectId,
        rememberProject: async () => {},
        invalidateQa,
      },
    })
    expect(runtime.skill.tools.map((t) => t.name)).toContain('create_presentation_team')
    const result = await runtime.skill.executeTool({
      id: 'create-team',
      name: 'create_presentation_team',
      input: { project_id: plan.projectId, plan_revision: 1 },
    })
    expect(result.isError, result.output).not.toBe(true)
    const id = JSON.parse(result.output).proposalId
    const proposals = runtime.proposals as ReturnType<typeof createStructuredProposalController>
    await vi.waitFor(() => expect(proposals.pending()?.lockReview?.state).not.toBe('checking'))
    const decision = proposals.waitForDecision(id)
    await proposals.confirm(id)
    expect((await decision).status).toBe('confirmed')
    expect(invalidateQa).not.toHaveBeenCalled()
    runtime.clearSession()
    expect(runtime.proposals.pending()).toBeUndefined()
  } finally {
    runtime?.dispose()
    vi.unstubAllGlobals()
    rmSync(root, { recursive: true, force: true })
  }
})

it('previews sharing a legal large source plan within the proposal budget and persists its complete snapshot', async () => {
  const root = mkdtempSync(join(tmpdir(), 'team-large-snapshot-'))
  const owner = 'e'.repeat(64)
  const service = createPresentationService({ userDataPath: root })
  const plan = benchmarkPlan()
  plan.sources = Array.from({ length: 7 }, (_, i) => ({
    ...plan.sources[0]!,
    id: i ? `large-source-${i}` : 'source',
    excerpt: 'a'.repeat(12000),
  }))
  const request = async (body: unknown, signal?: AbortSignal, team = false) => {
    const raw = Buffer.from(
      await service(
        body,
        signal ?? new AbortController().signal,
        team ? { version: 1, actorSubject: owner, pcSubject: owner } : undefined,
      ),
    ).toString('utf8')
    return new Response(raw, { status: JSON.parse(raw).error ? 400 : 200 })
  }
  const proposals = createStructuredProposalController()
  const skill = createPresentationTeamSkill({
    documentId: async () => 'large-document',
    request,
    teamAvailable: () => true,
    teamRequest: (body, signal) => request(body, signal, true),
    proposals,
  })
  try {
    expect(
      (
        await (
          await request({
            operation: 'save_plan',
            documentId: 'large-document',
            projectId: plan.projectId,
            expectedRevision: 0,
            plan,
          })
        ).json()
      ).revision,
    ).toBe(1)
    expect(Buffer.byteLength(JSON.stringify(plan))).toBeGreaterThan(64 * 1024)
    const result = await skill.executeTool({
      id: 'large-plan-share',
      name: 'create_presentation_team',
      input: { project_id: plan.projectId, plan_revision: 1 },
    })
    expect(result.isError, result.output).not.toBe(true)
    expect(Buffer.byteLength(JSON.stringify(proposals.pending()?.preview))).toBeLessThan(64 * 1024)
    const id = JSON.parse(result.output).proposalId
    const decision = proposals.waitForDecision(id)
    await proposals.confirm(id)
    expect((await decision).status).toBe('confirmed')
    const { presentationTeamId } = await import('@wiswork/pptx-engine/presentation-team')
    const teamId = await presentationTeamId(owner, 'large-document', plan.projectId)
    const saved = await (
      await request({ operation: 'team_project_read', teamId }, undefined, true)
    ).json()
    expect(saved.team.publishedPlan.plan).toEqual(plan)
    expect(saved.team.publishedPlan.plan.sources[6].excerpt).toHaveLength(12000)
  } finally {
    skill.clear()
    rmSync(root, { recursive: true, force: true })
  }
})

it('rejects an account switch at the actual write even when both accounts are authorized reviewers', async () => {
  const root = mkdtempSync(join(tmpdir(), 'team-account-switch-'))
  const owner = 'a'.repeat(64),
    first = 'b'.repeat(64),
    second = 'c'.repeat(64)
  let actor = owner
  let switchAtWrite = false
  const service = createPresentationService({ userDataPath: root })
  const identity = (): PresentationTeamContext => ({
    version: 1,
    actorSubject: actor,
    pcSubject: owner,
  })
  const request = async (body: unknown, signal?: AbortSignal, team = false) => {
    if (switchAtWrite && (body as { operation: string }).operation === 'team_comment_add')
      actor = second
    const raw = Buffer.from(
      await service(body, signal ?? new AbortController().signal, team ? identity() : undefined),
    ).toString('utf8')
    return new Response(raw, { status: JSON.parse(raw).error ? 400 : 200 })
  }
  const direct = async (body: Record<string, unknown>, write = false) =>
    (
      await request(write ? { ...body, expectedIdentity: identity() } : body, undefined, true)
    ).json()
  const proposals = createStructuredProposalController()
  const plan = benchmarkPlan()
  const skill = createPresentationTeamSkill({
    documentId: async () => 'switch-document',
    request,
    teamAvailable: () => true,
    teamRequest: (body, signal) => request(body, signal, true),
    proposals,
  })
  try {
    await request({
      operation: 'save_plan',
      documentId: 'switch-document',
      projectId: plan.projectId,
      expectedRevision: 0,
      plan,
    })
    let ledger = (
      await direct(
        {
          operation: 'team_project_create',
          documentId: 'switch-document',
          projectId: plan.projectId,
          planRevision: 1,
        },
        true,
      )
    ).team
    for (const memberSubject of [first, second])
      ledger = (
        await direct(
          {
            operation: 'team_member_set',
            teamId: ledger.teamId,
            expectedRevision: ledger.revision,
            memberSubject,
            role: 'reviewer',
          },
          true,
        )
      ).team
    actor = first
    const result = await skill.executeTool({
      id: 'review-account-switch',
      name: 'add_presentation_team_comment',
      input: {
        team_id: ledger.teamId,
        expected_revision: ledger.revision,
        comment: {
          id: 'never-written',
          targetKind: 'slide',
          targetId: plan.slides[0]!.id,
          text: '这条评论不应以另一个账号保存',
        },
      },
    })
    expect(result.isError, result.output).not.toBe(true)
    switchAtWrite = true
    await expect(proposals.confirm(JSON.parse(result.output).proposalId)).rejects.toThrow()
    actor = owner
    const saved = (await direct({ operation: 'team_project_read', teamId: ledger.teamId })).team
    expect(saved.comments).toEqual([])
    expect(saved.revision).toBe(ledger.revision)
  } finally {
    skill.clear()
    rmSync(root, { recursive: true, force: true })
  }
})

it('lets only the authenticated owner inspect private plans through the isolated team capability', async () => {
  const root = mkdtempSync(join(tmpdir(), 'team-private-plan-read-'))
  const service = createPresentationService({ userDataPath: root })
  const plan = benchmarkPlan(),
    owner = 'a'.repeat(64),
    reviewer = 'b'.repeat(64)
  const call = async (body: unknown, actor?: string) =>
    JSON.parse(
      Buffer.from(
        await service(
          body,
          new AbortController().signal,
          actor ? { version: 1, actorSubject: actor, pcSubject: owner } : undefined,
        ),
      ).toString('utf8'),
    )
  try {
    await call({
      operation: 'save_plan',
      documentId: 'private-document',
      projectId: plan.projectId,
      expectedRevision: 0,
      plan,
    })
    const read = {
      operation: 'team_plan_read',
      documentId: 'private-document',
      projectId: plan.projectId,
      planRevision: 1,
    }
    expect(await call(read)).toEqual({ error: 'access_denied' })
    expect(await call(read, reviewer)).toEqual({ error: 'access_denied' })
    expect(await call({ ...read, documentId: 'missing-document' }, reviewer)).toEqual({
      error: 'access_denied',
    })
    expect(await call(read, owner)).toEqual({ projectId: plan.projectId, revision: 1, plan })
    expect(
      await call(
        { ...read, expectedIdentity: { version: 1, actorSubject: owner, pcSubject: owner } },
        reviewer,
      ),
    ).toEqual({ error: 'invalid_request' })
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})
