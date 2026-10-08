import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it, vi } from 'vitest'
import { createPresentationService } from '../../shell/src/main/presentation-service.js'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { createOfficeHostRuntime } from '../src/agent/host-runtime.js'
import { createOfficeAgentSession } from '../src/agent/use-office-agent.js'
import type { StructuredProposalController } from '../src/agent/proposal-controller.js'
import { createOfficeWorkspaceUi } from '../src/App.js'

it('uses the actual workbench controller and existing confirmation card to publish, review and revoke without model calls or host writes', async () => {
  const root = mkdtempSync(join(tmpdir(), 'team-workbench-integration-'))
  const service = createPresentationService({ userDataPath: root })
  const owner = 'a'.repeat(64),
    reviewer = 'b'.repeat(64)
  let actor = owner
  const request = async (body: unknown, signal?: AbortSignal, team = false) => {
    const raw = Buffer.from(
      await service(
        body,
        signal ?? new AbortController().signal,
        team ? { version: 1, actorSubject: actor, pcSubject: owner } : undefined,
      ),
    ).toString('utf8')
    return new Response(raw, { status: JSON.parse(raw).error ? 400 : 200 })
  }
  const invalidateQa = vi.fn(async () => {})
  const stream = vi.fn()
  vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => true } } })
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      available: () => true,
      documentId: async () => 'workbench-document',
      request,
      teamAvailable: () => true,
      teamRequest: (body, signal) => request(body, signal, true),
      invalidateQa,
      lastProject: () => benchmarkPlan().projectId,
      rememberProject: async () => {},
    },
  })
  const session = createOfficeAgentSession({
    host: 'powerpoint',
    transport: { stream },
    skill: runtime.skill,
    proposals: runtime.proposals,
  })
  const ui = createOfficeWorkspaceUi(runtime)
  const team = ui.team!
  const proposals = runtime.proposals as StructuredProposalController
  const confirm = async (id: string) => {
    await vi.waitFor(() => expect(proposals.pending()?.lockReview?.state).not.toBe('checking'))
    expect(session.snapshot().proposal?.id).toBe(id)
    await session.confirm(id)
    expect(session.snapshot().error).toBeUndefined()
    await vi.waitFor(() => {
      expect(team.snapshot().proposalId).toBeUndefined()
      expect(team.snapshot().phase).toBe('idle')
      expect(team.snapshot().team).toBeDefined()
      expect(team.snapshot().error).toBeUndefined()
    })
    expect(session.snapshot().activity).toContain('团队')
  }
  try {
    const plan = benchmarkPlan()
    await request({
      operation: 'save_plan',
      documentId: 'workbench-document',
      projectId: plan.projectId,
      expectedRevision: 0,
      plan,
    })
    await team.refresh()
    expect(team.snapshot().identity?.actorSubject).toBe(owner)
    const create = await team.create(plan.projectId, 1)
    expect(team.snapshot().team).toBeUndefined()
    expect(session.snapshot().proposal?.id).toBe(create)
    await confirm(create)
    await vi.waitFor(() => expect(team.snapshot().team?.publishedPlan.plan).toEqual(plan))
    const teamId = team.snapshot().team!.teamId
    await confirm(await team.setMember(reviewer, 'reviewer'))
    expect(team.snapshot().team!.members).toEqual([{ subject: reviewer, role: 'reviewer' }])
    actor = reviewer
    await team.refresh(teamId)
    expect(team.snapshot().role).toBe('reviewer')
    const comment = await team.addComment('source', plan.sources[0]!.id, '请核对这条公开来源')
    expect(team.snapshot().team!.comments).toEqual([])
    session.reject()
    await vi.waitFor(() => expect(team.snapshot().proposalId).toBeUndefined())
    expect(team.snapshot().team).toBeUndefined()
    await team.refresh(teamId)
    expect(team.snapshot().team!.comments).toEqual([])
    expect(proposals.pending()).toBeUndefined()
    expect(comment).toBeTruthy()
    await confirm(await team.addComment('source', plan.sources[0]!.id, '再次提请核对'))
    expect(team.snapshot().team!.comments[0]).toMatchObject({
      authorSubject: reviewer,
      planRevision: 1,
    })
    actor = owner
    await team.refresh(teamId)
    await confirm(await team.revokeMember(reviewer))
    actor = reviewer
    await team.refresh(teamId)
    expect(team.snapshot().team).toBeUndefined()
    expect(team.snapshot().role).toBeUndefined()
    expect(team.snapshot().error).toBeTruthy()
    expect(invalidateQa).not.toHaveBeenCalled()
    expect(stream).not.toHaveBeenCalled()
    runtime.clearSession()
    expect(team.snapshot().identity).toBeUndefined()
  } finally {
    session.dispose()
    runtime.dispose()
    vi.unstubAllGlobals()
    rmSync(root, { recursive: true, force: true })
  }
})
