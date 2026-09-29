import { expect, it } from 'vitest'
import { createPresentationTeamController } from '../src/agent/presentation-team-controller.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
it('keeps a stable hidden snapshot and performs no unavailable tool calls', async () => {
  let calls = 0
  const controller = createPresentationTeamController({
    skill: {
      id: 'team',
      tools: [],
      systemPrompt: '',
      async executeTool() {
        calls++
        throw Error('secret')
      },
    },
    documentId: async () => 'doc',
    proposals: createStructuredProposalController(),
  })
  expect(controller.snapshot()).toBe(controller.snapshot())
  expect(controller.snapshot().available).toBe(false)
  await controller.refresh()
  expect(calls).toBe(0)
  expect(controller.snapshot().error).toBe('presentation_team_unavailable')
})

import { afterEach, vi } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createPresentationService } from '../../shell/src/main/presentation-service.js'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { createPresentationTeamSkill } from '../src/skills/powerpoint/presentation-team.js'
import {
  presentationTeamId,
  type PresentationTeamContext,
} from '@wiswork/pptx-engine/presentation-team'
const roots: string[] = []
afterEach(() => roots.splice(0).forEach((root) => rmSync(root, { recursive: true, force: true })))
async function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'team-controller-'))
  roots.push(root)
  const owner = 'a'.repeat(64),
    reviewer = 'b'.repeat(64),
    viewer = 'c'.repeat(64),
    plan = benchmarkPlan()
  let actor = owner,
    doc = 'doc',
    service = createPresentationService({ userDataPath: root })
  const requests: Record<string, unknown>[] = []
  const raw = async (body: Record<string, unknown>, team = true) => {
    requests.push(body)
    const bytes = await service(
      body,
      new AbortController().signal,
      team
        ? ({ version: 1, actorSubject: actor, pcSubject: owner } satisfies PresentationTeamContext)
        : undefined,
    )
    const value = JSON.parse(Buffer.from(bytes).toString())
    return Response.json(value, { status: value.error ? 400 : 200 })
  }
  expect(
    (
      await (
        await raw(
          {
            operation: 'save_plan',
            documentId: doc,
            projectId: plan.projectId,
            plan,
            expectedRevision: 0,
          },
          false,
        )
      ).json()
    ).revision,
  ).toBe(1)
  const proposals = createStructuredProposalController()
  const skill = createPresentationTeamSkill({
    proposals,
    documentId: async () => doc,
    request: (body) => raw(body as Record<string, unknown>, false),
    teamAvailable: () => true,
    teamRequest: (body) => raw(body as Record<string, unknown>),
  })
  const controller = createPresentationTeamController({
    skill,
    documentId: async () => doc,
    proposals,
  })
  const teamId = await presentationTeamId(owner, doc, plan.projectId)
  const confirm = async (id: string) => {
    await proposals.confirm(id)
    await vi.waitFor(() => expect(controller.snapshot().phase).toBe('idle'))
  }
  return {
    controller,
    skill,
    proposals,
    raw,
    requests,
    plan,
    teamId,
    owner,
    reviewer,
    viewer,
    confirm,
    setActor: (v: string) => (actor = v),
    setDoc: (v: string) => (doc = v),
    reopen: () => (service = createPresentationService({ userDataPath: root })),
  }
}
it('creates only a visible proposal then refreshes persisted publication after confirmation and restart', async () => {
  const f = await fixture()
  await f.controller.refresh()
  const cached = f.controller.snapshot()
  expect(cached).toBe(f.controller.snapshot())
  expect(Object.isFrozen(cached)).toBe(true)
  const id = await f.controller.create(f.plan.projectId, 1)
  expect(f.controller.snapshot().phase).toBe('awaiting_confirmation')
  expect(
    (await (await f.raw({ operation: 'team_project_read', teamId: f.teamId })).json()).error,
  ).toBe('not_found')
  await f.confirm(id)
  expect(f.controller.snapshot().team?.revision).toBe(1)
  expect(f.controller.snapshot().role).toBe('owner')
  f.reopen()
  await f.controller.refresh(f.teamId)
  expect(f.controller.snapshot().team?.publishedPlan.plan).toEqual(f.plan)
})
it('cancel and clear discard proposals without shared writes', async () => {
  const f = await fixture()
  await f.controller.refresh()
  await f.controller.create(f.plan.projectId, 1)
  f.proposals.reject()
  await vi.waitFor(() => expect(f.controller.snapshot().phase).toBe('idle'))
  await f.controller.create(f.plan.projectId, 1)
  f.controller.clear()
  expect(f.proposals.pending()).toBeUndefined()
  expect(f.controller.snapshot().team).toBeUndefined()
  expect(f.requests.filter((r) => r.operation === 'team_project_create')).toHaveLength(0)
})
it('uses current revision and actual actor roles; viewers and revoked reviewers cannot propose comments', async () => {
  const f = await fixture()
  await f.controller.refresh()
  await f.confirm(await f.controller.create(f.plan.projectId, 1))
  await f.confirm(await f.controller.setMember(f.reviewer, 'reviewer'))
  await f.confirm(await f.controller.setMember(f.viewer, 'viewer'))
  f.setActor(f.reviewer)
  await f.controller.refresh(f.teamId)
  expect(f.controller.snapshot().role).toBe('reviewer')
  await expect(f.controller.publish(1)).rejects.toThrow('access_denied')
  await f.controller.refresh(f.teamId)
  await f.confirm(await f.controller.addComment('slide', f.plan.slides[0]!.id, 'Untrusted comment'))
  expect(f.controller.snapshot().team?.comments[0]?.authorSubject).toBe(f.reviewer)
  f.setActor(f.viewer)
  await f.controller.refresh(f.teamId)
  expect(f.controller.snapshot().role).toBe('viewer')
  await expect(f.controller.addComment('slide', f.plan.slides[0]!.id, 'Denied')).rejects.toThrow(
    'access_denied',
  )
  f.setActor(f.owner)
  await f.controller.refresh(f.teamId)
  await f.confirm(await f.controller.revokeMember(f.reviewer))
  f.setActor(f.reviewer)
  await f.controller.refresh(f.teamId)
  expect(f.controller.snapshot().team).toBeUndefined()
  expect(f.controller.snapshot().error).toBe('access_denied')
})
it('rejects account and document switches before proposing any mutation', async () => {
  const f = await fixture()
  await f.controller.refresh()
  f.setActor(f.reviewer)
  await expect(f.controller.create(f.plan.projectId, 1)).rejects.toThrow(
    'presentation_team_account_changed',
  )
  expect(f.proposals.pending()).toBeUndefined()
  f.setActor(f.owner)
  await f.controller.refresh()
  f.setDoc('different')
  await expect(f.controller.create(f.plan.projectId, 1)).rejects.toThrow(
    'presentation_document_changed',
  )
  expect(f.requests.filter((r) => r.operation === 'team_project_create')).toHaveLength(0)
})
it('rejects CAS changes instead of silently adopting a newer ledger for a proposal', async () => {
  const f = await fixture()
  await f.controller.refresh()
  await f.confirm(await f.controller.create(f.plan.projectId, 1))
  const expectedIdentity = { version: 1, actorSubject: f.owner, pcSubject: f.owner }
  expect(
    (
      await (
        await f.raw({
          operation: 'team_member_set',
          teamId: f.teamId,
          expectedRevision: 1,
          expectedIdentity,
          memberSubject: f.reviewer,
          role: 'reviewer',
        })
      ).json()
    ).team.revision,
  ).toBe(2)
  await expect(f.controller.setMember(f.viewer, 'viewer')).rejects.toThrow('revision_conflict')
  expect(f.proposals.pending()).toBeUndefined()
})
it('suppresses a late identity read after clear and sanitizes callback failures', async () => {
  let resolve!: (v: { output: string; summary: string }) => void
  const proposals = createStructuredProposalController()
  const skill = {
    id: 'team',
    systemPrompt: '',
    tools: [{ name: 'read_presentation_team_identity', description: '', inputSchema: {} }],
    executeTool: () => new Promise<{ output: string; summary: string }>((r) => (resolve = r)),
  }
  const controller = createPresentationTeamController({
    skill,
    documentId: async () => 'doc',
    proposals,
  })
  const pending = controller.refresh()
  await vi.waitFor(() => expect(resolve).toBeDefined())
  controller.clear()
  resolve({
    summary: 'identity',
    output: JSON.stringify({
      identity: { version: 1, actorSubject: 'a'.repeat(64), pcSubject: 'a'.repeat(64) },
    }),
  })
  await pending
  expect(controller.snapshot().identity).toBeUndefined()
  expect(controller.snapshot().phase).toBe('idle')
})
it('confirmation after a document switch cannot publish and reports a safe failure', async () => {
  const f = await fixture()
  await f.controller.refresh()
  const id = await f.controller.create(f.plan.projectId, 1)
  f.setDoc('different')
  await expect(f.proposals.confirm(id)).rejects.toThrow()
  await vi.waitFor(() => expect(f.controller.snapshot().phase).toBe('idle'))
  expect(f.requests.filter((r) => r.operation === 'team_project_create')).toHaveLength(0)
  expect(f.controller.snapshot().team).toBeUndefined()
})
it('does not replace an awaiting visible confirmation with another action', async () => {
  const f = await fixture()
  await f.controller.refresh()
  const id = await f.controller.create(f.plan.projectId, 1)
  await expect(f.controller.refresh(f.teamId)).rejects.toThrow('presentation_team_busy')
  await expect(f.controller.create(f.plan.projectId, 1)).rejects.toThrow('presentation_team_busy')
  expect(f.proposals.pending()?.id).toBe(id)
  expect(f.requests.filter((r) => r.operation === 'team_project_create')).toHaveLength(0)
})
it('rejects a changed authenticated actor at confirmation instead of publishing under that account', async () => {
  const f = await fixture()
  await f.controller.refresh()
  const id = await f.controller.create(f.plan.projectId, 1)
  f.setActor(f.reviewer)
  await expect(f.proposals.confirm(id)).rejects.toThrow()
  await vi.waitFor(() => expect(f.controller.snapshot().phase).toBe('idle'))
  expect(f.requests.filter((r) => r.operation === 'team_project_create')).toHaveLength(0)
})
it('clear rejects a freshly created proposal while its document recheck is pending', async () => {
  const f = await fixture()
  let release!: (v: string) => void
  const controller = createPresentationTeamController({
    skill: f.skill,
    proposals: f.proposals,
    documentId: () =>
      f.proposals.pending() ? new Promise<string>((r) => (release = r)) : Promise.resolve('doc'),
  })
  await controller.refresh()
  const pending = controller.create(f.plan.projectId, 1)
  const rejected = expect(pending).rejects.toThrow('cancelled')
  await vi.waitFor(() => expect(release).toBeDefined())
  controller.clear()
  release('doc')
  await rejected
  expect(f.proposals.pending()).toBeUndefined()
  expect(f.requests.filter((r) => r.operation === 'team_project_create')).toHaveLength(0)
})
it('never exports a raw callback error in snapshots', async () => {
  const controller = createPresentationTeamController({
    skill: {
      id: 'team',
      tools: [{ name: 'read_presentation_team_identity', description: '', inputSchema: {} }],
      systemPrompt: '',
      executeTool() {
        throw Error('secret-token/private/source')
      },
    },
    documentId: async () => 'doc',
    proposals: createStructuredProposalController(),
  })
  await controller.refresh()
  expect(controller.snapshot().error).toBe('presentation_response_invalid')
  expect(JSON.stringify(controller.snapshot())).not.toContain('secret')
})
it('clears published team and identity if capability disappears during the final document recheck', async () => {
  const identity = { version: 1 as const, actorSubject: 'a'.repeat(64), pcSubject: 'a'.repeat(64) },
    plan = benchmarkPlan()
  const teamId = await presentationTeamId(identity.pcSubject, 'doc', plan.projectId)
  const team = {
    version: 1,
    teamId,
    documentId: 'doc',
    projectId: plan.projectId,
    ownerSubject: identity.actorSubject,
    revision: 1,
    createdAt: '2026-09-29T00:00:00.000Z',
    updatedAt: '2026-09-29T00:00:00.000Z',
    publishedPlan: { revision: 1, plan },
    members: [],
    comments: [],
  }
  let available = true,
    identities = 0,
    release!: (v: string) => void
  const controller = createPresentationTeamController({
    proposals: createStructuredProposalController(),
    skill: {
      id: 'team',
      systemPrompt: '',
      get tools() {
        return available
          ? [{ name: 'read_presentation_team_identity', description: '', inputSchema: {} }]
          : []
      },
      async executeTool(call) {
        if (call.name === 'read_presentation_team_identity') {
          identities++
          return { output: JSON.stringify({ identity }), summary: 'identity' }
        }
        return { output: JSON.stringify({ team }), summary: 'published' }
      },
    },
    documentId: () =>
      identities === 2 ? new Promise<string>((r) => (release = r)) : Promise.resolve('doc'),
  })
  const pending = controller.refresh(teamId)
  await vi.waitFor(() => expect(release).toBeDefined())
  available = false
  release('doc')
  await pending
  expect(controller.snapshot()).toEqual({
    available: false,
    phase: 'idle',
    error: 'presentation_team_unavailable',
  })
})
it('cancels its newly returned proposal if capability disappears before the post-tool guard', async () => {
  const f = await fixture()
  let available = true
  const controller = createPresentationTeamController({
    proposals: f.proposals,
    documentId: async () => 'doc',
    skill: {
      ...f.skill,
      get tools() {
        return available ? f.skill.tools : []
      },
      async executeTool(call, signal) {
        const result = await f.skill.executeTool(call, signal)
        if (call.name === 'create_presentation_team') available = false
        return result
      },
    },
  })
  await controller.refresh()
  await expect(controller.create(f.plan.projectId, 1)).rejects.toThrow(
    'presentation_team_unavailable',
  )
  expect(f.proposals.pending()).toBeUndefined()
  expect(controller.snapshot()).toEqual({
    available: false,
    phase: 'idle',
    error: 'presentation_team_unavailable',
  })
  expect(f.requests.filter((r) => r.operation === 'team_project_create')).toHaveLength(0)
})
