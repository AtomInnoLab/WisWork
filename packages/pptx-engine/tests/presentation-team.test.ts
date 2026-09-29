import { expect, it } from 'vitest'
import { parsePresentationTeamContext, parsePresentationTeamLedger } from '../src/presentation-team'
it('accepts only verified subject hashes and exact trusted context keys', () => {
  const context = { version: 1, actorSubject: 'a'.repeat(64), pcSubject: 'b'.repeat(64) }
  expect(parsePresentationTeamContext(context)).toEqual(context)
  expect(() => parsePresentationTeamContext({ ...context, authorLabel: 'owner' })).toThrow()
  expect(() => parsePresentationTeamContext({ ...context, actorSubject: 'Alice' })).toThrow()
  expect(() => parsePresentationTeamLedger({ version: 1 })).toThrow()
})
import { benchmarkPlan } from './fixtures/presentation-plan'
import { presentationTeamId } from '../src/presentation-team'
const owner = 'a'.repeat(64),
  member = 'b'.repeat(64),
  now = '2026-09-29T00:00:00.000Z'
const ledger = () => ({
  version: 1,
  teamId: 'team_' + 'f'.repeat(64),
  documentId: 'doc',
  projectId: benchmarkPlan().projectId,
  ownerSubject: owner,
  revision: 1,
  createdAt: now,
  updatedAt: now,
  publishedPlan: { revision: 1, plan: benchmarkPlan() },
  members: [{ subject: member, role: 'reviewer' }],
  comments: [],
})
it('clones a complete bounded snapshot without converting comments into QA approval', () => {
  const value = ledger(),
    parsed = parsePresentationTeamLedger(value)
  expect(parsed).toEqual(value)
  parsed.members[0]!.role = 'viewer'
  expect(value.members[0]!.role).toBe('reviewer')
  expect(parsed).not.toHaveProperty('qaPassed')
})
it.each([
  'owner-member',
  'duplicate-member',
  'unknown-role',
  'bad-date',
  'unknown-field',
  'plan-project',
] as const)('rejects invalid full ledger %s', (kind) => {
  const value = ledger()
  if (kind === 'owner-member') value.members[0]!.subject = owner
  if (kind === 'duplicate-member') value.members.push({ ...value.members[0]! })
  if (kind === 'unknown-role') value.members[0]!.role = 'owner'
  if (kind === 'bad-date') value.updatedAt = '2026-02-30T00:00:00.000Z'
  if (kind === 'unknown-field') Object.assign(value, { actorSubject: owner })
  if (kind === 'plan-project') value.projectId = 'other'
  expect(() => parsePresentationTeamLedger(value)).toThrow()
})
it('retains historical comment version while requiring current targets and canonical authenticated authors', () => {
  const value = {
    ...ledger(),
    publishedPlan: { revision: 2, plan: benchmarkPlan() },
    comments: [
      {
        id: 'historical',
        targetKind: 'source',
        targetId: 'removed-source',
        authorSubject: member,
        text: 'Historical only',
        planRevision: 1,
        state: 'open',
        createdAt: now,
        updatedAt: now,
      },
    ],
  }
  expect(parsePresentationTeamLedger(value).comments[0]!.planRevision).toBe(1)
  expect(() =>
    parsePresentationTeamLedger({
      ...value,
      comments: [{ ...value.comments[0], planRevision: 2 }],
    }),
  ).toThrow()
  expect(() =>
    parsePresentationTeamLedger({
      ...value,
      comments: [{ ...value.comments[0], authorSubject: 'author label' }],
    }),
  ).toThrow()
})
it('scopes deterministic team identity to owner/document/project', async () => {
  const id = await presentationTeamId(owner, 'doc', 'project')
  expect(id).toMatch(/^team_[a-f0-9]{64}$/)
  expect(await presentationTeamId(owner, 'other', 'project')).not.toBe(id)
  expect(await presentationTeamId(member, 'doc', 'project')).not.toBe(id)
})
