import { expect, it, vi } from 'vitest'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import { createPresentationCommentsSkill } from '../src/skills/powerpoint/presentation-comments.js'

it('requires confirmation for local review comments and keeps them separate from QA', async () => {
  const proposals = createStructuredProposalController()
  let ledger = {
    version: 1,
    documentId: 'doc-1',
    projectId: 'project-1',
    revision: 0,
    comments: [] as Record<string, unknown>[],
  }
  const request = vi.fn(async (body: unknown) => {
    const input = body as Record<string, unknown>
    if (input.operation === 'comment_add')
      ledger = {
        ...ledger,
        revision: 1,
        comments: [
          {
            ...(input.comment as object),
            planRevision: 1,
            state: 'open',
            createdAt: '2026-09-28T00:00:00.000Z',
            updatedAt: '2026-09-28T00:00:00.000Z',
          },
        ],
      }
    if (input.operation === 'comment_resolve')
      ledger = { ...ledger, revision: 2, comments: [{ ...ledger.comments[0], state: 'resolved' }] }
    return new Response(JSON.stringify(ledger))
  })
  const skill = createPresentationCommentsSkill({
    available: () => true,
    documentId: async () => 'doc-1',
    request,
    proposals,
  })
  const add = await skill.executeTool({
    id: 'add',
    name: 'add_presentation_review_comment',
    input: {
      project_id: 'project-1',
      expected_revision: 0,
      plan_revision: 1,
      comment: {
        id: 'comment-1',
        targetKind: 'slide',
        targetId: 'slide-1',
        authorLabel: '审阅人甲',
        text: '请核对结论',
      },
    },
  })
  expect(JSON.parse(add.output).status).toBe('awaiting_confirmation')
  expect(request).toHaveBeenCalledTimes(1) // Only a list read before confirmation.
  expect(ledger.comments).toEqual([])
  await proposals.confirm(JSON.parse(add.output).proposalId)
  expect(ledger.comments).toMatchObject([{ id: 'comment-1', state: 'open' }])
  const resolve = await skill.executeTool({
    id: 'resolve',
    name: 'resolve_presentation_review_comment',
    input: { project_id: 'project-1', expected_revision: 1, comment_id: 'comment-1' },
  })
  expect(ledger.comments[0]?.state).toBe('open')
  await proposals.confirm(JSON.parse(resolve.output).proposalId)
  expect(ledger.comments[0]?.state).toBe('resolved')
})
