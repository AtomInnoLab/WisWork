import { mkdtempSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it } from 'vitest'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'

it('pins local review comments to a plan revision and requires optimistic updates', async () => {
  const root = mkdtempSync(join(tmpdir(), 'wiswork-comments-'))
  try {
    const plan = benchmarkPlan()
    const first = createPresentationService({ userDataPath: root })
    const call = async (
      service: ReturnType<typeof createPresentationService>,
      input: Record<string, unknown>,
    ) =>
      JSON.parse(
        Buffer.from(
          await service(
            { documentId: 'doc-1', projectId: plan.projectId, ...input },
            new AbortController().signal,
          ),
        ).toString('utf8'),
      )
    expect(
      (await call(first, { operation: 'save_plan', expectedRevision: 0, plan })).revision,
    ).toBe(1)
    expect(await call(first, { operation: 'comment_list' })).toMatchObject({
      revision: 0,
      comments: [],
    })
    const comment = {
      id: 'comment-1',
      targetKind: 'slide',
      targetId: plan.slides[0]!.id,
      authorLabel: '审阅人甲',
      text: '请核对结论',
    }
    expect(
      await call(first, {
        operation: 'comment_add',
        expectedRevision: 0,
        planRevision: 1,
        comment,
      }),
    ).toMatchObject({ revision: 1, comments: [{ ...comment, planRevision: 1, state: 'open' }] })
    expect(await call(first, { operation: 'status' })).toMatchObject({
      reviewComments: {
        revision: 1,
        openCount: 1,
        resolvedCount: 0,
        recent: [{ id: 'comment-1', planRevision: 1 }],
      },
    })
    const reopened = createPresentationService({ userDataPath: root })
    expect(await call(reopened, { operation: 'comment_list' })).toMatchObject({
      revision: 1,
      comments: [{ id: 'comment-1' }],
    })
    expect(
      await call(reopened, {
        operation: 'comment_add',
        expectedRevision: 0,
        planRevision: 1,
        comment: { ...comment, id: 'comment-2' },
      }),
    ).toEqual({ error: 'revision_conflict' })
    expect(
      await call(reopened, {
        operation: 'comment_add',
        expectedRevision: 1,
        planRevision: 1,
        comment: { ...comment, id: 'comment-2', targetId: 'missing' },
      }),
    ).toEqual({ error: 'invalid_request' })
    expect(
      await call(reopened, {
        operation: 'comment_resolve',
        expectedRevision: 1,
        commentId: 'comment-1',
      }),
    ).toMatchObject({ revision: 2, comments: [{ id: 'comment-1', state: 'resolved' }] })
    const changed = structuredClone(plan)
    changed.slides[0]!.title = '修订标题'
    expect(
      (await call(reopened, { operation: 'save_plan', expectedRevision: 1, plan: changed }))
        .revision,
    ).toBe(2)
    expect(
      await call(reopened, {
        operation: 'comment_add',
        expectedRevision: 2,
        planRevision: 1,
        comment: { ...comment, id: 'comment-3' },
      }),
    ).toEqual({ error: 'invalid_request' })
    expect(await call(reopened, { operation: 'comment_list' })).toMatchObject({
      revision: 2,
      comments: [{ planRevision: 1, state: 'resolved' }],
    })
    expect(
      await call(reopened, {
        operation: 'comment_add',
        expectedRevision: 2,
        planRevision: 2,
        comment: {
          ...comment,
          id: 'comment-3',
          targetKind: 'source',
          targetId: changed.sources[0]!.id,
        },
      }),
    ).toMatchObject({
      revision: 3,
      comments: [expect.any(Object), { targetKind: 'source', planRevision: 2, state: 'open' }],
    })
    const directory = join(root, 'presentation-comments')
    writeFileSync(join(directory, readdirSync(directory)[0]!), '{broken')
    expect(await call(reopened, { operation: 'status' })).toMatchObject({
      commentsUnavailable: true,
      status: 'planned',
    })
    expect(await call(reopened, { operation: 'comment_list' })).toEqual({ error: 'invalid_state' })
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})
