import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it } from 'vitest'
import { createPresentationService } from '../src/main/presentation-service'

it('stores confirmed preferences separately from brand kits and survives restart', async () => {
  const root = mkdtempSync(join(tmpdir(), 'wiswork-preferences-'))
  try {
    const call = async (
      service: ReturnType<typeof createPresentationService>,
      input: Record<string, unknown>,
    ) =>
      JSON.parse(
        Buffer.from(
          await service({ documentId: 'doc-1', ...input }, new AbortController().signal),
        ).toString('utf8'),
      )
    const first = createPresentationService({ userDataPath: root })
    const preference = { projectId: 'project-1', changeId: 'edit-1', text: '标题尽量简短' }
    expect(await call(first, { operation: 'preference_save', preference })).toEqual({ preference })
    const reopened = createPresentationService({ userDataPath: root })
    expect(await call(reopened, { operation: 'preference_list', projectId: 'project-1' })).toEqual({
      preferences: [preference],
    })
    expect(await call(reopened, { operation: 'preference_list', projectId: 'project-2' })).toEqual({
      preferences: [],
    })
    expect(await call(reopened, { operation: 'brand_kit_list' })).toEqual({ brandKits: [] })
    expect(await call(reopened, { operation: 'preference_save', preference })).toEqual({
      preference,
    })
    expect(
      await call(reopened, {
        operation: 'preference_save',
        preference: { ...preference, text: '不同偏好' },
      }),
    ).toEqual({ error: 'revision_conflict' })
    expect(
      await call(reopened, {
        operation: 'preference_delete',
        projectId: 'project-1',
        changeId: 'edit-1',
      }),
    ).toEqual({ deleted: true })
    expect(await call(reopened, { operation: 'preference_list', projectId: 'project-1' })).toEqual({
      preferences: [],
    })
    expect(
      await call(reopened, {
        operation: 'preference_delete',
        projectId: 'project-1',
        changeId: 'edit-1',
      }),
    ).toEqual({ deleted: false })
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})
