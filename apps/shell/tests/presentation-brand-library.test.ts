import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it } from 'vitest'
import { createPresentationService } from '../src/main/presentation-service'

it('reuses exact brand kit revisions across PC service restarts and projects', async () => {
  const root = mkdtempSync(join(tmpdir(), 'wiswork-brand-library-'))
  try {
    const call = async (
      service: ReturnType<typeof createPresentationService>,
      input: Record<string, unknown>,
    ) =>
      JSON.parse(
        Buffer.from(
          await service({ documentId: 'deck-one', ...input }, new AbortController().signal),
        ).toString('utf8'),
      )
    const first = createPresentationService({ userDataPath: root })
    const kit = {
      id: 'research',
      revision: 1,
      name: '研究品牌',
      allowedColors: ['FFFFFF', '172033', '2255AA'],
    }
    expect(
      await call(first, { operation: 'brand_kit_save', expectedRevision: 0, brandKit: kit }),
    ).toEqual({ brandKit: kit })
    expect(await call(first, { operation: 'brand_kit_list' })).toEqual({ brandKits: [kit] })
    const changed = {
      ...kit,
      revision: 2,
      allowedColors: [...kit.allowedColors, '008844'],
      layoutComponents: [
        {
          id: 'title-body',
          name: '标题与正文',
          layout: 'content',
          slots: [
            { id: 'title', kind: 'text', x: 1, y: 1, w: 10, h: 1 },
            { id: 'body', kind: 'text', x: 1, y: 2.5, w: 10, h: 3 },
          ],
        },
      ],
    }
    expect(
      await call(first, { operation: 'brand_kit_save', expectedRevision: 1, brandKit: changed }),
    ).toEqual({ brandKit: changed })
    expect(
      await call(first, { operation: 'brand_kit_save', expectedRevision: 1, brandKit: changed }),
    ).toEqual({ error: 'revision_conflict' })
    const reopened = createPresentationService({ userDataPath: root })
    expect(
      await call(reopened, { operation: 'brand_kit_get', brandKitId: kit.id, revision: 1 }),
    ).toEqual({ brandKit: kit })
    expect(
      await call(reopened, {
        operation: 'brand_kit_get',
        brandKitId: kit.id,
        revision: 2,
        documentId: 'deck-two',
      }),
    ).toEqual({ brandKit: changed })
    expect(await call(reopened, { operation: 'brand_kit_list' })).toEqual({ brandKits: [changed] })
    expect(
      await call(reopened, {
        operation: 'brand_kit_save',
        expectedRevision: 2,
        brandKit: { ...changed, revision: 3, allowedColors: ['FFFFFF', 'ffffff'] },
      }),
    ).toEqual({ error: 'invalid_brand_kit' })
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})
