import { afterEach, expect, it, vi } from 'vitest'
import { mkdtempSync, rmSync, readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { createHash } from 'node:crypto'
import { PresentationLifecycleStore, PresentationStore } from '@wiswork/project-store'
import { PresentationPreferenceLibrary } from '../src/main/presentation-preferences'
import { PresentationManualObservationLibrary } from '../src/main/presentation-manual-observations'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'
const pause = vi.hoisted(() => ({ afterTemp: undefined as ((path: string) => void) | undefined }))
vi.mock('node:fs', async (original) => {
  const actual = await original<typeof import('node:fs')>()
  return {
    ...actual,
    writeFileSync: (...args: Parameters<typeof actual.writeFileSync>) => {
      const result = actual.writeFileSync(...args)
      if (typeof args[0] === 'string' && args[0].endsWith('.tmp')) pause.afterTemp?.(args[0])
      return result
    },
  }
})
const roots: string[] = []
afterEach(() => {
  pause.afterTemp = undefined
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const shape = {
  id: 'shape',
  name: 'Title',
  type: 'TextBox',
  left: 1,
  top: 2,
  width: 300,
  height: 40,
  text: 'before',
}
const approvalId = '12345678-1234-4123-8123-123456789abc'
const digest = (v: string) => createHash('sha256').update(v).digest('hex')
function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'sync-service-fences-'))
  roots.push(root)
  const service = createPresentationService({ userDataPath: root }),
    lifecycle = new PresentationLifecycleStore(root),
    store = new PresentationStore(root)
  const call = async (body: Record<string, unknown>) =>
    JSON.parse(Buffer.from(await service(body, new AbortController().signal)).toString())
  const freeze = (projectId = 'p', documentId = 'doc') => {
    const scope = { projectId, documentId },
      record = lifecycle.read(scope) ?? lifecycle.initialize(scope)
    lifecycle.beginDeletion(scope, record.revision, {
      deletionId: 'delete',
      reason: 'user',
      resources: [{ resourceId: 'own', kind: 'project', ownership: 'project_exclusive' }],
    })
  }
  return { root, call, lifecycle, store, freeze }
}
it.each(['preference', 'manual', 'comment'] as const)(
  'passes a fixed %s guard through real temporary write to final publication',
  async (kind) => {
    const f = fixture(),
      plan = benchmarkPlan(),
      projectId = kind === 'comment' ? plan.projectId : 'p'
    if (kind === 'comment') f.store.savePlan(projectId, 'doc', 0, plan)
    const namespace =
      kind === 'preference'
        ? 'presentation-preferences'
        : kind === 'manual'
          ? 'presentation-manual-observations'
          : 'presentation-comments'
    pause.afterTemp = (path) => {
      if (path.includes(`/${namespace}/`)) {
        pause.afterTemp = undefined
        f.freeze(projectId)
      }
    }
    const body =
      kind === 'preference'
        ? {
            operation: 'preference_save',
            documentId: 'doc',
            preference: { projectId, changeId: 'edit', text: 'Keep titles short' },
          }
        : kind === 'manual'
          ? {
              operation: 'manual_observation_begin',
              documentId: 'doc',
              projectId,
              observationId: 'o',
              slideId: 'slide',
              shape,
            }
          : {
              operation: 'comment_add',
              documentId: 'doc',
              projectId,
              expectedRevision: 0,
              planRevision: 1,
              comment: {
                id: 'c',
                targetKind: 'slide',
                targetId: plan.slides[0]!.id,
                authorLabel: 'User',
                text: 'Review',
              },
            }
    expect(await f.call(body)).toEqual({ error: 'revision_conflict' })
    expect(readdirSync(join(f.root, namespace))).toEqual([])
  },
)
it.each(['source', 'target'] as const)(
  'freezes distinct import %s scope before canonical publish',
  async (which) => {
    const f = fixture(),
      source = { documentId: 'source-doc', projectId: 'source', changeId: 'edit' },
      original = { projectId: source.projectId, changeId: 'edit', text: 'Use short titles' }
    new PresentationPreferenceLibrary(f.root).save(source.documentId, original)
    pause.afterTemp = (path) => {
      if (path.includes('/presentation-preferences/')) {
        pause.afterTemp = undefined
        f.freeze(
          which === 'source' ? source.projectId : 'target',
          which === 'source' ? source.documentId : 'target-doc',
        )
      }
    }
    expect(
      await f.call({
        operation: 'preference_import',
        documentId: 'target-doc',
        projectId: 'target',
        source,
        expectedTextDigest: digest(original.text),
        approvalId,
      }),
    ).toEqual({ error: 'revision_conflict' })
    expect(
      new PresentationPreferenceLibrary(f.root).get(source.documentId, source.projectId, 'edit'),
    ).toEqual(original)
    expect(new PresentationPreferenceLibrary(f.root).list('target-doc', 'target')).toEqual([])
    if (which === 'target')
      expect(
        f.lifecycle.read({ projectId: source.projectId, documentId: source.documentId }),
      ).toBeUndefined()
  },
)
it('creates preserving control for genuine pre-plan records and rejects tombstone writes', async () => {
  const f = fixture()
  const begun = await f.call({
    operation: 'manual_observation_begin',
    documentId: 'doc',
    projectId: 'p',
    observationId: 'o',
    slideId: 'slide',
    shape,
  })
  expect(begun.observation.before.shape).toEqual(shape)
  expect(f.lifecycle.read({ projectId: 'p', documentId: 'doc' })?.policy).toEqual({
    contentRetentionDays: null,
    auditRetentionDays: null,
  })
  expect(f.store.projectScope('p', 'doc')).toBeUndefined()
  f.freeze()
  expect(
    await f.call({
      operation: 'manual_observation_complete',
      documentId: 'doc',
      projectId: 'p',
      observationId: 'o',
      expectedBeforeDigest: begun.observation.before.digest,
      shape: { ...shape, text: 'after' },
    }),
  ).toEqual({ error: 'project_deleting' })
  expect(
    await f.call({
      operation: 'preference_save',
      documentId: 'doc',
      preference: { projectId: 'p', changeId: 'edit', text: 'new' },
    }),
  ).toEqual({ error: 'project_deleting' })
})
it('keeps actual legacy and empty reads byte-identical without control or body creation', async () => {
  const f = fixture(),
    library = new PresentationPreferenceLibrary(f.root),
    manual = new PresentationManualObservationLibrary(f.root)
  library.save('doc', { projectId: 'p', changeId: 'edit', text: 'short titles' })
  manual.begin('doc', 'p', 'o', 'slide', shape)
  const files = () =>
    readdirSync(f.root, { recursive: true, withFileTypes: true })
      .filter((v) => v.isFile())
      .map((v) => [
        join(v.parentPath, v.name),
        readFileSync(join(v.parentPath, v.name)).toString('base64'),
      ])
  const before = files()
  expect(
    await f.call({ operation: 'preference_list', documentId: 'doc', projectId: 'p' }),
  ).toMatchObject({ preferences: [{ changeId: 'edit' }] })
  expect(
    await f.call({
      operation: 'manual_observation_get',
      documentId: 'doc',
      projectId: 'p',
      observationId: 'o',
    }),
  ).toMatchObject({ observation: { projectId: 'p' } })
  expect(
    await f.call({ operation: 'preference_list', documentId: 'doc', projectId: 'unknown' }),
  ).toEqual({ preferences: [] })
  expect(
    await f.call({ operation: 'manual_observation_list', documentId: 'doc', projectId: 'unknown' }),
  ).toEqual({ observations: [] })
  expect(
    await f.call({
      operation: 'preference_delete',
      documentId: 'doc',
      projectId: 'unknown',
      changeId: 'edit',
    }),
  ).toEqual({ deleted: false })
  expect(files()).toEqual(before)
  expect(f.lifecycle.read({ projectId: 'p', documentId: 'doc' })).toBeUndefined()
})
it('does not create lifecycle control for malformed creation payloads', async () => {
  const f = fixture()
  expect(
    await f.call({
      operation: 'preference_save',
      documentId: 'doc',
      preference: { projectId: 'p', changeId: 'edit', text: '' },
    }),
  ).toEqual({ error: 'invalid_request' })
  expect(
    await f.call({
      operation: 'manual_observation_begin',
      documentId: 'doc',
      projectId: 'p',
      observationId: '../bad',
      slideId: 'slide',
      shape,
    }),
  ).toEqual({ error: 'invalid_request' })
  expect(readdirSync(f.root)).toEqual([])
})

it.each(['add_shape', 'add_revision', 'resolve_id', 'resolve_revision'] as const)(
  'rejects invalid legacy comment %s before creating lifecycle control',
  async (kind) => {
    const f = fixture(),
      plan = benchmarkPlan()
    f.store.savePlan(plan.projectId, 'doc', 0, plan)
    const body = kind.startsWith('add')
      ? {
          operation: 'comment_add',
          documentId: 'doc',
          projectId: plan.projectId,
          expectedRevision: kind === 'add_revision' ? -1 : 0,
          planRevision: 1,
          comment:
            kind === 'add_shape'
              ? { text: '' }
              : {
                  id: 'c',
                  targetKind: 'slide',
                  targetId: plan.slides[0]!.id,
                  authorLabel: 'User',
                  text: 'Review',
                },
        }
      : {
          operation: 'comment_resolve',
          documentId: 'doc',
          projectId: plan.projectId,
          expectedRevision: kind === 'resolve_revision' ? -1 : 0,
          commentId: kind === 'resolve_id' ? '' : 'c',
        }
    expect(await f.call(body)).toEqual({ error: 'invalid_request' })
    expect(f.lifecycle.read({ documentId: 'doc', projectId: plan.projectId })).toBeUndefined()
  },
)
