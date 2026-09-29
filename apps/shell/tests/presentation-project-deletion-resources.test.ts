import { afterEach, expect, it, vi } from 'vitest'
import { createHash } from 'node:crypto'
import {
  mkdtempSync,
  mkdirSync,
  writeFileSync,
  readFileSync,
  existsSync,
  rmSync,
  symlinkSync,
  utimesSync,
  renameSync,
  readdirSync,
  unlinkSync,
  constants,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join, dirname } from 'node:path'
import { PresentationStore, PresentationLifecycleStore } from '@wiswork/project-store'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { inspectPresentationProjectInventory } from '../src/main/presentation-project-inventory'
import { removePresentationProjectDeletionResource } from '../src/main/presentation-project-deletion-resources'
const proofOpenHook = vi.hoisted(() => ({
  run: undefined as undefined | ((path: string, flags: unknown) => void),
}))
vi.mock('node:fs', async (load) => {
  const actual = await load<typeof import('node:fs')>()
  return {
    ...actual,
    openSync: (...args: Parameters<typeof actual.openSync>) => {
      proofOpenHook.run?.(String(args[0]), args[1])
      return actual.openSync(...args)
    },
  }
})
const roots: string[] = []
afterEach(() => {
  proofOpenHook.run = undefined
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const hash = (v: string) => createHash('sha256').update(v).digest('hex')
async function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'project-delete-resource-'))
  roots.push(root)
  const store = new PresentationStore(root),
    plan = benchmarkPlan(),
    scope = { projectId: plan.projectId, documentId: 'doc' }
  store.savePlan(scope.projectId, scope.documentId, 0, plan)
  store.savePlan('other', scope.documentId, 0, { ...plan, projectId: 'other' })
  const source = join(root, 'projects', 'presentations', hash(scope.projectId)),
    other = join(root, 'projects', 'presentations', hash('other'))
  writeFileSync(join(source, 'payload.bin'), 'private')
  const inventory = await inspectPresentationProjectInventory({ userDataPath: root, ...scope })
  const resource = inventory.resources.find((r) => r.kind === 'project')!
  const lifecycle = new PresentationLifecycleStore(root)
  lifecycle.initialize(scope)
  const deletionId = 'delete'
  lifecycle.beginDeletion(scope, 0, {
    deletionId,
    reason: 'user',
    resources: [
      { resourceId: resource.resourceId, kind: 'project', ownership: 'project_exclusive' },
    ],
  })
  const options = {
    userDataPath: root,
    scope,
    deletionId,
    resourceId: resource.resourceId,
    expectedRevision: 1,
  }
  const holding = join(
    root,
    'presentation-project-deletion-work',
    hash(scope.projectId),
    hash(deletionId),
    hash(resource.resourceId),
    'data',
  )
  return { root, scope, lifecycle, source, other, resource, options, holding }
}
it('removes the proved project resource while preserving another project and minimal lifecycle', async () => {
  const f = await fixture(),
    before = readFileSync(join(f.other, 'plan.json'))
  expect(await removePresentationProjectDeletionResource(f.options)).toMatchObject({
    status: 'removed',
    resourceId: f.resource.resourceId,
  })
  expect(existsSync(f.source)).toBe(false)
  expect(readFileSync(join(f.other, 'plan.json'))).toEqual(before)
  expect(f.lifecycle.read(f.scope)?.state).toBe('deleting')
  expect(await removePresentationProjectDeletionResource(f.options)).toMatchObject({
    status: 'removed',
  })
})
it('preserves quarantined content after an interruption and resumes the same resource', async () => {
  const f = await fixture()
  const first = await removePresentationProjectDeletionResource({
    ...f.options,
    assertDeleting: () => {
      if (existsSync(f.holding)) throw Error('resource_busy')
    },
  })
  expect(first).toMatchObject({ status: 'failed', code: 'resource_busy' })
  expect(existsSync(f.source)).toBe(false)
  expect(readFileSync(join(f.holding, 'payload.bin'), 'utf8')).toBe('private')
  expect(await removePresentationProjectDeletionResource(f.options)).toMatchObject({
    status: 'removed',
  })
  expect(existsSync(f.holding)).toBe(false)
})
it('rejects a symlink without touching its outside target or original project', async () => {
  const f = await fixture(),
    outside = join(f.root, 'outside.bin')
  writeFileSync(outside, 'keep')
  symlinkSync(outside, join(f.source, 'foreign.bin'))
  expect(await removePresentationProjectDeletionResource(f.options)).toMatchObject({
    status: 'retained',
    code: 'ownership_unproven',
  })
  expect(readFileSync(outside, 'utf8')).toBe('keep')
  expect(readFileSync(join(f.source, 'payload.bin'), 'utf8')).toBe('private')
})
it('refuses a reappeared source while its exact quarantine still exists', async () => {
  const f = await fixture()
  await removePresentationProjectDeletionResource({
    ...f.options,
    assertDeleting: () => {
      if (existsSync(f.holding)) throw Error('resource_busy')
    },
  })
  mkdirSync(f.source)
  writeFileSync(join(f.source, 'project.json'), JSON.stringify({ version: 1, ...f.scope }))
  writeFileSync(join(f.source, 'new.bin'), 'new')
  expect(await removePresentationProjectDeletionResource(f.options)).toMatchObject({
    status: 'failed',
    code: 'resource_busy',
  })
  expect(readFileSync(join(f.source, 'new.bin'), 'utf8')).toBe('new')
  expect(readFileSync(join(f.holding, 'payload.bin'), 'utf8')).toBe('private')
})

it('preserves changed same-inode bytes whose original mtime was restored after quarantine', async () => {
  const f = await fixture(),
    fixed = new Date('2026-09-29T00:00:00.000Z')
  utimesSync(join(f.source, 'payload.bin'), fixed, fixed)
  let changed = false
  expect(
    await removePresentationProjectDeletionResource({
      ...f.options,
      assertDeleting: () => {
        if (!changed && existsSync(f.holding)) {
          changed = true
          writeFileSync(join(f.holding, 'payload.bin'), 'foreign')
          utimesSync(join(f.holding, 'payload.bin'), fixed, fixed)
        }
      },
    }),
  ).toMatchObject({ status: 'retained', code: 'ownership_unproven' })
  expect(changed).toBe(true)
  expect(readFileSync(join(f.holding, 'payload.bin'), 'utf8')).toBe('foreign')
})
it('does not write proof bytes into a foreign directory substituted persistently at leaf open', async () => {
  const f = await fixture(),
    foreign = join(f.root, 'foreign')
  mkdirSync(foreign)
  writeFileSync(join(foreign, 'keep.txt'), 'keep')
  let replaced = false
  proofOpenHook.run = (path, flags) => {
    if (
      !replaced &&
      (flags === 'wx' ||
        flags ===
          (constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | constants.O_NOFOLLOW)) &&
      path.includes('presentation-project-deletion-work') &&
      path.endsWith('.tmp')
    ) {
      replaced = true
      const parent = dirname(path)
      renameSync(parent, parent + '-original')
      symlinkSync(foreign, parent, 'dir')
    }
  }
  expect(await removePresentationProjectDeletionResource(f.options)).toMatchObject({
    status: 'retained',
    code: 'ownership_unproven',
  })
  expect(replaced).toBe(true)
  expect(readFileSync(join(f.source, 'payload.bin'), 'utf8')).toBe('private')
  expect(readdirSync(foreign)).toEqual(['keep.txt'])
  expect(readFileSync(join(foreign, 'keep.txt'), 'utf8')).toBe('keep')
})

it('reopens its own interrupted proof for a legal pre-epoch source file timestamp', async () => {
  const f = await fixture(),
    old = new Date('1960-01-01T00:00:00.000Z')
  utimesSync(join(f.source, 'payload.bin'), old, old)
  expect(
    await removePresentationProjectDeletionResource({
      ...f.options,
      assertDeleting: () => {
        if (existsSync(f.holding)) throw Error('resource_busy')
      },
    }),
  ).toMatchObject({ status: 'failed', code: 'resource_busy' })
  expect(readFileSync(join(f.holding, 'payload.bin'), 'utf8')).toBe('private')
  expect(await removePresentationProjectDeletionResource(f.options)).toMatchObject({
    status: 'removed',
  })
  expect(existsSync(f.holding)).toBe(false)
})

it('does not publish a foreign proof leaf substituted after writing the owned FD', async () => {
  const f = await fixture(),
    work = dirname(f.holding)
  let replaced = false
  const result = await removePresentationProjectDeletionResource({
    ...f.options,
    assertDeleting: () => {
      if (existsSync(f.holding)) throw Error('resource_busy')
      if (!replaced && existsSync(work)) {
        const name = readdirSync(work).find((n) => n.endsWith('.tmp'))
        if (name) {
          const path = join(work, name)
          if (readFileSync(path).length > 0) {
            replaced = true
            unlinkSync(path)
            writeFileSync(path, 'foreign proof')
          }
        }
      }
    },
  })
  expect(replaced).toBe(true)
  expect(result).toMatchObject({ status: 'retained', code: 'ownership_unproven' })
  expect(readFileSync(join(f.source, 'payload.bin'), 'utf8')).toBe('private')
  expect(existsSync(f.holding)).toBe(false)
  expect(
    readFileSync(
      join(
        work,
        readdirSync(work).find((n) => n.endsWith('.tmp'))!,
      ),
      'utf8',
    ),
  ).toBe('foreign proof')
})
it('resumes proof disposal after raw content cleanup was interrupted', async () => {
  const f = await fixture()
  const proof = join(dirname(f.holding), 'proof.json')
  const first = await removePresentationProjectDeletionResource({
    ...f.options,
    assertDeleting: () => {
      if (!existsSync(f.source) && !existsSync(f.holding) && existsSync(proof))
        throw Error('resource_busy')
    },
  })
  expect(first).toMatchObject({ status: 'failed', code: 'resource_busy' })
  expect(existsSync(proof)).toBe(true)
  expect(await removePresentationProjectDeletionResource(f.options)).toMatchObject({
    status: 'removed',
  })
  expect(existsSync(proof)).toBe(false)
})
