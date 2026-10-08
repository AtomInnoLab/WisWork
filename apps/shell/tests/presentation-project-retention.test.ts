import { afterEach, expect, it } from 'vitest'
import { existsSync, mkdirSync, mkdtempSync, rmSync, symlinkSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { PresentationLifecycleStore, PresentationStore } from '@wiswork/project-store'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { inspectPresentationProjectInventory } from '../src/main/presentation-project-inventory'
import {
  createPresentationProjectRetentionService,
  presentationRetentionEnabled,
} from '../src/main/presentation-project-retention'
import { registerPresentationProjectWork } from '../src/main/presentation-project-work'

const roots: string[] = []
it('requires both explicit PC switches before automatic retention can run', () => {
  expect(presentationRetentionEnabled({})).toBe(false)
  expect(presentationRetentionEnabled({ WISWORK_PPT_PROJECT_GOVERNANCE_ENABLED: '1' })).toBe(false)
  expect(presentationRetentionEnabled({ WISWORK_PPT_RETENTION_AUTOMATION_ENABLED: '1' })).toBe(
    false,
  )
  expect(
    presentationRetentionEnabled({
      WISWORK_PPT_PROJECT_GOVERNANCE_ENABLED: '1',
      WISWORK_PPT_RETENTION_AUTOMATION_ENABLED: '1',
    }),
  ).toBe(true)
})
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { force: true, recursive: true })
})
function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'ppt-retention-'))
  roots.push(root)
  const scope = { projectId: benchmarkPlan().projectId, documentId: 'doc' }
  const store = new PresentationStore(root)
  store.savePlan(scope.projectId, scope.documentId, 0, benchmarkPlan())
  const life = new PresentationLifecycleStore(root)
  let tail = Promise.resolve()
  const lock = async () => {
    const previous = tail
    let release!: () => void
    tail = new Promise<void>((done) => {
      release = done
    })
    await previous
    return release
  }
  const now = () => new Date(Date.now() + 31 * 86400000)
  const make = () =>
    createPresentationProjectRetentionService({ userDataPath: root, acquireProjectLock: lock, now })
  return { root, scope, store, life, make }
}

it('never cleans an existing project with the default null content policy', async () => {
  const f = fixture()
  f.life.initialize(f.scope)
  expect(await f.make().tick()).toEqual({
    considered: 1,
    started: 0,
    resumed: 0,
    pruned: 0,
    skipped: 1,
  })
  expect(f.store.projectScope(f.scope.projectId, f.scope.documentId)).toEqual(f.scope)
  expect(f.life.read(f.scope)?.state).toBe('active')
})

it('starts an expired saved policy through the durable retention deletion intent', async () => {
  const f = fixture()
  const initial = f.life.initialize(f.scope)
  f.life.setPolicy(f.scope, initial.revision, {
    contentRetentionDays: 30,
    auditRetentionDays: null,
  })
  const result = await f.make().tick()
  expect(result).toEqual({ considered: 1, started: 1, resumed: 0, pruned: 0, skipped: 0 })
  const tombstone = f.life.readControl(f.scope)
  expect(tombstone?.state).toBe('deleted')
  expect(tombstone?.deletion?.reason).toBe('retention')
  expect(f.store.projectScope(f.scope.projectId, f.scope.documentId)).toBeUndefined()
  expect(existsSync(f.root)).toBe(true)
})

it('prunes expired anonymous audit after content deletion without reviving the project', async () => {
  const f = fixture()
  const initial = f.life.initialize(f.scope)
  f.life.setPolicy(f.scope, initial.revision, {
    contentRetentionDays: 30,
    auditRetentionDays: 1,
  })
  expect((await f.make().tick()).started).toBe(1)
  expect((await f.make().tick()).pruned).toBe(1)
  expect(() => f.life.initialize(f.scope)).toThrow('project_deleted')
  expect(() => f.life.exportAudit(f.scope)).toThrow('project_deleted')
  expect((await f.make().tick()).considered).toBe(0)
})

it('waits until the actual content activity exceeds the saved retention interval', async () => {
  const f = fixture()
  const initial = f.life.initialize(f.scope)
  f.life.setPolicy(f.scope, initial.revision, {
    contentRetentionDays: 60,
    auditRetentionDays: null,
  })
  expect(await f.make().tick()).toEqual({
    considered: 1,
    started: 0,
    resumed: 0,
    pruned: 0,
    skipped: 1,
  })
  expect(f.store.projectScope(f.scope.projectId, f.scope.documentId)).toEqual(f.scope)
})

it('skips an expired project while its actual foreground work is registered', async () => {
  const f = fixture()
  const initial = f.life.initialize(f.scope)
  f.life.setPolicy(f.scope, initial.revision, {
    contentRetentionDays: 30,
    auditRetentionDays: null,
  })
  const work = registerPresentationProjectWork({ scope: { root: f.root, ...f.scope } })
  try {
    expect(await f.make().tick()).toEqual({
      considered: 1,
      started: 0,
      resumed: 0,
      pruned: 0,
      skipped: 1,
    })
    expect(f.life.read(f.scope)?.state).toBe('active')
  } finally {
    work.finish()
  }
  expect((await f.make().tick()).started).toBe(1)
})

it('never follows a symlinked control directory during enumeration', async () => {
  const f = fixture()
  const external = mkdtempSync(join(tmpdir(), 'ppt-retention-external-'))
  roots.push(external)
  const name = createHash('sha256').update('external').digest('hex')
  mkdirSync(join(f.root, 'presentation-project-lifecycles'))
  symlinkSync(external, join(f.root, 'presentation-project-lifecycles', name))
  expect(await f.make().tick()).toEqual({
    considered: 0,
    started: 0,
    resumed: 0,
    pruned: 0,
    skipped: 0,
  })
  expect(existsSync(external)).toBe(true)
})

it('resumes the original durable retention intent after an interrupted pass', async () => {
  const f = fixture()
  const initial = f.life.initialize(f.scope)
  const policy = f.life.setPolicy(f.scope, initial.revision, {
    contentRetentionDays: 30,
    auditRetentionDays: null,
  })
  const inventory = await inspectPresentationProjectInventory({ userDataPath: f.root, ...f.scope })
  const project = inventory.resources.find((resource) => resource.kind === 'project')!
  f.life.beginDeletion(f.scope, policy.revision, {
    deletionId: 'original_retention',
    reason: 'retention',
    resources: [
      { resourceId: project.resourceId, kind: 'project', ownership: 'project_exclusive' },
    ],
  })
  expect(await f.make().tick()).toEqual({
    considered: 1,
    started: 0,
    resumed: 1,
    pruned: 0,
    skipped: 0,
  })
  expect(f.life.readControl(f.scope)?.deletion?.deletionId).toBe('original_retention')
  expect(f.life.readControl(f.scope)?.state).toBe('deleted')
})

it('two service instances cannot start two retention intents for the same project', async () => {
  const f = fixture()
  const initial = f.life.initialize(f.scope)
  f.life.setPolicy(f.scope, initial.revision, {
    contentRetentionDays: 30,
    auditRetentionDays: null,
  })
  const results = await Promise.all([f.make().tick(), f.make().tick()])
  expect(results.reduce((sum, result) => sum + result.started, 0)).toBe(1)
  expect(f.life.readControl(f.scope)?.deletion?.reason).toBe('retention')
})
