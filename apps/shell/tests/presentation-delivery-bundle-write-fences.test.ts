import { afterEach, expect, it, vi } from 'vitest'
import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import JSZip from 'jszip'
import { PresentationLifecycleStore } from '@wiswork/project-store'
import {
  capturePresentationProjectReadLease,
  capturePresentationProjectWriteLease,
} from '../src/main/presentation-project-write-lease'
import { stopPresentationProjectWork } from '../src/main/presentation-project-work'
import {
  deliveryBundleFixture,
  cleanupDeliveryBundleFixtures,
} from './helpers/delivery-bundle-fixture'
const fsHook = vi.hoisted(() => ({
  beforeOpen: undefined as ((path: unknown, flags: unknown) => void) | undefined,
}))
vi.mock('node:fs/promises', async (original) => {
  const actual = await original<typeof import('node:fs/promises')>()
  return {
    ...actual,
    open: (...args: Parameters<typeof actual.open>) => {
      fsHook.beforeOpen?.(args[0], args[1])
      return actual.open(...args)
    },
  }
})
afterEach(() => {
  fsHook.beforeOpen = undefined
  vi.restoreAllMocks()
  cleanupDeliveryBundleFixtures()
})
const hash = (value: string | Uint8Array) => createHash('sha256').update(value).digest('hex')
const paths = (f: Awaited<ReturnType<typeof deliveryBundleFixture>>) => {
  const project = join(
    f.root,
    'presentation-delivery-bundles',
    hash(f.base.documentId),
    hash(f.base.projectId),
  )
  return { project, dir: join(project, hash(f.base.bundleId)) }
}
const tree = (root: string) =>
  readdirSync(root, { recursive: true, withFileTypes: true })
    .filter((e) => e.isFile())
    .map((e) => [
      join(e.parentPath, e.name),
      readFileSync(join(e.parentPath, e.name)).toString('base64'),
    ])
    .sort()
async function guardedFixture(onGuard?: () => void) {
  let lifecycle!: PresentationLifecycleStore
  const f = await deliveryBundleFixture(undefined, (root, store) => {
    lifecycle = new PresentationLifecycleStore(root)
    return {
      captureProjectLease: (scope, mode, signal) => {
        const input = {
          store: lifecycle,
          scope,
          signal,
          readExistingProject: (s: typeof scope) => store.projectScope(s.projectId, s.documentId),
        }
        if (mode === 'read') return capturePresentationProjectReadLease(input)
        const lease = capturePresentationProjectWriteLease(input)
        return {
          assertCurrent: () => {
            onGuard?.()
            lease.assertWritable()
          },
        }
      },
    }
  })
  lifecycle.initialize({ documentId: f.base.documentId, projectId: f.base.projectId })
  const freeze = () =>
    lifecycle.beginDeletion({ documentId: f.base.documentId, projectId: f.base.projectId }, 0, {
      deletionId: 'delete',
      reason: 'user',
      resources: [
        { resourceId: 'bundle', kind: 'delivery_bundles', ownership: 'project_exclusive' },
      ],
    })
  return { ...f, lifecycle, freeze }
}
it('pure list observes a missing bundle namespace without creating directories', async () => {
  const f = await deliveryBundleFixture(),
    before = tree(f.root)
  const scope = {
    documentId: f.base.documentId,
    projectId: f.base.projectId,
    requestId: f.base.requestId,
  }
  expect(
    await f.service({ ...scope, operation: 'delivery_bundle_list' }, new AbortController().signal),
  ).toEqual({ bundles: [] })
  expect(existsSync(join(f.root, 'presentation-delivery-bundles'))).toBe(false)
  expect(tree(f.root)).toEqual(before)
})
it('metadata retains interrupted receipt bytes and does not remove another invocation staging files', async () => {
  const f = await deliveryBundleFixture()
  const receipt = await f.begin()
  const { project, dir } = paths(f),
    stage = join(project, '.tmp-12345678-1234-4234-8234-123456789abc')
  mkdirSync(stage)
  writeFileSync(join(stage, 'bundle.zip'), 'other invocation')
  writeFileSync(join(dir, 'bundle.zip'), f.raw)
  const before = tree(f.root)
  expect(await f.call('metadata')).toEqual(receipt)
  expect(tree(f.root)).toEqual(before)
  expect(await f.call('chunk', { offset: 0, base64: f.raw.toString('base64') })).toMatchObject({
    receivedBytes: f.raw.length,
  })
  expect(existsSync(stage)).toBe(true)
})
it('freezes after actual ZIP validation awaits without publishing a ready receipt', async () => {
  const f = await guardedFixture()
  await f.upload()
  const before = tree(f.root)
  let release!: () => void, enter!: () => void
  const gate = new Promise<void>((r) => (release = r)),
    entered = new Promise<void>((r) => (enter = r)),
    load = JSZip.loadAsync.bind(JSZip)
  vi.spyOn(JSZip, 'loadAsync').mockImplementation(async (...args) => {
    enter()
    await gate
    return load(...args)
  })
  const pending = f.call('finish')
  await entered
  f.freeze()
  release()
  await expect(pending).rejects.toThrow('revision_conflict')
  // Lifecycle metadata is expected to change; every bundle byte stays unchanged.
  expect(tree(join(f.root, 'presentation-delivery-bundles'))).toEqual(
    before.filter(([p]) => p!.includes('presentation-delivery-bundles')),
  )
})
it('drains a real bundle validation before declaring scoped work finished', async () => {
  const f = await guardedFixture()
  await f.upload()
  let release!: () => void, enter!: () => void
  const gate = new Promise<void>((r) => (release = r)),
    entered = new Promise<void>((r) => (enter = r)),
    load = JSZip.loadAsync.bind(JSZip)
  vi.spyOn(JSZip, 'loadAsync').mockImplementation(async (...args) => {
    enter()
    await gate
    return load(...args)
  })
  const pending = f.call('finish'),
    settled = pending.catch((e) => e)
  await entered
  f.freeze()
  let drained = false
  const drain = stopPresentationProjectWork({
    root: f.root,
    documentId: f.base.documentId,
    projectId: f.base.projectId,
  }).then(() => {
    drained = true
  })
  await new Promise<void>((done) => setImmediate(done))
  const premature = drained
  release()
  await drain
  expect(premature).toBe(false)
  expect(drained).toBe(true)
  expect(await settled).toBeInstanceOf(Error)
})

it('blocks chunk canonical publication when freeze follows the actual temporary fd open', async () => {
  let armed = false,
    frozen = false
  const f: Awaited<ReturnType<typeof guardedFixture>> = await guardedFixture(() => {
    if (
      armed &&
      !frozen &&
      readdirSync(f.root, { recursive: true }).some((name) => String(name).endsWith('.tmp'))
    ) {
      frozen = true
      f.freeze()
    }
  })
  await f.begin()
  const before = tree(join(f.root, 'presentation-delivery-bundles'))
  armed = true
  await expect(f.call('chunk', { offset: 0, base64: f.raw.toString('base64') })).rejects.toThrow(
    'revision_conflict',
  )
  expect(frozen).toBe(true)
  expect(tree(join(f.root, 'presentation-delivery-bundles'))).toEqual(before)
})
it('fixes scope and manifest aliases before the first private await', async () => {
  const f = await deliveryBundleFixture(),
    original = structuredClone(f.manifest),
    request = {
      ...f.base,
      operation: 'delivery_bundle_begin',
      sha256: f.base.bundleId,
      sizeBytes: f.raw.length,
      manifest: f.manifest,
    }
  const pending = f.service(request, new AbortController().signal)
  request.documentId = 'foreign'
  request.bundleId = 'f'.repeat(64)
  request.manifest.projectId = 'foreign'
  expect(await pending).toMatchObject({
    documentId: f.base.documentId,
    bundleId: f.base.bundleId,
    manifest: original,
  })
})

it('never removes another invocation file when exclusive temporary open fails', async () => {
  const f = await deliveryBundleFixture()
  await f.begin()
  let collision: string | undefined
  fsHook.beforeOpen = (path, flags) => {
    if (
      !collision &&
      flags === 'wx' &&
      String(path).includes('bundle.zip.') &&
      String(path).endsWith('.tmp')
    ) {
      collision = String(path)
      writeFileSync(collision, 'other invocation')
    }
  }
  await expect(f.call('chunk', { offset: 0, base64: f.raw.toString('base64') })).rejects.toThrow(
    'invalid_state',
  )
  expect(collision).toBeDefined()
  expect(readFileSync(collision!, 'utf8')).toBe('other invocation')
  expect(readFileSync(join(paths(f).dir, 'bundle.zip'))).toHaveLength(0)
})
