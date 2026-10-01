import { expect, it } from 'vitest'
import { parsePresentationDeliveryBundleManifest } from '@wiswork/project-store/presentation-delivery-bundle'
import { createPresentationDeliveryBundleService } from '../src/main/presentation-delivery-bundles'
it('rejects unknown manifest fields and operations before writing', async () => {
  expect(() => parsePresentationDeliveryBundleManifest({ version: 1 })).toThrow('invalid_state')
  const service = createPresentationDeliveryBundleService({
    userDataPath: '/tmp/unused-delivery-bundle',
  })
  await expect(
    service({ operation: 'delivery_bundle_unknown' }, new AbortController().signal),
  ).rejects.toThrow('invalid_request')
})
import { writeFileSync, symlinkSync, mkdirSync, existsSync } from 'node:fs'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { PNG } from 'pngjs'
import { afterEach } from 'vitest'
import { parsePresentationDeliveryBundleReceipt } from '@wiswork/project-store/presentation-delivery-bundle'
import {
  deliveryBundleFixture as fixture,
  cleanupDeliveryBundleFixtures,
} from './helpers/delivery-bundle-fixture'
afterEach(cleanupDeliveryBundleFixtures)
const hash = (v: string | Uint8Array) => createHash('sha256').update(v).digest('hex')
it('uploads with overlapping retry, restores ready metadata/read/list and explicit cleanup', async () => {
  const f = await fixture()
  const begun = await f.begin()
  expect(parsePresentationDeliveryBundleReceipt(begun).state).toBe('uploading')
  expect(await f.begin()).toEqual(begun)
  await f.call('chunk', { offset: 0, base64: f.raw.subarray(0, 50).toString('base64') })
  await f.call('chunk', { offset: 25, base64: f.raw.subarray(25).toString('base64') })
  const ready = await f.call('finish')
  expect(parsePresentationDeliveryBundleReceipt(ready).state).toBe('ready')
  expect(await f.call('finish')).toEqual(ready)
  const restarted = createPresentationDeliveryBundleService({ userDataPath: f.root })
  expect(
    await restarted(
      { ...f.base, operation: 'delivery_bundle_metadata' },
      new AbortController().signal,
    ),
  ).toEqual(ready)
  expect(await f.call('read', { offset: 0, length: 128 * 1024 })).toEqual({
    bundleId: f.base.bundleId,
    offset: 0,
    totalBytes: f.raw.length,
    base64: f.raw.toString('base64'),
  })
  const { bundleId, ...identity } = f.base
  expect(
    await restarted(
      { ...identity, operation: 'delivery_bundle_list' },
      new AbortController().signal,
    ),
  ).toEqual({ bundles: [ready] })
  expect(await f.call('delete')).toEqual({ bundleId, deleted: true })
  await expect(f.call('metadata')).rejects.toThrow('not_found')
})
it('accepts eight bounded host screenshots and rejects a corrupt image before publication', async () => {
  const png = PNG.sync.write(new PNG({ width: 2, height: 2 }))
  const f = await fixture((files) => {
    for (let page = 1; page <= 8; page++) files.set(`page-${page}.png`, png)
  })
  await f.upload()
  expect(parsePresentationDeliveryBundleReceipt(await f.call('finish')).state).toBe('ready')
  const broken = await fixture((files) => {
    for (let page = 1; page <= 8; page++)
      files.set(`page-${page}.png`, page === 8 ? Buffer.from('corrupt png') : png)
  })
  await broken.upload()
  await expect(broken.call('finish')).rejects.toThrow('unsupported_file')
})
it.each([3, 20])('accepts %i contiguous host screenshots through PC publication', async (count) => {
  const png = PNG.sync.write(new PNG({ width: 2, height: 2 }))
  const f = await fixture((files) => {
    for (let page = 1; page <= count; page++) files.set(`page-${page}.png`, png)
  })
  await f.upload()
  const ready = parsePresentationDeliveryBundleReceipt(await f.call('finish'))
  expect(ready.state).toBe('ready')
  expect(ready.manifest.files.filter((file) => /^page-\d+\.png$/.test(file.name))).toHaveLength(
    count,
  )
})
it('rejects an invalid screenshot on the twentieth page before PC publication', async () => {
  const png = PNG.sync.write(new PNG({ width: 2, height: 2 }))
  const f = await fixture((files) => {
    for (let page = 1; page <= 20; page++)
      files.set(`page-${page}.png`, page === 20 ? Buffer.from('corrupt png') : png)
  })
  await f.upload()
  await expect(f.call('finish')).rejects.toThrow('unsupported_file')
})
it('rejects overlap conflict, mismatched identity and frozen evidence before publication', async () => {
  const f = await fixture((files) => {
    const v = JSON.parse(files.get('evidence.json')!.toString())
    v.inputDigest = 'a'.repeat(64)
    files.set('evidence.json', Buffer.from(JSON.stringify(v)))
  })
  await f.upload()
  await expect(
    f.call('chunk', { offset: 0, base64: Buffer.from('conflict').toString('base64') }),
  ).rejects.toThrow('attachment_conflict')
  await expect(f.call('finish')).rejects.toThrow('invalid_state')
  expect(parsePresentationDeliveryBundleReceipt(await f.call('metadata')).state).toBe('uploading')
  await expect(f.call('metadata', { requestId: 'other' })).rejects.toThrow('not_found')
  await expect(
    f.call('begin', {
      sha256: f.base.bundleId,
      sizeBytes: f.raw.length,
      manifest: { ...f.manifest, documentId: 'other' },
    }),
  ).rejects.toThrow('invalid_state')
})
it('rejects corrupt cache safely without changing existing bundle', async () => {
  const f = await fixture()
  await f.upload()
  await f.call('finish')
  const dir = join(
    f.root,
    'presentation-delivery-bundles',
    hash('doc'),
    hash(f.base.projectId),
    hash(f.base.bundleId),
  )
  writeFileSync(join(dir, 'bundle.zip'), 'changed')
  await expect(f.call('metadata')).rejects.toThrow('invalid_state')
  writeFileSync(join(dir, 'metadata.json'), 'private disk details')
  await expect(f.call('metadata')).rejects.toThrow(/^invalid_state$/)
})
it('enforces reserved project capacity and rejects source symlinks', async () => {
  const f = await fixture()
  for (let i = 0; i < 5; i++)
    await f.call('begin', {
      bundleId: hash(String(i)),
      sha256: hash(String(i)),
      sizeBytes: 20 * 1024 * 1024,
      manifest: f.manifest,
    })
  await expect(f.begin()).rejects.toThrow('quota_exceeded')
  const g = await fixture()
  symlinkSync(f.root, join(g.root, 'presentation-delivery-bundles'))
  await expect(g.begin()).rejects.toThrow('invalid_state')
})
it('rejects duplicate ZIP central entries and per-file digest mismatch', async () => {
  const f = await fixture()
  const broken = Buffer.from(f.raw)
  const signatures = []
  for (let i = 0; i < broken.length - 4; i++)
    if (broken.readUInt32LE(i) === 0x02014b50) signatures.push(i)
  const source = signatures[4]!,
    target = signatures[5]!
  const nameLength = broken.readUInt16LE(source + 28)
  expect(broken.readUInt16LE(target + 28)).toBe(nameLength)
  broken.copy(broken, target + 46, source + 46, source + 46 + nameLength)
  const base = { ...f.base, bundleId: hash(broken) }
  const service = createPresentationDeliveryBundleService({ userDataPath: f.root })
  const call = (operation: string, extra: Record<string, unknown> = {}) =>
    service(
      { ...base, operation: 'delivery_bundle_' + operation, ...extra },
      new AbortController().signal,
    )
  await call('begin', { sha256: base.bundleId, sizeBytes: broken.length, manifest: f.manifest })
  await call('chunk', { offset: 0, base64: broken.toString('base64') })
  await expect(call('finish')).rejects.toThrow('unsupported_file')
  const g = await fixture()
  g.manifest.files[0]!.sha256 = 'a'.repeat(64)
  await g.upload()
  await expect(g.call('finish')).rejects.toThrow('digest_mismatch')
})
it('bounds outstanding bundle count and classic ZIP expanded sizes', async () => {
  const f = await fixture()
  for (let i = 0; i < 32; i++)
    await f.call('begin', {
      bundleId: hash(String(i)),
      sha256: hash(String(i)),
      sizeBytes: 1,
      manifest: f.manifest,
    })
  await expect(f.begin()).rejects.toThrow('quota_exceeded')
  const g = await fixture()
  const raw = Buffer.from(g.raw)
  for (let i = 0; i < raw.length - 4; i++)
    if (raw.readUInt32LE(i) === 0x02014b50) {
      raw.writeUInt32LE(21 * 1024 * 1024, i + 24)
      break
    }
  const service = createPresentationDeliveryBundleService({ userDataPath: g.root })
  const base = { ...g.base, bundleId: hash(raw) }
  const call = (operation: string, extra: Record<string, unknown> = {}) =>
    service(
      { ...base, operation: 'delivery_bundle_' + operation, ...extra },
      new AbortController().signal,
    )
  await call('begin', { sha256: base.bundleId, sizeBytes: raw.length, manifest: g.manifest })
  await call('chunk', { offset: 0, base64: raw.toString('base64') })
  await expect(call('finish')).rejects.toThrow('unsupported_file')
})
it('rejects inconsistent manifests and ready receipts and returns detached copies', async () => {
  const f = await fixture()
  const parsed = parsePresentationDeliveryBundleManifest(f.manifest)
  parsed.files[0]!.name = 'changed'
  expect(f.manifest.files[0]!.name).toBe('presentation.pptx')
  const screenshots = Array.from({ length: 8 }, (_, index) => ({
    name: `page-${index + 1}.png`,
    sizeBytes: 128,
    sha256: hash(`page-${index + 1}`),
  }))
  expect(
    parsePresentationDeliveryBundleManifest({
      ...f.manifest,
      files: [...f.manifest.files, ...screenshots],
      checks: { ...f.manifest.checks, pageScreenshots: 'captured_unreviewed' },
    }).files,
  ).toHaveLength(16)
  expect(
    parsePresentationDeliveryBundleManifest({
      ...f.manifest,
      files: [...f.manifest.files, ...screenshots.slice(0, 3)],
      checks: { ...f.manifest.checks, pageScreenshots: 'captured_unreviewed' },
    }).files,
  ).toHaveLength(11)
  expect(() =>
    parsePresentationDeliveryBundleManifest({
      ...f.manifest,
      files: [...f.manifest.files, screenshots[0]!, screenshots[2]!],
      checks: { ...f.manifest.checks, pageScreenshots: 'captured_unreviewed' },
    }),
  ).toThrow('invalid_state')
  for (const manifest of [
    { ...f.manifest, extra: true },
    { ...f.manifest, createdAt: '2026-02-30T00:00:00.000Z' },
    { ...f.manifest, checks: { ...f.manifest.checks, completion: 'complete' } },
    { ...f.manifest, checks: { ...f.manifest.checks, pdf: 'included' } },
    { ...f.manifest, files: [...f.manifest.files.slice(1), f.manifest.files[1]] },
  ])
    expect(() => parsePresentationDeliveryBundleManifest(manifest)).toThrow('invalid_state')
  const receipt = await f.begin()
  expect(() =>
    parsePresentationDeliveryBundleReceipt({
      ...(receipt as object),
      state: 'ready',
      completedAt: new Date().toISOString(),
    }),
  ).toThrow('invalid_state')
})
it('keeps metadata read-only and restores a durable chunk through explicit exact-overlap retry', async () => {
  const f = await fixture()
  await f.begin()
  const dir = join(
    f.root,
    'presentation-delivery-bundles',
    hash('doc'),
    hash(f.base.projectId),
    hash(f.base.bundleId),
  )
  writeFileSync(join(dir, 'bundle.zip'), f.raw)
  const recovered = await f.call('metadata')
  expect(parsePresentationDeliveryBundleReceipt(recovered).receivedBytes).toBe(0)
  expect(
    parsePresentationDeliveryBundleReceipt(
      await f.call('chunk', { offset: 0, base64: f.raw.toString('base64') }),
    ).receivedBytes,
  ).toBe(f.raw.length)
  expect(parsePresentationDeliveryBundleReceipt(await f.call('finish')).state).toBe('ready')
})

it('preserves recognized crash stages and other invocation temporary files on readonly restart', async () => {
  const f = await fixture()
  await f.upload()
  const ready = await f.call('finish')
  const project = join(f.root, 'presentation-delivery-bundles', hash('doc'), hash(f.base.projectId))
  const dir = join(project, hash(f.base.bundleId))
  const uuid = '12345678-1234-4234-8234-123456789abc'
  const stage = join(project, '.tmp-' + uuid)
  mkdirSync(stage)
  writeFileSync(join(stage, 'bundle.zip'), 'orphan')
  const zipTmp = join(dir, 'bundle.zip.' + uuid + '.tmp'),
    metadataTmp = join(dir, 'metadata.json.' + uuid + '.tmp')
  writeFileSync(zipTmp, 'orphan')
  writeFileSync(metadataTmp, 'orphan')
  writeFileSync(join(dir, 'keep.txt'), 'user artifact')
  const otherProject = join(
    f.root,
    'presentation-delivery-bundles',
    hash('doc'),
    hash('other-project'),
  )
  mkdirSync(otherProject)
  const active = join(otherProject, '.tmp-' + uuid)
  mkdirSync(active)
  writeFileSync(join(active, 'bundle.zip'), 'active')
  const restart = createPresentationDeliveryBundleService({ userDataPath: f.root })
  expect(
    await restart(
      { ...f.base, operation: 'delivery_bundle_metadata' },
      new AbortController().signal,
    ),
  ).toEqual(ready)
  expect(existsSync(stage)).toBe(true)
  expect(existsSync(zipTmp)).toBe(true)
  expect(existsSync(metadataTmp)).toBe(true)
  expect(existsSync(active)).toBe(true)
  expect(existsSync(join(dir, 'keep.txt'))).toBe(true)
  expect(await f.call('read', { offset: 0, length: 128 * 1024 })).toMatchObject({
    base64: f.raw.toString('base64'),
  })
  symlinkSync(
    join(dir, 'bundle.zip'),
    join(dir, 'bundle.zip.12345678-1234-4234-8234-123456789abd.tmp'),
  )
  await expect(f.call('metadata')).rejects.toThrow('invalid_state')
  expect(existsSync(join(dir, 'bundle.zip'))).toBe(true)
})
it('accepts 4096 character document identities and independent host and PC clocks', async () => {
  const f = await fixture()
  const future = {
    ...f.manifest,
    documentId: 'd'.repeat(4096),
    createdAt: '2099-01-01T00:00:00.000Z',
  }
  expect(parsePresentationDeliveryBundleManifest(future).documentId).toHaveLength(4096)
  expect(() =>
    parsePresentationDeliveryBundleManifest({ ...future, documentId: 'd'.repeat(4097) }),
  ).toThrow('invalid_state')
  expect(
    parsePresentationDeliveryBundleReceipt({
      version: 1,
      documentId: future.documentId,
      projectId: future.projectId,
      requestId: future.requestId,
      bundleId: f.base.bundleId,
      sha256: f.base.bundleId,
      sizeBytes: f.raw.length,
      receivedBytes: 0,
      state: 'uploading',
      createdAt: '2026-09-29T00:00:00.000Z',
      manifest: future,
    }).createdAt,
  ).toBe('2026-09-29T00:00:00.000Z')
})
