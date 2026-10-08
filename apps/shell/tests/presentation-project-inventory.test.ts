import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { PresentationTeamStore } from '../src/main/presentation-team'
import { afterEach, expect, it, vi } from 'vitest'
import { createHash } from 'node:crypto'
import {
  mkdtempSync,
  mkdirSync,
  writeFileSync,
  rmSync,
  symlinkSync,
  truncateSync,
  renameSync,
  readFileSync,
  readdirSync,
} from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { presentationDeliveryBundleFiles } from '@wiswork/project-store/presentation-delivery-bundle'
import {
  inspectPresentationProjectInventory,
  MAX_PRESENTATION_INVENTORY_BYTES,
  MAX_PRESENTATION_INVENTORY_FILES,
} from '../src/main/presentation-project-inventory'
const openHook = vi.hoisted(() => ({
  after: undefined as undefined | ((path: string) => void),
  futureDirectory: undefined as string | undefined,
}))
vi.mock('node:fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs')>()
  return {
    ...actual,
    lstatSync: (...args: Parameters<typeof actual.lstatSync>) => {
      const value = actual.lstatSync(...args)
      if (value && String(args[0]) === openHook.futureDirectory) value.birthtimeMs = 9e15
      return value
    },
    openSync: (...args: Parameters<typeof actual.openSync>) => {
      const fd = actual.openSync(...args)
      openHook.after?.(String(args[0]))
      return fd
    },
  }
})
const roots: string[] = []
afterEach(() => {
  openHook.after = undefined
  openHook.futureDirectory = undefined
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const hash = (v: string) => createHash('sha256').update(v).digest('hex')
function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'project-inventory-'))
  roots.push(root)
  return root
}
function file(root: string, path: string, value: unknown) {
  const target = join(root, path)
  mkdirSync(join(target, '..'), { recursive: true })
  writeFileSync(target, typeof value === 'string' ? value : JSON.stringify(value))
  return target
}
function project(root: string, id = 'p', doc = 'doc') {
  file(root, `projects/presentations/${hash(id)}/project.json`, {
    version: 1,
    projectId: id,
    documentId: doc,
  })
  return `projects/presentations/${hash(id)}`
}
it('ignores sibling temporary-directory activity while checking a project inventory', async () => {
  const root = fixture()
  project(root)
  const sibling = `${root}-peer`
  roots.push(sibling)
  openHook.after = (path) => {
    if (!path.endsWith('/project.json')) return
    openHook.after = undefined
    mkdirSync(sibling)
  }
  const inventory = await inspectPresentationProjectInventory({
    userDataPath: root,
    documentId: 'doc',
    projectId: 'p',
    observeActivity: {},
  })
  expect(inventory.activity?.decidable).toBe(true)
})
it('scans actual project/research/delivery/page namespaces while retaining same-document shared originals for both projects', async () => {
  const root = fixture()
  project(root)
  project(root, 'other')
  const state = {
    version: 1,
    documentId: 'doc',
    projectId: 'p',
    revision: 0,
    totalRecords: 0,
    records: [],
  }
  file(root, `presentation-research/${hash('doc')}/${hash('p')}/state.json`, {
    state,
    checksum: hash(JSON.stringify(state)),
  })
  const time = '2026-09-29T00:00:00.000Z',
    receipt = {
      version: 1,
      documentId: 'doc',
      projectId: 'p',
      requestId: 'r',
      bundleId: hash('bundle'),
      sha256: hash('bundle'),
      sizeBytes: 10,
      receivedBytes: 0,
      state: 'uploading',
      createdAt: time,
      manifest: {
        version: 1,
        scope: 'current_office_document',
        documentId: 'doc',
        projectId: 'p',
        requestId: 'r',
        planRevision: 1,
        inputDigest: hash('input'),
        planDigest: hash('plan'),
        createdAt: time,
        files: presentationDeliveryBundleFiles.map((name) => ({
          name,
          sizeBytes: 1,
          sha256: hash(name),
        })),
        checks: {
          completion: 'not_verified',
          sourceAuthority: 'not_verified',
          timeliness: 'not_verified',
          roundTrip: 'not_run',
          hostQa: 'not_checked',
          pdf: 'not_requested',
        },
      },
    }
  file(
    root,
    `presentation-delivery-bundles/${hash('doc')}/${hash('p')}/${hash(receipt.bundleId)}/metadata.json`,
    { receipt, checksum: hash(JSON.stringify(receipt)) },
  )
  file(root, `presentation-page-backups/${hash('p')}/${hash('backup')}/metadata.json`, {
    backupId: 'backup',
    projectId: 'p',
    documentId: 'doc',
    requestId: 'r',
    parentRequestId: 'parent',
    pageId: 'page',
    hostSlideId: 'host',
    slideIds: ['host'],
    sha256: hash('bytes'),
    sizeBytes: 1,
    parentInputDigest: hash('parent'),
    inputDigest: hash('input'),
    status: 'uploading',
  })
  for (const kind of [
    'presentation-attachments',
    'presentation-existing-page-backups',
    'presentation-master-backups',
    'presentation-package-backups',
  ])
    file(root, `${kind}/${hash('doc')}/opaque/raw.bin`, 'original retained bytes')
  const a = await inspectPresentationProjectInventory({
      userDataPath: root,
      projectId: 'p',
      documentId: 'doc',
    }),
    b = await inspectPresentationProjectInventory({
      userDataPath: root,
      projectId: 'other',
      documentId: 'doc',
    })
  expect(a.resources.filter((r) => r.disposition === 'candidate').map((r) => r.kind)).toEqual(
    expect.arrayContaining(['project', 'research', 'delivery_bundles', 'page_backups']),
  )
  for (const kind of [
    'attachments',
    'existing_page_backups',
    'master_backups',
    'package_backups',
  ]) {
    expect(a.resources.find((r) => r.kind === kind)).toMatchObject({
      ownership: 'document_shared',
      disposition: 'retained',
      fileCount: 1,
    })
    expect(b.resources.find((r) => r.kind === kind)).toMatchObject({
      ownership: 'document_shared',
      disposition: 'retained',
      fileCount: 1,
    })
  }
  expect(JSON.stringify(a)).not.toContain(root)
  expect(JSON.stringify(a)).not.toContain('original retained bytes')
  expect(JSON.stringify(a)).not.toContain(hash('doc'))
  expect(a.scope).toBe('known_presentation_namespaces')
  expect(a.complete).toBe(true)
  expect(
    await inspectPresentationProjectInventory({
      userDataPath: root,
      projectId: 'p',
      documentId: 'doc',
    }),
  ).toEqual(a)
})
it.each(['project', 'research', 'delivery', 'page', 'preferences'])(
  'refuses wrong internal %s ownership instead of trusting hashed paths',
  async (kind) => {
    const root = fixture()
    const dir = project(root)
    const wrong = { version: 1, documentId: 'other-doc', projectId: 'p' }
    if (kind === 'project') file(root, `${dir}/project.json`, wrong)
    if (kind === 'research')
      file(root, `presentation-research/${hash('doc')}/${hash('p')}/state.json`, {
        state: wrong,
        checksum: hash(JSON.stringify(wrong)),
      })
    if (kind === 'delivery')
      file(root, `presentation-delivery-bundles/${hash('doc')}/${hash('p')}/b/metadata.json`, {
        receipt: wrong,
        checksum: hash(JSON.stringify(wrong)),
      })
    if (kind === 'page') file(root, `presentation-page-backups/${hash('p')}/b/metadata.json`, wrong)
    if (kind === 'preferences')
      file(root, `presentation-preferences/${hash(JSON.stringify(['doc', 'p']))}.json`, {
        ...wrong,
        preferences: [],
      })
    await expect(
      inspectPresentationProjectInventory({
        userDataPath: root,
        projectId: 'p',
        documentId: 'doc',
      }),
    ).rejects.toThrow('presentation_inventory_invalid')
  },
)
it('refuses symlink namespaces and leaves, traversal scope IDs, oversized total bytes and cancellation', async () => {
  const root = fixture()
  project(root)
  const foreign = fixture()
  file(foreign, 'secret.txt', 'foreign')
  symlinkSync(foreign, join(root, 'presentation-package-backups'))
  await expect(
    inspectPresentationProjectInventory({ userDataPath: root, projectId: 'p', documentId: 'doc' }),
  ).rejects.toThrow('presentation_inventory_invalid')
  rmSync(join(root, 'presentation-package-backups'))
  const path = file(root, `presentation-attachments/${hash('doc')}/oversize.bin`, 'x')
  truncateSync(path, MAX_PRESENTATION_INVENTORY_BYTES + 1)
  await expect(
    inspectPresentationProjectInventory({ userDataPath: root, projectId: 'p', documentId: 'doc' }),
  ).rejects.toThrow('presentation_inventory_budget')
  await expect(
    inspectPresentationProjectInventory({
      userDataPath: root,
      projectId: '../p',
      documentId: 'doc',
    }),
  ).rejects.toThrow('invalid_request')
  const signal = AbortSignal.abort()
  await expect(
    inspectPresentationProjectInventory({
      userDataPath: root,
      projectId: 'p',
      documentId: 'doc',
      signal,
    }),
  ).rejects.toThrow('cancelled')
})
it('retains acquisition, release receipts and lifecycle controls; reads only explicitly referenced global brands', async () => {
  const root = fixture(),
    dir = project(root)
  const history = {
    version: 1,
    scope: 'remote_material_acquisition',
    documentId: 'doc',
    revision: 0,
    totalAttempts: 0,
    records: [],
  }
  file(root, `presentation-acquisition-history/${hash('doc')}.json`, {
    history,
    checksum: hash(JSON.stringify(history)),
  })
  file(root, `presentation-existing-page-backups/.released/${hash('doc')}/receipt.json`, {
    documentId: 'doc',
  })
  file(root, `presentation-project-lifecycles/${hash('p')}/lifecycle.json`, {
    version: 1,
    documentId: 'doc',
    projectId: 'p',
  })
  file(root, `${dir}/plan.json`, {
    version: 1,
    documentId: 'doc',
    projectId: 'p',
    plan: { brandKit: { id: 'brand' } },
  })
  file(root, `presentation-brand-kits/${hash('brand')}.json`, {
    version: 1,
    id: 'brand',
    revisions: [],
  })
  file(
    root,
    `presentation-brand-kits/${hash('unreferenced')}.json`,
    'corrupt unrelated global content',
  )
  const result = await inspectPresentationProjectInventory({
    userDataPath: root,
    documentId: 'doc',
    projectId: 'p',
  })
  for (const kind of [
    'acquisition_history',
    'existing_page_releases',
    'lifecycle_control',
    'brand_reference',
  ])
    expect(result.resources.find((r) => r.kind === kind)?.disposition).toBe('retained')
  expect(result.resources.filter((r) => r.kind === 'brand_reference')).toHaveLength(1)
})
it('reports orphaned scoped resources and unbound root staging as retained unproven', async () => {
  const root = fixture()
  file(root, `presentation-preferences/${hash(JSON.stringify(['doc', 'p']))}.json`, {
    version: 1,
    documentId: 'doc',
    projectId: 'p',
    preferences: [],
  })
  file(root, 'presentation-page-backups/.tmp-orphan/raw.pptx', 'partial bytes')
  const result = await inspectPresentationProjectInventory({
    userDataPath: root,
    documentId: 'doc',
    projectId: 'p',
  })
  expect(result.resources).toHaveLength(2)
  expect(
    result.resources.every((r) => r.ownership === 'unproven' && r.disposition === 'retained'),
  ).toBe(true)
})
it('refuses a symlink leaf and a directory exceeding the fixed scan-entry budget', async () => {
  const root = fixture()
  project(root)
  const foreign = file(root, 'foreign.bin', 'secret')
  const path = join(root, 'presentation-attachments', hash('doc'))
  mkdirSync(path, { recursive: true })
  symlinkSync(foreign, join(path, 'linked.bin'))
  await expect(
    inspectPresentationProjectInventory({ userDataPath: root, documentId: 'doc', projectId: 'p' }),
  ).rejects.toThrow('presentation_inventory_invalid')
  rmSync(join(path, 'linked.bin'))
  for (let i = 0; i <= MAX_PRESENTATION_INVENTORY_FILES; i++)
    writeFileSync(join(path, String(i)), '')
  await expect(
    inspectPresentationProjectInventory({ userDataPath: root, documentId: 'doc', projectId: 'p' }),
  ).rejects.toThrow('presentation_inventory_budget')
})

it('retains real team published ledgers without treating owner ACL data as private-project deletion permission', async () => {
  const root = fixture(),
    plan = benchmarkPlan(),
    documentId = 'doc',
    projectId = plan.projectId
  project(root, projectId, documentId)
  const owner = 'a'.repeat(64),
    time = '2026-09-29T00:00:00.000Z',
    teamId = 'team_' + hash(JSON.stringify([owner, documentId, projectId]))
  new PresentationTeamStore(root).write(
    {
      version: 1,
      teamId,
      documentId,
      projectId,
      ownerSubject: owner,
      revision: 1,
      createdAt: time,
      updatedAt: time,
      publishedPlan: { revision: 1, plan },
      members: [],
      comments: [],
    },
    undefined,
  )
  const result = await inspectPresentationProjectInventory({
    userDataPath: root,
    documentId,
    projectId,
  })
  expect(result.resources.find((r) => r.kind === 'teams')).toMatchObject({
    ownership: 'unproven',
    disposition: 'retained',
    fileCount: 1,
  })
  expect(result.deletionPerformed).toBe(false)
  expect(JSON.stringify(result)).not.toContain(owner)
})

it('refuses persistent parent replacement even when the opened metadata bytes are identical', async () => {
  const root = fixture(),
    dir = join(root, project(root)),
    bytes = readFileSync(join(dir, 'project.json'))
  let replaced = false
  openHook.after = (path) => {
    if (path === join(dir, 'project.json') && !replaced) {
      replaced = true
      renameSync(dir, dir + '-old')
      mkdirSync(dir)
      writeFileSync(join(dir, 'project.json'), bytes)
    }
  }
  await expect(
    inspectPresentationProjectInventory({
      userDataPath: root,
      documentId: 'doc',
      projectId: 'p',
      observeActivity: {},
    }),
  ).rejects.toThrow('presentation_inventory_invalid')
  expect(replaced).toBe(true)
})
it('accepts existing 17MiB project metadata capacity without an 8MiB preview regression', async () => {
  const root = fixture(),
    dir = project(root),
    path = join(root, dir, 'project.json'),
    json = readFileSync(path, 'utf8')
  writeFileSync(path, json + ' '.repeat(17 * 1024 * 1024 - Buffer.byteLength(json)))
  const result = await inspectPresentationProjectInventory({
    userDataPath: root,
    documentId: 'doc',
    projectId: 'p',
  })
  expect(result.resources.find((r) => r.kind === 'project')?.bytes).toBe(17 * 1024 * 1024)
})
it('accepts an existing 64MiB research state envelope using the actual checksum and ownership', async () => {
  const root = fixture()
  project(root)
  const state = {
    version: 1,
    documentId: 'doc',
    projectId: 'p',
    revision: 0,
    totalRecords: 0,
    records: [],
  }
  const path = file(root, `presentation-research/${hash('doc')}/${hash('p')}/state.json`, {
      state,
      checksum: hash(JSON.stringify(state)),
    }),
    json = readFileSync(path, 'utf8')
  writeFileSync(path, json + ' '.repeat(64 * 1024 * 1024 - Buffer.byteLength(json)))
  const result = await inspectPresentationProjectInventory({
    userDataPath: root,
    documentId: 'doc',
    projectId: 'p',
  })
  expect(result.resources.find((r) => r.kind === 'research')?.bytes).toBe(64 * 1024 * 1024)
})
it('counts both full 4096-blob document backup namespaces including every metadata and raw file', async () => {
  const root = fixture()
  project(root)
  for (const namespace of ['presentation-master-backups', 'presentation-package-backups']) {
    const directory = join(root, namespace, hash('doc'), hash('change'))
    mkdirSync(directory, { recursive: true })
    for (let i = 0; i < 4096; i++) {
      writeFileSync(join(directory, `blob-${i}.json`), JSON.stringify({ documentId: 'doc' }))
      writeFileSync(join(directory, `blob-${i}.bin`), 'x')
    }
  }
  const result = await inspectPresentationProjectInventory({
    userDataPath: root,
    documentId: 'doc',
    projectId: 'p',
  })
  expect(result.resources.find((r) => r.kind === 'master_backups')?.fileCount).toBe(8192)
  expect(result.resources.find((r) => r.kind === 'package_backups')?.fileCount).toBe(8192)
  expect(result.totals.fileCount).toBe(16385)
})

it('still refuses project and research metadata beyond their existing individual file limits', async () => {
  const root = fixture(),
    dir = project(root),
    path = join(root, dir, 'project.json'),
    json = readFileSync(path, 'utf8')
  writeFileSync(path, json + ' '.repeat(17 * 1024 * 1024 + 1 - Buffer.byteLength(json)))
  await expect(
    inspectPresentationProjectInventory({ userDataPath: root, documentId: 'doc', projectId: 'p' }),
  ).rejects.toThrow('presentation_inventory_invalid')
  writeFileSync(path, json)
  const state = {
      version: 1,
      documentId: 'doc',
      projectId: 'p',
      revision: 0,
      totalRecords: 0,
      records: [],
    },
    research = file(root, `presentation-research/${hash('doc')}/${hash('p')}/state.json`, {
      state,
      checksum: hash(JSON.stringify(state)),
    }),
    body = readFileSync(research, 'utf8')
  writeFileSync(research, body + ' '.repeat(64 * 1024 * 1024 + 1 - Buffer.byteLength(body)))
  await expect(
    inspectPresentationProjectInventory({ userDataPath: root, documentId: 'doc', projectId: 'p' }),
  ).rejects.toThrow('presentation_inventory_invalid')
})

it.each(['dynamic', 'throwing'] as const)(
  'bounds %s error message getters at the public API',
  async (kind) => {
    let reads = 0
    const error = new Error()
    Object.defineProperty(error, 'message', {
      get() {
        reads++
        if (kind === 'throwing') throw new Error('private-token=secret')
        return reads === 1 ? 'cancelled' : 'private-token=secret'
      },
    })
    const options = {
      userDataPath: fixture(),
      projectId: 'p',
      get documentId(): string {
        throw error
      },
    }
    const caught = await inspectPresentationProjectInventory(options).catch((value) => value)
    expect(caught).toBeInstanceOf(Error)
    expect(caught === error).toBe(false)
    expect(caught.message).toBe(kind === 'dynamic' ? 'cancelled' : 'presentation_inventory_invalid')
    expect(caught.cause).toBeUndefined()
    expect(reads).toBe(1)
  },
)

it('reads an empty 4096-character research document scope without creating or deleting files and rejects overflow', async () => {
  const root = fixture(),
    documentId = 'd'.repeat(4096)
  expect(
    await inspectPresentationProjectInventory({ userDataPath: root, documentId, projectId: 'p' }),
  ).toMatchObject({
    complete: true,
    deletionPerformed: false,
    resources: [],
    totals: { fileCount: 0, bytes: 0 },
  })
  expect(readdirSync(root, { recursive: true })).toEqual([])
  await expect(
    inspectPresentationProjectInventory({
      userDataPath: root,
      documentId: 'd'.repeat(4097),
      projectId: 'p',
    }),
  ).rejects.toThrow('invalid_request')
  expect(readdirSync(root, { recursive: true })).toEqual([])
})
it('internal activity observes only bound project files and excludes lifecycle control and other project', async () => {
  const root = fixture(),
    path = project(root),
    other = project(root, 'other')
  const { utimesSync } = await import('node:fs')
  const old = new Date('2020-01-01T00:00:00Z')
  utimesSync(join(root, path, 'project.json'), old, old)
  file(root, `${path}/private.bin`, 'own')
  const alien = file(root, `${other}/foreign.bin`, 'foreign')
  const before = Date.now()
  utimesSync(alien, new Date('2035-01-01'), new Date('2035-01-01'))
  const result = await inspectPresentationProjectInventory({
    userDataPath: root,
    documentId: 'doc',
    projectId: 'p',
    observeActivity: { policySavedAt: '2026-09-29T00:00:00.000Z' },
  })
  expect(result.activity).toMatchObject({
    scope: { documentId: 'doc', projectId: 'p' },
    decidable: true,
  })
  expect(Date.parse(result.activity!.latestActivityAt!)).toBeGreaterThanOrEqual(before - 1000)
  expect(result.resources.some((v) => v.kind === 'project')).toBe(true)
  expect(
    await inspectPresentationProjectInventory({
      userDataPath: root,
      documentId: 'doc',
      projectId: 'p',
    }),
  ).not.toHaveProperty('activity')
})

it('policy baseline and rollback-resistant ctime delay eligibility without including control activity', async () => {
  const root = fixture(),
    path = project(root),
    scope = { userDataPath: root, documentId: 'doc', projectId: 'p' }
  const before = await inspectPresentationProjectInventory({ ...scope, observeActivity: {} })
  const future = '2035-01-01T00:00:00.000Z'
  const policy = await inspectPresentationProjectInventory({
    ...scope,
    observeActivity: { policySavedAt: future },
  })
  expect(Date.parse(policy.activity!.latestActivityAt!)).toBeGreaterThanOrEqual(Date.parse(future))
  const ctl = file(root, `presentation-project-lifecycles/${hash('p')}/lifecycle.json`, {
    version: 1,
    projectId: 'p',
    documentId: 'doc',
  })
  const after = await inspectPresentationProjectInventory({ ...scope, observeActivity: {} })
  expect(after.activity!.latestActivityAt).toBe(before.activity!.latestActivityAt)
  expect(after.activity!.decidable).toBe(true)
  const { utimesSync } = await import('node:fs')
  utimesSync(join(root, path, 'project.json'), new Date('2020-01-01'), new Date('2020-01-01'))
  const rollback = await inspectPresentationProjectInventory({ ...scope, observeActivity: {} })
  expect(Date.parse(rollback.activity!.latestActivityAt!)).toBeGreaterThanOrEqual(
    Date.parse(before.activity!.latestActivityAt!),
  )
  expect(ctl).toBeTruthy()
})
it('shared or unbound bytes block an activity eligibility proof without body leakage', async () => {
  const root = fixture(),
    scope = { userDataPath: root, documentId: 'doc', projectId: 'p' }
  expect(
    (await inspectPresentationProjectInventory({ ...scope, observeActivity: {} })).activity,
  ).toMatchObject({ decidable: false, reason: 'unbound_project' })
  project(root)
  file(root, `presentation-attachments/${hash('doc')}/shared.bin`, 'secret private body')
  const r = await inspectPresentationProjectInventory({ ...scope, observeActivity: {} })
  expect(r.activity).toMatchObject({ decidable: false, reason: 'shared_or_unproven' })
  expect(JSON.stringify(r.activity)).not.toContain('secret')
  expect(Object.keys(r.activity!)).toEqual(['scope', 'decidable', 'latestActivityAt', 'reason'])
  await expect(
    inspectPresentationProjectInventory({
      ...scope,
      observeActivity: { policySavedAt: 'not-a-date' },
    }),
  ).rejects.toThrow('invalid_request')
  await expect(
    inspectPresentationProjectInventory({
      ...scope,
      observeActivity: { policySavedAt: '2026-02-30T00:00:00.000Z' },
    }),
  ).rejects.toThrow('invalid_request')
})

it('returns undecidable for an unrepresentable directory timestamp and does not coerce policy objects', async () => {
  const root = fixture(),
    path = project(root),
    scope = { userDataPath: root, documentId: 'doc', projectId: 'p' }
  let coerced = 0
  const policySavedAt = {
    [Symbol.toPrimitive]() {
      coerced++
      return '2026-09-29T00:00:00.000Z'
    },
  }
  await expect(
    inspectPresentationProjectInventory({ ...scope, observeActivity: { policySavedAt } as never }),
  ).rejects.toThrow('invalid_request')
  expect(coerced).toBe(0)
  openHook.futureDirectory = join(root, path)
  const value = await inspectPresentationProjectInventory({ ...scope, observeActivity: {} })
  expect(value.activity).toMatchObject({ decidable: false, reason: 'invalid_timestamp' })
  expect(
    value.activity!.latestActivityAt === null ||
      Number.isFinite(Date.parse(value.activity!.latestActivityAt)),
  ).toBe(true)
})
it('rejects changing activity options without running accessors or losing the policy baseline', async () => {
  const root = fixture()
  project(root)
  let reads = 0
  const accessor = {
    get policySavedAt() {
      reads++
      return reads === 1 ? '2035-01-01T00:00:00.000Z' : undefined
    },
  }
  await expect(
    inspectPresentationProjectInventory({
      userDataPath: root,
      documentId: 'doc',
      projectId: 'p',
      observeActivity: accessor,
    }),
  ).rejects.toThrow('invalid_request')
  expect(reads).toBe(0)
  await expect(
    inspectPresentationProjectInventory({
      userDataPath: root,
      documentId: 'doc',
      projectId: 'p',
      observeActivity: { extra: true } as never,
    }),
  ).rejects.toThrow('invalid_request')
})

it('observes recent bound research and preference activity without another project time', async () => {
  const root = fixture()
  project(root)
  project(root, 'other')
  const base = { userDataPath: root, documentId: 'doc', projectId: 'p' }
  const state = {
    version: 1,
    documentId: 'doc',
    projectId: 'p',
    revision: 0,
    totalRecords: 0,
    records: [],
  }
  const research = file(root, `presentation-research/${hash('doc')}/${hash('p')}/state.json`, {
    state,
    checksum: hash(JSON.stringify(state)),
  })
  const { utimesSync } = await import('node:fs')
  utimesSync(research, new Date('2033-01-01T00:00:00Z'), new Date('2033-01-01T00:00:00Z'))
  const first = await inspectPresentationProjectInventory({ ...base, observeActivity: {} })
  expect(Date.parse(first.activity!.latestActivityAt!)).toBeGreaterThanOrEqual(
    Date.parse('2033-01-01T00:00:00Z'),
  )
  const pref = file(root, `presentation-preferences/${hash(JSON.stringify(['doc', 'p']))}.json`, {
    version: 1,
    documentId: 'doc',
    projectId: 'p',
    preferences: [],
  })
  utimesSync(pref, new Date('2034-01-01T00:00:00Z'), new Date('2034-01-01T00:00:00Z'))
  const second = await inspectPresentationProjectInventory({ ...base, observeActivity: {} })
  expect(Date.parse(second.activity!.latestActivityAt!)).toBeGreaterThanOrEqual(
    Date.parse('2034-01-01T00:00:00Z'),
  )
  const alien = file(
    root,
    `presentation-preferences/${hash(JSON.stringify(['doc', 'other']))}.json`,
    { version: 1, documentId: 'doc', projectId: 'other', preferences: [] },
  )
  utimesSync(alien, new Date('2035-01-01T00:00:00Z'), new Date('2035-01-01T00:00:00Z'))
  const after = await inspectPresentationProjectInventory({ ...base, observeActivity: {} })
  expect(after.activity).toEqual(second.activity)
  expect(after.activity!.decidable).toBe(true)
})
