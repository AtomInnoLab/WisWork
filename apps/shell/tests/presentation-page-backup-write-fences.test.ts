import { afterEach, expect, it, vi } from 'vitest'
import { createHash } from 'node:crypto'
import {
  constants,
  mkdtempSync,
  rmSync,
  existsSync,
  readFileSync,
  readdirSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { PresentationLifecycleStore } from '@wiswork/project-store'
import { createPresentationService } from '../src/main/presentation-service'
import { createPresentationPageBackupService } from '../src/main/presentation-page-backups'
import {
  capturePresentationProjectWriteLease,
  capturePresentationProjectReadLease,
} from '../src/main/presentation-project-write-lease'
import { stopPresentationProjectWork } from '../src/main/presentation-project-work'
const hooks = vi.hoisted(() => ({
  open: undefined as undefined | ((path: string, flags: unknown) => void | Promise<void>),
  stat: undefined as undefined | ((path: string, flags: unknown) => void | Promise<void>),
  rename: undefined as undefined | (() => void | Promise<void>),
  zip: undefined as undefined | (() => void | Promise<void>),
}))
vi.mock('node:fs/promises', async (load) => {
  const actual = await load<typeof import('node:fs/promises')>()
  return {
    ...actual,
    rename: async (...args: Parameters<typeof actual.rename>) => {
      await hooks.rename?.()
      return actual.rename(...args)
    },
    open: async (...args: Parameters<typeof actual.open>) => {
      const handle = await actual.open(...args)
      await hooks.open?.(String(args[0]), args[1])
      const stat = handle.stat.bind(handle)
      handle.stat = async (...values: Parameters<typeof handle.stat>) => {
        const result = await stat(...values)
        await hooks.stat?.(String(args[0]), args[1])
        return result
      }
      return handle
    },
  }
})
vi.mock('@wiswork/pptx-engine', async (load) => {
  const actual = await load<typeof import('@wiswork/pptx-engine')>()
  return {
    ...actual,
    openPptx: async (...args: Parameters<typeof actual.openPptx>) => {
      const opened = await actual.openPptx(...args)
      await hooks.zip?.()
      return opened
    },
  }
})
const roots: string[] = []
afterEach(() => {
  hooks.rename = undefined
  hooks.open = undefined
  hooks.stat = undefined
  hooks.zip = undefined
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const hash = (value: string | Buffer) => createHash('sha256').update(value).digest('hex')
async function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'page-backup-fence-'))
  roots.push(root)
  const deck = benchmarkPlannedDeck(),
    pageId = deck.slides[0]!.id,
    scope = { documentId: 'doc', projectId: deck.id },
    pc = createPresentationService({ userDataPath: root })
  const call = async (operation: string, fields = {}) =>
    JSON.parse(
      Buffer.from(
        await pc({ operation, ...scope, ...fields }, new AbortController().signal),
      ).toString(),
    )
  await call('save_plan', { expectedRevision: 0, plan: benchmarkPlan() })
  await call('production_begin', { requestId: 'parent', planRevision: 1, deck })
  await call('production_run', { requestId: 'parent' })
  await call('production_rebuild_page', {
    parentRequestId: 'parent',
    requestId: 'child',
    pageId,
    slide: { ...deck.slides[0], notes: 'revision' },
  })
  await call('production_run', { requestId: 'child' })
  const raw = Buffer.from(
      (await compilePresentationDeck({ ...deck, slides: [deck.slides[0]] })).bytes,
    ),
    lifecycle = new PresentationLifecycleStore(root),
    revision = lifecycle.read(scope)?.revision ?? lifecycle.initialize(scope).revision
  const service = createPresentationPageBackupService({
    userDataPath: root,
    captureProjectLease: ({ scope, operation, signal }) =>
      (['page_backup_status', 'page_backup_read'].includes(operation)
        ? capturePresentationProjectReadLease
        : capturePresentationProjectWriteLease)({
        store: lifecycle,
        scope,
        signal,
        readExistingProject: () => undefined,
      }),
  })
  const body = {
      backupId: 'backup',
      requestId: 'child',
      pageId,
      hostSlideId: '256',
      slideIds: ['256', '257'],
      sha256: hash(raw),
      sizeBytes: raw.length,
    },
    dir = join(root, 'presentation-page-backups', hash(scope.projectId), hash(body.backupId))
  const backup = (operation: string, fields = {}) =>
    service({ operation, ...scope, ...fields }, new AbortController().signal)
  const freeze = () =>
    lifecycle.beginDeletion(scope, revision, {
      deletionId: 'delete',
      reason: 'user',
      resources: [{ resourceId: 'pages', kind: 'page_backups', ownership: 'project_exclusive' }],
    })
  return { root, scope, raw, body, dir, backup, freeze }
}
it.each(['page_backup_status', 'page_backup_read'])(
  'keeps missing readonly namespaces absent for %s',
  async (operation) => {
    const f = await fixture()
    await expect(
      f.backup(operation, {
        backupId: 'missing',
        ...(operation === 'page_backup_read' ? { offset: 0, length: 1 } : {}),
      }),
    ).rejects.toThrow('not_found')
    expect(existsSync(join(f.root, 'presentation-page-backups'))).toBe(false)
  },
)
it('refuses begin canonical publication after an awaited staging file open sees freeze', async () => {
  const f = await fixture()
  hooks.open = (path) => {
    if (path.includes('raw.pptx.') && path.endsWith('.tmp')) {
      hooks.open = undefined
      f.freeze()
    }
  }
  await expect(f.backup('page_backup_begin', f.body)).rejects.toThrow('revision_conflict')
  expect(existsSync(f.dir)).toBe(false)
})
it('does not append raw bytes after append stat returns under a frozen lease', async () => {
  const f = await fixture()
  await f.backup('page_backup_begin', f.body)
  hooks.stat = (path, flags) => {
    if (
      path.endsWith('raw.pptx') &&
      flags === (constants.O_RDWR | constants.O_APPEND | constants.O_NOFOLLOW)
    ) {
      hooks.stat = undefined
      f.freeze()
    }
  }
  await expect(
    f.backup('page_backup_chunk', {
      backupId: f.body.backupId,
      offset: 0,
      base64: f.raw.toString('base64'),
    }),
  ).rejects.toThrow('revision_conflict')
  expect(readFileSync(join(f.dir, 'raw.pptx')).length).toBe(0)
})
it('does not mark ready after package ZIP verification completes after freeze', async () => {
  const f = await fixture()
  await f.backup('page_backup_begin', f.body)
  await f.backup('page_backup_chunk', {
    backupId: f.body.backupId,
    offset: 0,
    base64: f.raw.toString('base64'),
  })
  const before = readFileSync(join(f.dir, 'metadata.json'))
  hooks.zip = () => {
    hooks.zip = undefined
    f.freeze()
  }
  await expect(f.backup('page_backup_finish', { backupId: f.body.backupId })).rejects.toThrow(
    'revision_conflict',
  )
  expect(readFileSync(join(f.dir, 'metadata.json'))).toEqual(before)
})
it('drains the actual in-flight file open before the request settles and creates no canonical backup', async () => {
  const f = await fixture()
  let entered!: () => void, resume!: () => void
  const arrived = new Promise<void>((r) => {
      entered = r
    }),
    wait = new Promise<void>((r) => {
      resume = r
    })
  hooks.open = async (path) => {
    if (path.includes('raw.pptx.') && path.endsWith('.tmp')) {
      hooks.open = undefined
      entered()
      await wait
    }
  }
  const pending = f.backup('page_backup_begin', f.body)
  await arrived
  f.freeze()
  let settled = false
  const drained = stopPresentationProjectWork({ root: f.root, ...f.scope }).then(() => {
    settled = true
  })
  await Promise.resolve()
  await Promise.resolve()
  expect(settled).toBe(false)
  resume()
  await expect(pending).rejects.toThrow('aborted')
  await drained
  expect(existsSync(f.dir)).toBe(false)
})

it.each([
  { operation: 'page_backup_read', offset: 0, length: 0 },
  { operation: 'page_backup_chunk', offset: -1, base64: 'YWJj' },
  {
    operation: 'page_backup_chunk',
    offset: 0,
    base64: Buffer.alloc(128 * 1024 + 1).toString('base64'),
  },
])('rejects invalid operation inputs before lifecycle admission ($operation)', async (body) => {
  const root = mkdtempSync(join(tmpdir(), 'page-backup-invalid-'))
  roots.push(root)
  const capture = vi.fn(() => {
    throw Error('admission_should_not_run')
  })
  const service = createPresentationPageBackupService({
    userDataPath: root,
    captureProjectLease: capture,
  })
  await expect(
    service(
      { ...body, documentId: 'doc', projectId: 'p', backupId: 'backup' },
      new AbortController().signal,
    ),
  ).rejects.toThrow('invalid_request')
  expect(capture).not.toHaveBeenCalled()
  expect(readdirSync(root)).toEqual([])
})
it('owns caller request arrays and digest across an awaited staging open', async () => {
  const f = await fixture()
  let entered!: () => void, resume!: () => void
  const arrived = new Promise<void>((r) => {
      entered = r
    }),
    wait = new Promise<void>((r) => {
      resume = r
    })
  hooks.open = async (path) => {
    if (path.includes('raw.pptx.') && path.endsWith('.tmp')) {
      hooks.open = undefined
      entered()
      await wait
    }
  }
  const pending = f.backup('page_backup_begin', f.body)
  await arrived
  f.body.slideIds[0] = 'foreign'
  f.body.hostSlideId = 'foreign'
  f.body.backupId = 'foreign'
  f.body.sha256 = 'b'.repeat(64)
  resume()
  expect(await pending).toMatchObject({
    backupId: 'backup',
    hostSlideId: '256',
    slideIds: ['256', '257'],
    sha256: hash(f.raw),
  })
  expect(JSON.parse(readFileSync(join(f.dir, 'metadata.json'), 'utf8'))).toMatchObject({
    backupId: 'backup',
    hostSlideId: '256',
    slideIds: ['256', '257'],
    sha256: hash(f.raw),
  })
})

it('publishes with a synchronous rename and no async lease gap', async () => {
  const f = await fixture()
  await f.backup('page_backup_begin', f.body)
  await f.backup('page_backup_chunk', {
    backupId: f.body.backupId,
    offset: 0,
    base64: f.raw.toString('base64'),
  })
  const before = readFileSync(join(f.dir, 'metadata.json'))
  const asyncRename = vi.fn(() => {
    f.freeze()
  })
  hooks.rename = asyncRename
  const result = await f
    .backup('page_backup_finish', { backupId: f.body.backupId })
    .catch((error) => error)
  expect(asyncRename).not.toHaveBeenCalled()
  expect(result).toMatchObject({ status: 'ready' })
  expect(readFileSync(join(f.dir, 'metadata.json'))).not.toEqual(before)
})

it('independent: does not delete a reused temporary pathname after relinquishing it at publication', async () => {
  const f = await fixture()
  await f.backup('page_backup_begin', f.body)
  await f.backup('page_backup_chunk', {
    backupId: f.body.backupId,
    offset: 0,
    base64: f.raw.toString('base64'),
  })
  let temp = ''
  hooks.open = (path) => {
    if (path.includes('metadata.json.') && path.endsWith('.tmp')) temp = path
    else if (path === f.dir && temp) {
      writeFileSync(temp, 'foreign-invocation')
      hooks.open = undefined
    }
  }
  await f.backup('page_backup_finish', { backupId: f.body.backupId })
  expect(existsSync(temp)).toBe(true)
  expect(readFileSync(temp, 'utf8')).toBe('foreign-invocation')
})
