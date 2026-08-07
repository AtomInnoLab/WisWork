import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { AtomicWriteCommittedError, atomicWriteFile } from '../src/atomic-write.js'
import { openLatexProject } from '../src/project.js'
import { SnapshotStore, type SnapshotStoreOptions } from '../src/snapshot.js'

interface SnapshotStorageHooks {
  writeIndex?: (path: string, data: Uint8Array) => Promise<void>
  removeSnapshotDirectory?: (path: string) => Promise<void>
}

type SnapshotOptionsWithHooks = SnapshotStoreOptions & { storageHooks: SnapshotStorageHooks }

describe('snapshot index failure reconciliation', () => {
  const roots: string[] = []

  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  async function setup() {
    const root = await mkdtemp(join(tmpdir(), 'latex-snapshot-index-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    const cache = join(root, 'cache')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'before')
    return { root, cache, project: await openLatexProject(projectRoot) }
  }

  it('readbacks a committed index error instead of deleting its referenced snapshot', async () => {
    const { cache, project } = await setup()
    let writes = 0
    const options: SnapshotOptionsWithHooks = {
      storageHooks: {
        writeIndex: async (path, data) => {
          writes += 1
          await atomicWriteFile(path, data)
          throw new AtomicWriteCommittedError('durability', new Error('injected index sync'))
        },
      },
    }
    const snapshots = new SnapshotStore(cache, options as never)
    const created = await snapshots.create('project-1', project, ['main.tex'])
    expect(writes).toBe(1)
    expect((await snapshots.list('project-1')).map((item) => item.id)).toEqual([created.id])
    expect(
      JSON.parse(await readFile(join(cache, 'snapshots', created.id, 'manifest.json'), 'utf8')).id,
    ).toBe(created.id)
  })

  it('keeps an unindexed orphan for retry when prune deletion fails', async () => {
    const { cache, project } = await setup()
    let removals = 0
    const options: SnapshotOptionsWithHooks = {
      maxSnapshots: 1,
      storageHooks: {
        removeSnapshotDirectory: async () => {
          removals += 1
          throw new Error('injected prune failure')
        },
      },
    }
    const snapshots = new SnapshotStore(cache, options as never)
    const first = await snapshots.create('project-1', project, ['main.tex'])
    await project.saveText('main.tex', 'after')
    const second = await snapshots.create('project-1', project, ['main.tex'])

    expect(removals).toBe(1)
    expect((await snapshots.list('project-1')).map((item) => item.id)).toEqual([second.id])
    expect(
      JSON.parse(await readFile(join(cache, 'snapshots', first.id, 'manifest.json'), 'utf8')).id,
    ).toBe(first.id)
  })

  it('keeps a newly indexed snapshot when post-commit cleanup persistence fails', async () => {
    const { cache, project } = await setup()
    let writes = 0
    const snapshots = new SnapshotStore(cache, {
      maxSnapshots: 1,
      storageHooks: {
        writeIndex: async (path, data) => {
          writes += 1
          if (writes === 3) throw new Error('injected uncommitted cleanup index failure')
          await atomicWriteFile(path, data)
        },
      },
    })
    await snapshots.create('project-1', project, ['main.tex'])
    await project.saveText('main.tex', 'after')
    const created = await snapshots.create('project-1', project, ['main.tex'])

    expect((await snapshots.list('project-1')).map((item) => item.id)).toEqual([created.id])
    expect(
      JSON.parse(await readFile(join(cache, 'snapshots', created.id, 'manifest.json'), 'utf8')).id,
    ).toBe(created.id)
  })
})
