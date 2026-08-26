import { createHash } from 'node:crypto'
import { access, mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { openLatexProject } from '../src/project.js'
import { SnapshotStore } from '../src/snapshot.js'

const sha = (value: string) => createHash('sha256').update(value).digest('hex')

describe('SnapshotStore', () => {
  const roots: string[] = []
  async function sandbox() {
    const root = await mkdtemp(join(tmpdir(), 'latex-snapshot-'))
    roots.push(root)
    return root
  }
  afterEach(async () =>
    Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true }))),
  )

  it('stores versioned changed text and absent markers under caller cache, then restores both', async () => {
    const root = await sandbox()
    const projectRoot = join(root, 'project')
    const cache = join(root, 'user-data-cache')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'before')
    const project = await openLatexProject(projectRoot)
    const store = new SnapshotStore(cache)
    const snapshot = await store.create('project-1', project, ['main.tex', 'new.tex'])
    await project.saveText('main.tex', 'after')
    await project.saveText('new.tex', 'created')
    await store.restore('project-1', snapshot.id, project, {
      expectedAbsentHashes: new Map([['new.tex', sha('created')]]),
    })
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('before')
    await expect(access(join(projectRoot, 'new.tex'))).rejects.toMatchObject({ code: 'ENOENT' })
    const manifest = JSON.parse(
      await readFile(join(cache, 'snapshots', snapshot.id, 'manifest.json'), 'utf8'),
    )
    expect(manifest.schemaVersion).toBe(1)
    expect(manifest.entries[1].kind).toBe('absent')
    await expect(access(join(projectRoot, '.wiswork'))).rejects.toMatchObject({ code: 'ENOENT' })
  })

  it('prunes quota while protecting the current rollback snapshot', async () => {
    const root = await sandbox()
    const projectRoot = join(root, 'project')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), '1234')
    const project = await openLatexProject(projectRoot)
    const store = new SnapshotStore(join(root, 'cache'), { maxSnapshots: 2, maxBytes: 10000 })
    const first = await store.create('project-1', project, ['main.tex'])
    await store.setCurrentRollback('project-1', first.id)
    await project.saveText('main.tex', '5678')
    const second = await store.create('project-1', project, ['main.tex'])
    await project.saveText('main.tex', '9012')
    const third = await store.create('project-1', project, ['main.tex'])
    expect((await store.list('project-1')).map((item) => item.id).sort()).toEqual(
      [first.id, third.id].sort(),
    )
    await expect(store.restore('project-1', second.id, project)).rejects.toThrow(/not found/i)
  })

  it('fails creation when quota cannot retain the protected rollback point and new snapshot', async () => {
    const root = await sandbox()
    const projectRoot = join(root, 'project')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'long content')
    const project = await openLatexProject(projectRoot)
    const store = new SnapshotStore(join(root, 'cache'), { maxSnapshots: 1, maxBytes: 10000 })
    const first = await store.create('project-1', project, ['main.tex'])
    await store.setCurrentRollback('project-1', first.id)
    await expect(store.create('project-1', project, ['main.tex'])).rejects.toThrow(/quota/i)
  })

  it('discards an unprotected tentative snapshot but never the current rollback point', async () => {
    const root = await sandbox()
    const projectRoot = join(root, 'project')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'before')
    const project = await openLatexProject(projectRoot)
    const store = new SnapshotStore(join(root, 'cache'))

    const tentative = await store.create('project-1', project, ['main.tex'])
    await expect(store.discard('project-1', tentative.id)).resolves.toBe(true)
    await expect(store.list('project-1')).resolves.toEqual([])
    await expect(store.getFileHashes('project-1', tentative.id)).rejects.toThrow(/not found/i)

    const protectedSnapshot = await store.create('project-1', project, ['main.tex'])
    await store.setCurrentRollback('project-1', protectedSnapshot.id)
    await expect(store.discard('project-1', protectedSnapshot.id)).rejects.toThrow(/rollback/i)
    await expect(store.list('project-1')).resolves.toContainEqual(protectedSnapshot)
  })
})
