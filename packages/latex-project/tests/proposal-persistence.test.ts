import { createHash } from 'node:crypto'
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { AtomicWriteCommittedError } from '../src/atomic-write.js'
import { openLatexProject } from '../src/project.js'
import { ProposalStore } from '../src/proposal.js'
import { SnapshotStore } from '../src/snapshot.js'

const sha = (text: string) => createHash('sha256').update(text).digest('hex')

describe('persistent proposal state', () => {
  const roots: string[] = []

  async function sandbox() {
    const root = await mkdtemp(join(tmpdir(), 'latex-proposal-state-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'before')
    return { root, projectRoot, project: await openLatexProject(projectRoot) }
  }

  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  it('persists consumption across ProposalStore reconstruction', async () => {
    const { root, project } = await sandbox()
    const cache = join(root, 'cache')
    const snapshots = new SnapshotStore(cache)
    const first = new ProposalStore(cache, snapshots)
    await first.create({
      id: 'persistent',
      projectId: 'project-1',
      expiresAt: Date.now() + 10_000,
      files: [{ path: 'main.tex', beforeSha256: sha('before'), afterText: 'after' }],
    })
    await first.apply('persistent', 'project-1', project)
    const reconstructed = new ProposalStore(cache, new SnapshotStore(cache))
    await expect(reconstructed.apply('persistent', 'project-1', project)).rejects.toThrow(
      /consumed/i,
    )
  })

  it('consumes an expired proposal without changing project files', async () => {
    const { root, project, projectRoot } = await sandbox()
    let now = 100
    const cache = join(root, 'cache')
    const store = new ProposalStore(cache, new SnapshotStore(cache), { now: () => now })
    await store.create({
      id: 'expires',
      projectId: 'project-1',
      expiresAt: 101,
      files: [{ path: 'main.tex', beforeSha256: sha('before'), afterText: 'after' }],
    })
    now = 102
    await expect(store.apply('expires', 'project-1', project)).rejects.toThrow(/expired/i)
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('before')
    await expect(store.apply('expires', 'project-1', project)).rejects.toThrow(/consumed/i)
  })

  it('rejects duplicate normalized file targets and unsafe cache identifiers', async () => {
    const { root } = await sandbox()
    const store = new ProposalStore(join(root, 'cache'), new SnapshotStore(join(root, 'cache')))
    await expect(
      store.create({
        id: 'duplicates',
        projectId: 'project-1',
        expiresAt: Date.now() + 10_000,
        files: [
          { path: 'chapters/./one.tex', beforeSha256: null, afterText: 'one' },
          { path: 'chapters/one.tex', beforeSha256: null, afterText: 'two' },
        ],
      }),
    ).rejects.toThrow(/duplicate/i)
    await expect(
      store.create({
        id: '../escape',
        projectId: 'project-1',
        expiresAt: Date.now() + 10_000,
        files: [{ path: 'main.tex', beforeSha256: null, afterText: 'x' }],
      }),
    ).rejects.toThrow(/ID/i)
  })

  it('keeps a committed apply when rollback-point persistence reports a committed error', async () => {
    const { root, project, projectRoot } = await sandbox()
    const cache = join(root, 'cache')
    class FailingSnapshotStore extends SnapshotStore {
      override async setCurrentRollback(projectId: string, snapshotId: string | null) {
        await super.setCurrentRollback(projectId, snapshotId)
        if (snapshotId) {
          throw new AtomicWriteCommittedError('durability', new Error('injected index sync'))
        }
      }
    }
    const snapshots = new FailingSnapshotStore(cache)
    const store = new ProposalStore(cache, snapshots)
    await store.create({
      id: 'metadata-failure',
      projectId: 'project-1',
      expiresAt: Date.now() + 10_000,
      files: [{ path: 'main.tex', beforeSha256: sha('before'), afterText: 'after' }],
    })
    const applied = await store.apply('metadata-failure', 'project-1', project)
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('after')
    expect(await snapshots.getCurrentRollback('project-1')).toBe(applied.snapshotId)
    await expect(store.undo('project-1', applied.snapshotId, project)).resolves.toMatchObject({
      restored: true,
    })
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('before')
  })
})
