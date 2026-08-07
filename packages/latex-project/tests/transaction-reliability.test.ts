import { createHash } from 'node:crypto'
import { access, mkdtemp, mkdir, readFile, rm, unlink, utimes, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { openLatexProject, type LatexProject } from '../src/project.js'
import { ProposalStore } from '../src/proposal.js'
import { SnapshotStore } from '../src/snapshot.js'

const sha = (text: string) => createHash('sha256').update(text).digest('hex')

describe('transaction reliability boundaries', () => {
  const roots: string[] = []

  async function setup() {
    const root = await mkdtemp(join(tmpdir(), 'latex-transaction-reliability-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    const cache = join(root, 'cache')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'a.tex'), 'A0')
    await writeFile(join(projectRoot, 'b.tex'), 'B0')
    return { root, projectRoot, cache, project: await openLatexProject(projectRoot) }
  }

  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  it('rechecks every baseline after snapshot creation and before any write', async () => {
    const { cache, project, projectRoot } = await setup()
    class DeletingSnapshotStore extends SnapshotStore {
      override async create(projectId: string, target: LatexProject, paths: readonly string[]) {
        const result = await super.create(projectId, target, paths)
        await unlink(join(projectRoot, 'a.tex'))
        return result
      }
    }
    const proposals = new ProposalStore(cache, new DeletingSnapshotStore(cache))
    await proposals.create({
      id: 'post-snapshot-conflict',
      projectId: 'project-1',
      expiresAt: Date.now() + 10_000,
      files: [
        { path: 'a.tex', beforeSha256: sha('A0'), afterText: 'A1' },
        { path: 'b.tex', beforeSha256: sha('B0'), afterText: 'B1' },
      ],
    })
    await expect(proposals.apply('post-snapshot-conflict', 'project-1', project)).rejects.toThrow(
      /baseline/i,
    )
    await expect(access(join(projectRoot, 'a.tex'))).rejects.toMatchObject({ code: 'ENOENT' })
    expect(await readFile(join(projectRoot, 'b.tex'), 'utf8')).toBe('B0')
  })

  it('rejects undo after a later user edit without partially restoring the batch', async () => {
    const { cache, project, projectRoot } = await setup()
    const snapshots = new SnapshotStore(cache)
    const proposals = new ProposalStore(cache, snapshots)
    await proposals.create({
      id: 'undo-conflict',
      projectId: 'project-1',
      expiresAt: Date.now() + 10_000,
      files: [
        { path: 'a.tex', beforeSha256: sha('A0'), afterText: 'A1' },
        { path: 'b.tex', beforeSha256: sha('B0'), afterText: 'B1' },
      ],
    })
    const { snapshotId } = await proposals.apply('undo-conflict', 'project-1', project)
    await project.saveText('a.tex', 'external edit')
    await expect(proposals.undo('project-1', snapshotId, project)).rejects.toThrow(
      /changed|conflict/i,
    )
    expect(await readFile(join(projectRoot, 'a.tex'), 'utf8')).toBe('external edit')
    expect(await readFile(join(projectRoot, 'b.tex'), 'utf8')).toBe('B1')
  })

  it('preflights every snapshot payload before restoring the first file', async () => {
    const { cache, project, projectRoot } = await setup()
    const snapshots = new SnapshotStore(cache)
    const snapshot = await snapshots.create('project-1', project, ['a.tex', 'b.tex'])
    await project.saveText('a.tex', 'A1')
    await project.saveText('b.tex', 'B1')
    await writeFile(join(cache, 'snapshots', snapshot.id, 'files', '1.txt'), 'corrupt')
    await expect(snapshots.restore('project-1', snapshot.id, project)).rejects.toThrow(/integrity/i)
    expect(await readFile(join(projectRoot, 'a.tex'), 'utf8')).toBe('A1')
    expect(await readFile(join(projectRoot, 'b.tex'), 'utf8')).toBe('B1')
  })

  it('rejects a rollback record with an extra unsafe path before deleting a created file', async () => {
    const { cache, project, projectRoot } = await setup()
    const snapshots = new SnapshotStore(cache)
    const proposals = new ProposalStore(cache, snapshots)
    await proposals.create({
      id: 'rollback-record',
      projectId: 'project-1',
      expiresAt: Date.now() + 10_000,
      files: [{ path: 'new.tex', beforeSha256: null, afterText: 'new' }],
    })
    const { snapshotId } = await proposals.apply('rollback-record', 'project-1', project)
    const recordPath = join(cache, 'proposals', 'rollbacks', `${snapshotId}.json`)
    const record = JSON.parse(await readFile(recordPath, 'utf8'))
    record.expectedAbsentHashes['../escape.tex'] = '0'.repeat(64)
    await writeFile(recordPath, JSON.stringify(record))
    await expect(proposals.undo('project-1', snapshotId, project)).rejects.toThrow(/rollback/i)
    expect(await readFile(join(projectRoot, 'new.tex'), 'utf8')).toBe('new')
  })

  it.each(['dead-owner', 'old-empty'])('recovers a stale snapshot index lock: %s', async (kind) => {
    const { cache, project } = await setup()
    const snapshotsRoot = join(cache, 'snapshots')
    const snapshots = new SnapshotStore(cache)
    const lockPath = join(snapshotsRoot, 'index.lock')
    await mkdir(snapshotsRoot, { recursive: true })
    if (kind === 'dead-owner') {
      await writeFile(
        lockPath,
        JSON.stringify({
          schemaVersion: 1,
          pid: 2_147_483_647,
          token: 'a'.repeat(32),
          createdAt: 0,
        }),
      )
    } else {
      await writeFile(lockPath, '')
      await utimes(lockPath, new Date(0), new Date(0))
    }
    await expect(snapshots.create('project-1', project, ['a.tex'])).resolves.toMatchObject({
      projectId: 'project-1',
    })
  })

  it('does not steal a snapshot index lock owned by the live process', async () => {
    const { cache, project } = await setup()
    const snapshotsRoot = join(cache, 'snapshots')
    const lockPath = join(snapshotsRoot, 'index.lock')
    await mkdir(snapshotsRoot, { recursive: true })
    const lock = {
      schemaVersion: 1,
      pid: process.pid,
      token: 'b'.repeat(32),
      createdAt: Date.now(),
    }
    await writeFile(lockPath, JSON.stringify(lock))
    const snapshots = new SnapshotStore(cache)
    await expect(snapshots.create('project-1', project, ['a.tex'])).rejects.toThrow()
    expect(JSON.parse(await readFile(lockPath, 'utf8'))).toEqual(lock)
  })
})
