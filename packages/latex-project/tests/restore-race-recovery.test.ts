import { createHash } from 'node:crypto'
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { openLatexProject, type LatexProject } from '../src/project.js'
import { ProposalStore } from '../src/proposal.js'
import { SnapshotStore, type SnapshotRestoreOptions } from '../src/snapshot.js'

const sha = (text: string) => createHash('sha256').update(text).digest('hex')

interface RecoverableProposalStore {
  recover(projectId: string, project: LatexProject): Promise<{ recovered: number }>
}

describe('conditional restore and restart recovery', () => {
  const roots: string[] = []

  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  async function setup() {
    const root = await mkdtemp(join(tmpdir(), 'latex-restore-race-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    const cache = join(root, 'cache')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'a.tex'), 'A0')
    await writeFile(join(projectRoot, 'b.tex'), 'B0')
    return { root, projectRoot, cache, project: await openLatexProject(projectRoot) }
  }

  it('does not overwrite a user edit inserted after undo preflight', async () => {
    const { cache, projectRoot, project } = await setup()
    class RaceSnapshotStore extends SnapshotStore {
      injectRace = false

      override async restore(
        projectId: string,
        snapshotId: string,
        target: LatexProject,
        options: SnapshotRestoreOptions = {},
      ) {
        if (this.injectRace) {
          this.injectRace = false
          await target.saveText('b.tex', 'external edit')
        }
        return super.restore(projectId, snapshotId, target, options)
      }
    }
    const snapshots = new RaceSnapshotStore(cache)
    const proposals = new ProposalStore(cache, snapshots)
    await proposals.create({
      id: 'undo-race',
      projectId: 'project-1',
      expiresAt: Date.now() + 10_000,
      files: [
        { path: 'a.tex', beforeSha256: sha('A0'), afterText: 'A1' },
        { path: 'b.tex', beforeSha256: sha('B0'), afterText: 'B1' },
      ],
    })
    const { snapshotId } = await proposals.apply('undo-race', 'project-1', project)
    snapshots.injectRace = true
    await expect(proposals.undo('project-1', snapshotId, project)).rejects.toThrow(
      /changed|conflict/i,
    )
    expect(await readFile(join(projectRoot, 'a.tex'), 'utf8')).toBe('A1')
    expect(await readFile(join(projectRoot, 'b.tex'), 'utf8')).toBe('external edit')
  })

  it('does not overwrite a user edit inserted after rollback preflight', async () => {
    const { cache, projectRoot, project } = await setup()
    class RaceSnapshotStore extends SnapshotStore {
      override async restore(
        projectId: string,
        snapshotId: string,
        target: LatexProject,
        options: SnapshotRestoreOptions = {},
      ) {
        await target.saveText('a.tex', 'external edit')
        return super.restore(projectId, snapshotId, target, options)
      }
    }
    let writes = 0
    const proposals = new ProposalStore(cache, new RaceSnapshotStore(cache), {
      writeText: async (target, path, text) => {
        writes += 1
        await target.saveText(path, text)
        if (writes === 1) throw new Error('injected write failure')
      },
    })
    await proposals.create({
      id: 'rollback-race',
      projectId: 'project-1',
      expiresAt: Date.now() + 10_000,
      files: [
        { path: 'a.tex', beforeSha256: sha('A0'), afterText: 'A1' },
        { path: 'b.tex', beforeSha256: sha('B0'), afterText: 'B1' },
      ],
    })
    await expect(proposals.apply('rollback-race', 'project-1', project)).rejects.toThrow()
    expect(await readFile(join(projectRoot, 'a.tex'), 'utf8')).toBe('external edit')
    expect(await readFile(join(projectRoot, 'b.tex'), 'utf8')).toBe('B0')
  })

  it('continues an undo journal after restart from a partially restored batch', async () => {
    const { cache, projectRoot, project } = await setup()
    const snapshots = new SnapshotStore(cache)
    const proposals = new ProposalStore(cache, snapshots)
    await proposals.create({
      id: 'undo-restart',
      projectId: 'project-1',
      expiresAt: Date.now() + 10_000,
      files: [
        { path: 'a.tex', beforeSha256: sha('A0'), afterText: 'A1' },
        { path: 'b.tex', beforeSha256: sha('B0'), afterText: 'B1' },
      ],
    })
    const { snapshotId } = await proposals.apply('undo-restart', 'project-1', project)
    await project.saveText('a.tex', 'A0')
    const transactionsRoot = join(cache, 'proposals', 'transactions')
    await mkdir(transactionsRoot, { recursive: true })
    await writeFile(
      join(transactionsRoot, 'undo-restart.json'),
      `${JSON.stringify({
        schemaVersion: 1,
        id: 'undo-restart',
        operation: 'undo',
        phase: 'restoring',
        projectId: 'project-1',
        projectRevision: 1,
        nextProjectRevision: 2,
        snapshotId,
        previousRollback: snapshotId,
        files: [
          { path: 'a.tex', beforeSha256: sha('A0'), afterSha256: sha('A1') },
          { path: 'b.tex', beforeSha256: sha('B0'), afterSha256: sha('B1') },
        ],
      })}\n`,
    )

    const reconstructed = new ProposalStore(cache, new SnapshotStore(cache))
    await expect(
      (reconstructed as unknown as RecoverableProposalStore).recover('project-1', project),
    ).resolves.toEqual({ recovered: 1 })
    expect(await readFile(join(projectRoot, 'a.tex'), 'utf8')).toBe('A0')
    expect(await readFile(join(projectRoot, 'b.tex'), 'utf8')).toBe('B0')
  })
})
