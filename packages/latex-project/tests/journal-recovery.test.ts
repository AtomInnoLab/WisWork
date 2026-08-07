import { createHash } from 'node:crypto'
import { access, mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { openLatexProject, type LatexProject } from '../src/project.js'
import { ProposalStore } from '../src/proposal.js'
import { SnapshotStore } from '../src/snapshot.js'

const sha = (text: string) => createHash('sha256').update(text).digest('hex')

interface RecoverableProposalStore {
  recover(projectId: string, project: LatexProject): Promise<{ recovered: number }>
}

describe('persistent proposal transaction recovery', () => {
  const roots: string[] = []

  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  async function setup() {
    const root = await mkdtemp(join(tmpdir(), 'latex-journal-recovery-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    const cache = join(root, 'cache')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'a.tex'), 'A0')
    await writeFile(join(projectRoot, 'b.tex'), 'B0')
    const project = await openLatexProject(projectRoot)
    const snapshots = new SnapshotStore(cache)
    const snapshot = await snapshots.create('project-1', project, ['a.tex', 'b.tex'])
    return { root, projectRoot, cache, project, snapshots, snapshot }
  }

  async function writeJournal(cache: string, snapshotId: string, phase: 'prepared' | 'committed') {
    const root = join(cache, 'proposals', 'transactions')
    await mkdir(root, { recursive: true })
    const journal = {
      schemaVersion: 1,
      id: 'recovery-transaction',
      operation: 'apply',
      phase,
      projectId: 'project-1',
      projectRevision: 0,
      nextProjectRevision: 1,
      snapshotId,
      previousRollback: null,
      files: [
        { path: 'a.tex', beforeSha256: sha('A0'), afterSha256: sha('A1') },
        { path: 'b.tex', beforeSha256: sha('B0'), afterSha256: sha('B1') },
      ],
    }
    const path = join(root, `${journal.id}.json`)
    await writeFile(path, `${JSON.stringify(journal)}\n`)
    return path
  }

  it.each(['none-written', 'one-written', 'all-written'])(
    'restores all-before from a prepared journal after %s',
    async (stage) => {
      const { projectRoot, cache, project, snapshot } = await setup()
      if (stage !== 'none-written') await project.saveText('a.tex', 'A1')
      if (stage === 'all-written') await project.saveText('b.tex', 'B1')
      const journalPath = await writeJournal(cache, snapshot.id, 'prepared')

      const reconstructed = new ProposalStore(cache, new SnapshotStore(cache))
      await expect(
        (reconstructed as unknown as RecoverableProposalStore).recover('project-1', project),
      ).resolves.toEqual({ recovered: 1 })
      expect(await readFile(join(projectRoot, 'a.tex'), 'utf8')).toBe('A0')
      expect(await readFile(join(projectRoot, 'b.tex'), 'utf8')).toBe('B0')
      await expect(access(journalPath)).rejects.toMatchObject({ code: 'ENOENT' })
    },
  )

  it.each(['before-metadata', 'after-rollback-record', 'after-current-index'])(
    'finishes all-after from a committed journal crashed %s and preserves undo',
    async (stage) => {
      const { projectRoot, cache, project, snapshots, snapshot } = await setup()
      await project.saveText('a.tex', 'A1')
      await project.saveText('b.tex', 'B1')
      const journalPath = await writeJournal(cache, snapshot.id, 'committed')
      const rollbackRoot = join(cache, 'proposals', 'rollbacks')
      if (stage !== 'before-metadata') {
        await mkdir(rollbackRoot, { recursive: true })
        await writeFile(
          join(rollbackRoot, `${snapshot.id}.json`),
          `${JSON.stringify({
            schemaVersion: 1,
            snapshotId: snapshot.id,
            projectId: 'project-1',
            expectedAbsentHashes: {},
            expectedCurrentHashes: { 'a.tex': sha('A1'), 'b.tex': sha('B1') },
          })}\n`,
        )
      }
      if (stage === 'after-current-index') {
        await snapshots.setCurrentRollback('project-1', snapshot.id)
      }

      const reconstructedSnapshots = new SnapshotStore(cache)
      const reconstructed = new ProposalStore(cache, reconstructedSnapshots)
      await expect(
        (reconstructed as unknown as RecoverableProposalStore).recover('project-1', project),
      ).resolves.toEqual({ recovered: 1 })
      expect(await readFile(join(projectRoot, 'a.tex'), 'utf8')).toBe('A1')
      expect(await readFile(join(projectRoot, 'b.tex'), 'utf8')).toBe('B1')
      expect(await reconstructedSnapshots.getCurrentRollback('project-1')).toBe(snapshot.id)
      await expect(access(journalPath)).rejects.toMatchObject({ code: 'ENOENT' })

      await expect(reconstructed.undo('project-1', snapshot.id, project)).resolves.toMatchObject({
        restored: true,
      })
      expect(await readFile(join(projectRoot, 'a.tex'), 'utf8')).toBe('A0')
      expect(await readFile(join(projectRoot, 'b.tex'), 'utf8')).toBe('B0')
    },
  )
})
