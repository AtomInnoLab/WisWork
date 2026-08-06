import { createHash } from 'node:crypto'
import { access, mkdtemp, mkdir, rename, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { openLatexProject } from '../src/project.js'
import { ProposalStore } from '../src/proposal.js'
import { SnapshotStore } from '../src/snapshot.js'

const sha = (text: string) => createHash('sha256').update(text).digest('hex')

describe('conditional delete quarantine recovery', () => {
  const roots: string[] = []

  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  it('finishes a quarantined absent-marker delete from the persisted undo journal', async () => {
    const root = await mkdtemp(join(tmpdir(), 'latex-delete-quarantine-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    const cache = join(root, 'cache')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'before')
    const project = await openLatexProject(projectRoot)
    const snapshots = new SnapshotStore(cache)
    const proposals = new ProposalStore(cache, snapshots)
    await proposals.create({
      id: 'quarantine-crash',
      projectId: 'project-1',
      expiresAt: Date.now() + 10_000,
      files: [{ path: 'new.tex', beforeSha256: null, afterText: 'proposal file' }],
    })
    const { snapshotId } = await proposals.apply('quarantine-crash', 'project-1', project)

    const transactionId = `undo-${snapshotId}`
    const quarantineName = `.wiswork-delete-${snapshotId}-${createHash('sha256')
      .update('new.tex')
      .digest('hex')
      .slice(0, 16)}`
    const quarantinePath = join(projectRoot, quarantineName)
    await rename(join(projectRoot, 'new.tex'), quarantinePath)
    const transactionsRoot = join(cache, 'proposals', 'transactions')
    await mkdir(transactionsRoot, { recursive: true })
    const journalPath = join(transactionsRoot, `${transactionId}.json`)
    await writeFile(
      journalPath,
      `${JSON.stringify({
        schemaVersion: 1,
        id: transactionId,
        operation: 'undo',
        phase: 'restoring',
        projectId: 'project-1',
        projectRevision: 1,
        nextProjectRevision: 2,
        snapshotId,
        previousRollback: snapshotId,
        files: [{ path: 'new.tex', beforeSha256: null, afterSha256: sha('proposal file') }],
      })}\n`,
    )

    const reconstructed = new ProposalStore(cache, new SnapshotStore(cache))
    await expect(reconstructed.recover('project-1', project)).resolves.toEqual({ recovered: 1 })
    await expect(access(join(projectRoot, 'new.tex'))).rejects.toMatchObject({ code: 'ENOENT' })
    await expect(access(quarantinePath)).rejects.toMatchObject({ code: 'ENOENT' })
    await expect(access(journalPath)).rejects.toMatchObject({ code: 'ENOENT' })
  })
})
