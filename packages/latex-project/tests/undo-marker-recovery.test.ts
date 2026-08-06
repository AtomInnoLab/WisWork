import { createHash } from 'node:crypto'
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { openLatexProject } from '../src/project.js'
import { ProposalStore } from '../src/proposal.js'
import { SnapshotStore } from '../src/snapshot.js'

const sha = (text: string) => createHash('sha256').update(text).digest('hex')

describe('undo marker recovery', () => {
  const roots: string[] = []

  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  it('recognizes an already restored snapshot after marker persistence failed', async () => {
    const root = await mkdtemp(join(tmpdir(), 'latex-undo-marker-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    const cache = join(root, 'cache')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'before')
    const project = await openLatexProject(projectRoot)
    const snapshots = new SnapshotStore(cache)
    const proposals = new ProposalStore(cache, snapshots)
    await proposals.create({
      id: 'marker-failure',
      projectId: 'project-1',
      expiresAt: Date.now() + 10_000,
      files: [{ path: 'main.tex', beforeSha256: sha('before'), afterText: 'after' }],
    })
    const { snapshotId } = await proposals.apply('marker-failure', 'project-1', project)

    const undoneRoot = join(cache, 'proposals', 'undone')
    await rm(undoneRoot, { recursive: true })
    await writeFile(undoneRoot, 'block marker persistence')
    await expect(proposals.undo('project-1', snapshotId, project)).rejects.toThrow()
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('before')

    await rm(undoneRoot)
    await mkdir(undoneRoot)
    await expect(proposals.undo('project-1', snapshotId, project)).resolves.toEqual({
      snapshotId,
      restored: false,
    })
    expect(await snapshots.getCurrentRollback('project-1')).toBeUndefined()
  })
})
