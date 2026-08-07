import { createHash } from 'node:crypto'
import { access, mkdtemp, mkdir, readFile, readdir, rename, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { openLatexProject } from '../src/project.js'
import { ProposalStore } from '../src/proposal.js'
import { SnapshotStore } from '../src/snapshot.js'

const sha = (text: string) => createHash('sha256').update(text).digest('hex')

describe('conditional delete and persistent prune retry', () => {
  const roots: string[] = []

  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  async function setup() {
    const root = await mkdtemp(join(tmpdir(), 'latex-delete-prune-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    const cache = join(root, 'cache')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'before')
    return { root, projectRoot, cache }
  }

  it('preserves an external atomic replacement between delete hash validation and unlink', async () => {
    const { projectRoot, cache } = await setup()
    let injected = false
    const project = await openLatexProject(projectRoot, {
      pathHooks: {
        beforeConditionalDelete: async (path: string) => {
          if (injected || !path.endsWith('new.tex')) return
          injected = true
          const replacement = join(projectRoot, 'replacement.tex')
          await writeFile(replacement, 'external edit')
          await rename(replacement, join(projectRoot, 'new.tex'))
        },
      },
    } as never)
    const snapshots = new SnapshotStore(cache)
    const proposals = new ProposalStore(cache, snapshots)
    await proposals.create({
      id: 'delete-race',
      projectId: 'project-1',
      expiresAt: Date.now() + 10_000,
      files: [{ path: 'new.tex', beforeSha256: null, afterText: 'proposal file' }],
    })
    const { snapshotId } = await proposals.apply('delete-race', 'project-1', project)

    await expect(proposals.undo('project-1', snapshotId, project)).rejects.toThrow(
      /changed|conflict|conditional delete/i,
    )
    expect(injected).toBe(true)
    expect(await readFile(join(projectRoot, 'new.tex'), 'utf8')).toBe('external edit')
    await expect(
      access(join(cache, 'proposals', 'transactions', `undo-${snapshotId}.json`)),
    ).resolves.toBeUndefined()
  })

  it('preserves an external in-place edit after delete hash validation', async () => {
    const { projectRoot, cache } = await setup()
    let injected = false
    const project = await openLatexProject(projectRoot, {
      pathHooks: {
        beforeConditionalDelete: async (path: string) => {
          if (injected || !path.endsWith('new.tex')) return
          injected = true
          await writeFile(path, 'external in-place edit')
        },
      },
    } as never)
    const proposals = new ProposalStore(cache, new SnapshotStore(cache))
    await proposals.create({
      id: 'delete-in-place-race',
      projectId: 'project-1',
      expiresAt: Date.now() + 10_000,
      files: [{ path: 'new.tex', beforeSha256: null, afterText: 'proposal file' }],
    })
    const { snapshotId } = await proposals.apply('delete-in-place-race', 'project-1', project)

    await expect(proposals.undo('project-1', snapshotId, project)).rejects.toThrow(
      /changed|conflict/i,
    )
    expect(await readFile(join(projectRoot, 'new.tex'), 'utf8')).toBe('external in-place edit')
  })

  it('persists a failed prune and retries it from a reconstructed Store', async () => {
    const { projectRoot, cache } = await setup()
    const project = await openLatexProject(projectRoot)
    let failRemoval = true
    const options = {
      maxSnapshots: 1,
      storageHooks: {
        removeSnapshotDirectory: async (path: string) => {
          if (failRemoval) throw new Error('injected persistent prune failure')
          await rm(path, { recursive: true, force: true })
        },
      },
    }
    const firstStore = new SnapshotStore(cache, options)
    const first = await firstStore.create('project-1', project, ['main.tex'])
    await project.saveText('main.tex', 'second')
    await firstStore.create('project-1', project, ['main.tex'])

    const indexPath = join(cache, 'snapshots', 'index.json')
    const failedIndex = JSON.parse(await readFile(indexPath, 'utf8'))
    expect(failedIndex.pendingDeletes.map((item: { id: string }) => item.id)).toEqual([first.id])
    failRemoval = false
    const reconstructed = new SnapshotStore(cache, options)
    await project.saveText('main.tex', 'third')
    await reconstructed.create('project-1', project, ['main.tex'])

    await expect(access(join(cache, 'snapshots', first.id))).rejects.toMatchObject({
      code: 'ENOENT',
    })
    expect(JSON.parse(await readFile(indexPath, 'utf8')).pendingDeletes).toEqual([])
  })

  it('finishes a quarantined created-file delete from an undo journal after restart', async () => {
    const { projectRoot, cache } = await setup()
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
    const quarantineName = `.wiswork-delete-${snapshotId}-${sha('new.tex').slice(0, 16)}`
    const quarantinePath = join(projectRoot, quarantineName)
    await rename(join(projectRoot, 'new.tex'), quarantinePath)

    const transactionsRoot = join(cache, 'proposals', 'transactions')
    await mkdir(transactionsRoot, { recursive: true })
    const journalPath = join(transactionsRoot, 'undo-quarantine-crash.json')
    await writeFile(
      journalPath,
      `${JSON.stringify({
        schemaVersion: 1,
        id: 'undo-quarantine-crash',
        operation: 'undo',
        phase: 'restoring',
        projectId: 'project-1',
        projectRevision: 1,
        nextProjectRevision: 2,
        snapshotId,
        previousRollback: snapshotId,
        files: [
          {
            path: 'new.tex',
            beforeSha256: null,
            afterSha256: sha('proposal file'),
          },
        ],
      })}\n`,
    )

    const reconstructed = new ProposalStore(cache, new SnapshotStore(cache))
    await expect(reconstructed.recover('project-1', project)).resolves.toEqual({ recovered: 1 })
    await expect(access(join(projectRoot, 'new.tex'))).rejects.toMatchObject({ code: 'ENOENT' })
    await expect(access(quarantinePath)).rejects.toMatchObject({ code: 'ENOENT' })
    await expect(access(journalPath)).rejects.toMatchObject({ code: 'ENOENT' })
  })

  it('recovers when project-directory durability fails after quarantine rename', async () => {
    const { projectRoot, cache } = await setup()
    let syncAttempts = 0
    const project = await openLatexProject(projectRoot, {
      pathHooks: {
        syncConditionalDeleteDirectory: async () => {
          syncAttempts += 1
          if (syncAttempts === 1) throw new Error('injected project directory sync failure')
        },
      },
    } as never)
    const snapshots = new SnapshotStore(cache)
    const proposals = new ProposalStore(cache, snapshots)
    await proposals.create({
      id: 'delete-durability',
      projectId: 'project-1',
      expiresAt: Date.now() + 10_000,
      files: [{ path: 'new.tex', beforeSha256: null, afterText: 'proposal file' }],
    })
    const { snapshotId } = await proposals.apply('delete-durability', 'project-1', project)
    await expect(proposals.undo('project-1', snapshotId, project)).rejects.toThrow(/sync/i)
    const journalPath = join(cache, 'proposals', 'transactions', `undo-${snapshotId}.json`)
    await expect(access(journalPath)).resolves.toBeUndefined()

    const reconstructedProject = await openLatexProject(projectRoot)
    const reconstructed = new ProposalStore(cache, new SnapshotStore(cache))
    await expect(reconstructed.recover('project-1', reconstructedProject)).resolves.toEqual({
      recovered: 1,
    })
    await expect(access(join(projectRoot, 'new.tex'))).rejects.toMatchObject({ code: 'ENOENT' })
    await expect(access(journalPath)).rejects.toMatchObject({ code: 'ENOENT' })
  })

  it('keeps snapshot directory growth bounded while persistent prune keeps failing', async () => {
    const { projectRoot, cache } = await setup()
    const project = await openLatexProject(projectRoot)
    const options = {
      maxSnapshots: 1,
      storageHooks: {
        removeSnapshotDirectory: async () => {
          throw new Error('injected persistent prune failure')
        },
      },
    }
    const firstStore = new SnapshotStore(cache, options)
    await firstStore.create('project-1', project, ['main.tex'])
    await project.saveText('main.tex', 'second')
    await firstStore.create('project-1', project, ['main.tex'])
    const countAfterFailure = await snapshotDirectoryCount(cache)

    const reconstructed = new SnapshotStore(cache, options)
    await project.saveText('main.tex', 'third')
    await expect(reconstructed.create('project-1', project, ['main.tex'])).rejects.toThrow(
      /pending|prune|cleanup/i,
    )
    await expect(reconstructed.create('project-1', project, ['main.tex'])).rejects.toThrow(
      /pending|prune|cleanup/i,
    )
    expect(await snapshotDirectoryCount(cache)).toBe(countAfterFailure)
    expect(countAfterFailure).toBeLessThanOrEqual(2)
  })
})

async function snapshotDirectoryCount(cache: string): Promise<number> {
  const entries = await readdir(join(cache, 'snapshots'), { withFileTypes: true })
  return entries.filter((entry) => entry.isDirectory() && /^[a-f0-9]{32}$/.test(entry.name)).length
}
