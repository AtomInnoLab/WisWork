import { createHash } from 'node:crypto'
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { openLatexProject } from '../src/project.js'
import { ProposalStore } from '../src/proposal.js'
import { SnapshotStore } from '../src/snapshot.js'

const sha = (text: string) => createHash('sha256').update(text).digest('hex')

describe('persistent cache validation', () => {
  const roots: string[] = []

  async function setup() {
    const root = await mkdtemp(join(tmpdir(), 'latex-cache-validation-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    const cache = join(root, 'cache')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'before')
    return { root, projectRoot, cache, project: await openLatexProject(projectRoot) }
  }

  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  it('rejects project identifiers that are unsafe as index keys', async () => {
    const { cache } = await setup()
    const proposals = new ProposalStore(cache, new SnapshotStore(cache))
    await expect(
      proposals.create({
        id: 'reserved-project-id',
        projectId: '__proto__',
        expiresAt: Date.now() + 10_000,
        files: [{ path: 'main.tex', beforeSha256: sha('before'), afterText: 'after' }],
      }),
    ).rejects.toThrow(/project ID/i)
  })

  it('revalidates a persisted proposal before consuming its file paths', async () => {
    const { cache, project, projectRoot } = await setup()
    const proposals = new ProposalStore(cache, new SnapshotStore(cache))
    await proposals.create({
      id: 'tampered',
      projectId: 'project-1',
      expiresAt: Date.now() + 10_000,
      files: [{ path: 'main.tex', beforeSha256: sha('before'), afterText: 'after' }],
    })
    const path = join(cache, 'proposals', 'pending', 'tampered.json')
    const stored = JSON.parse(await readFile(path, 'utf8'))
    stored.files[0].path = '../escape.tex'
    await writeFile(path, JSON.stringify(stored))
    await expect(proposals.apply('tampered', 'project-1', project)).rejects.toThrow(
      /traversal|path/i,
    )
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('before')
  })

  it('rejects a tampered snapshot manifest before restoring any entry', async () => {
    const { cache, project, projectRoot } = await setup()
    const snapshots = new SnapshotStore(cache)
    const snapshot = await snapshots.create('project-1', project, ['main.tex'])
    const manifestPath = join(cache, 'snapshots', snapshot.id, 'manifest.json')
    const manifest = JSON.parse(await readFile(manifestPath, 'utf8'))
    manifest.entries[0].file = '../../outside.txt'
    await writeFile(manifestPath, JSON.stringify(manifest))
    await project.saveText('main.tex', 'after')
    await expect(snapshots.restore('project-1', snapshot.id, project)).rejects.toThrow(/manifest/i)
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('after')
  })
})
