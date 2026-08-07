import { createHash } from 'node:crypto'
import { access, mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { AtomicWriteCommittedError } from '../src/atomic-write.js'
import { openLatexProject, type LatexProject } from '../src/project.js'
import { ProposalStore } from '../src/proposal.js'
import { SnapshotStore } from '../src/snapshot.js'

const sha = (value: string) => createHash('sha256').update(value).digest('hex')

describe('ProposalStore transactions', () => {
  const roots: string[] = []
  async function setup(options: ConstructorParameters<typeof ProposalStore>[2] = {}) {
    const root = await mkdtemp(join(tmpdir(), 'latex-proposal-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'a.tex'), 'A0')
    await writeFile(join(projectRoot, 'b.tex'), 'B0')
    const project = await openLatexProject(projectRoot)
    const snapshots = new SnapshotStore(join(root, 'cache'))
    const proposals = new ProposalStore(join(root, 'cache'), snapshots, options)
    return { root, projectRoot, project, proposals, snapshots }
  }
  afterEach(async () =>
    Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true }))),
  )

  it('applies a multi-file proposal once and returns a protected rollback snapshot', async () => {
    const { project, projectRoot, proposals, snapshots } = await setup()
    await proposals.create({
      id: 'p1',
      projectId: 'project-1',
      expiresAt: Date.now() + 10000,
      files: [
        { path: 'a.tex', beforeSha256: sha('A0'), afterText: 'A1' },
        { path: 'new.tex', beforeSha256: null, afterText: 'N1' },
      ],
    })
    const applied = await proposals.apply('p1', 'project-1', project)
    expect(await readFile(join(projectRoot, 'a.tex'), 'utf8')).toBe('A1')
    expect(await readFile(join(projectRoot, 'new.tex'), 'utf8')).toBe('N1')
    expect(await snapshots.getCurrentRollback('project-1')).toBe(applied.snapshotId)
    await expect(proposals.apply('p1', 'project-1', project)).rejects.toThrow(/consumed/i)
  })

  it('rejects expired, conflicting, unsafe, binary, invalid UTF-8, oversized, and deletion proposals', async () => {
    const { project, proposals } = await setup({ now: () => 100, maxFileBytes: 3 })
    const invalid = [
      {
        id: 'path',
        expiresAt: 101,
        files: [{ path: '../a.tex', beforeSha256: sha('A0'), afterText: 'A1' }],
      },
      {
        id: 'binary',
        expiresAt: 101,
        files: [{ path: 'image.png', beforeSha256: null, afterText: 'x' }],
      },
      {
        id: 'utf8',
        expiresAt: 101,
        files: [{ path: 'a.tex', beforeSha256: sha('A0'), afterText: '\ud800' }],
      },
      {
        id: 'large',
        expiresAt: 101,
        files: [{ path: 'a.tex', beforeSha256: sha('A0'), afterText: '1234' }],
      },
      {
        id: 'delete',
        expiresAt: 101,
        files: [{ path: 'a.tex', beforeSha256: sha('A0'), afterText: null }],
      },
    ]
    for (const item of invalid) {
      await expect(proposals.create({ projectId: 'project-1', ...item } as never)).rejects.toThrow()
    }
    await proposals.create({
      id: 'hash',
      projectId: 'project-1',
      expiresAt: 101,
      files: [{ path: 'a.tex', beforeSha256: sha('wrong'), afterText: 'A1' }],
    })
    await expect(proposals.apply('hash', 'project-1', project)).rejects.toThrow(/baseline/i)
  })

  it.each(['ordinary', 'committed'])(
    'rolls back the whole batch after a %s partial failure',
    async (kind) => {
      let writes = 0
      const { project, projectRoot, proposals } = await setup({
        writeText: async (target: LatexProject, path: string, text: string) => {
          writes += 1
          await target.saveText(path, text)
          if (writes === 2)
            throw kind === 'committed'
              ? new AtomicWriteCommittedError('durability', new Error('sync'))
              : new Error('injected')
        },
      })
      await proposals.create({
        id: `partial-${kind}`,
        projectId: 'project-1',
        expiresAt: Date.now() + 10000,
        files: [
          { path: 'a.tex', beforeSha256: sha('A0'), afterText: 'A1' },
          { path: 'b.tex', beforeSha256: sha('B0'), afterText: 'B1' },
        ],
      })
      await expect(proposals.apply(`partial-${kind}`, 'project-1', project)).rejects.toThrow()
      expect(await readFile(join(projectRoot, 'a.tex'), 'utf8')).toBe('A0')
      expect(await readFile(join(projectRoot, 'b.tex'), 'utf8')).toBe('B0')
    },
  )

  it('undoes exactly once, restores hashes, and treats a verified repeat as idempotent', async () => {
    const { project, projectRoot, proposals } = await setup()
    await proposals.create({
      id: 'undo',
      projectId: 'project-1',
      expiresAt: Date.now() + 10000,
      files: [
        { path: 'a.tex', beforeSha256: sha('A0'), afterText: 'A1' },
        { path: 'new.tex', beforeSha256: null, afterText: 'N1' },
      ],
    })
    const { snapshotId } = await proposals.apply('undo', 'project-1', project)
    expect(await proposals.undo('project-1', snapshotId, project)).toEqual({
      snapshotId,
      restored: true,
    })
    expect(sha(await readFile(join(projectRoot, 'a.tex'), 'utf8'))).toBe(sha('A0'))
    await expect(access(join(projectRoot, 'new.tex'))).rejects.toMatchObject({ code: 'ENOENT' })
    expect(await proposals.undo('project-1', snapshotId, project)).toEqual({
      snapshotId,
      restored: false,
    })
  })
})
