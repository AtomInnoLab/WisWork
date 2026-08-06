import { createHash } from 'node:crypto'
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { openLatexProject } from '../src/project.js'
import { ProposalStore } from '../src/proposal.js'
import { SnapshotStore } from '../src/snapshot.js'

const sha = (text: string) => createHash('sha256').update(text).digest('hex')

describe('per-project proposal serialization', () => {
  const roots: string[] = []

  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  it('allows only one Store to commit proposals created at the same project revision', async () => {
    const root = await mkdtemp(join(tmpdir(), 'latex-project-lock-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    const cache = join(root, 'cache')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'a.tex'), 'A0')
    await writeFile(join(projectRoot, 'b.tex'), 'B0')
    const project = await openLatexProject(projectRoot)

    // Proposal metadata is shared, while independent snapshot stores remove unrelated index-lock
    // contention from this test. The serialization boundary must be the project revision itself.
    const first = new ProposalStore(cache, new SnapshotStore(join(root, 'snapshots-first')))
    const second = new ProposalStore(cache, new SnapshotStore(join(root, 'snapshots-second')))
    await first.create({
      id: 'same-revision-a',
      projectId: 'project-1',
      expiresAt: Date.now() + 10_000,
      files: [{ path: 'a.tex', beforeSha256: sha('A0'), afterText: 'A1' }],
    })
    await second.create({
      id: 'same-revision-b',
      projectId: 'project-1',
      expiresAt: Date.now() + 10_000,
      files: [{ path: 'b.tex', beforeSha256: sha('B0'), afterText: 'B1' }],
    })

    const results = await Promise.allSettled([
      first.apply('same-revision-a', 'project-1', project),
      second.apply('same-revision-b', 'project-1', project),
    ])
    expect(results.filter((result) => result.status === 'fulfilled')).toHaveLength(1)
    expect(results.filter((result) => result.status === 'rejected')).toHaveLength(1)

    const values = [
      await readFile(join(projectRoot, 'a.tex'), 'utf8'),
      await readFile(join(projectRoot, 'b.tex'), 'utf8'),
    ]
    expect(values.filter((value) => value.endsWith('0'))).toHaveLength(1)
    expect(values.filter((value) => value.endsWith('1'))).toHaveLength(1)
  })
})
