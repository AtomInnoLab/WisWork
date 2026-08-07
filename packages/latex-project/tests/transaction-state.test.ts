import { mkdtemp, mkdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import {
  ProjectRevisionConflictError,
  ProjectTransactionState,
  type ProjectTransactionJournal,
} from '../src/transaction-state.js'

const hash = 'a'.repeat(64)

function journal(phase: 'prepared' | 'committed' = 'prepared'): ProjectTransactionJournal {
  return {
    schemaVersion: 1,
    id: 'transaction-1',
    operation: 'apply',
    phase,
    projectId: 'project-1',
    projectRevision: 0,
    nextProjectRevision: 1,
    snapshotId: 'b'.repeat(32),
    previousRollback: null,
    files: [{ path: 'main.tex', beforeSha256: null, afterSha256: hash }],
  }
}

describe('ProjectTransactionState', () => {
  const roots: string[] = []

  async function cache() {
    const root = await mkdtemp(join(tmpdir(), 'latex-transaction-state-'))
    roots.push(root)
    return root
  }

  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  it('persists strict journals and permits only prepared-to-committed apply transition', async () => {
    const state = new ProjectTransactionState(await cache())
    await state.writeJournal(journal())
    expect(await state.readJournal('transaction-1')).toEqual(journal())
    expect(await state.listJournals('project-1')).toEqual([journal()])
    await state.writeJournal(journal('committed'))
    expect((await state.readJournal('transaction-1')).phase).toBe('committed')

    await expect(
      state.writeJournal({
        ...journal('committed'),
        id: 'unsafe-journal',
        files: [{ path: '../escape.tex', beforeSha256: null, afterSha256: hash }],
      }),
    ).rejects.toThrow(/invalid/i)
    await expect(state.writeJournal(journal('prepared'))).rejects.toThrow(/phase/i)
  })

  it('advances a project revision idempotently and rejects a stale expected revision', async () => {
    const state = new ProjectTransactionState(await cache())
    expect(await state.readRevision('project-1')).toBe(0)
    await state.advanceRevision('project-1', 0, 1)
    expect(await state.readRevision('project-1')).toBe(1)
    await expect(state.advanceRevision('project-1', 0, 1)).resolves.toBeUndefined()
    await expect(state.advanceRevision('project-1', 2, 3)).rejects.toBeInstanceOf(
      ProjectRevisionConflictError,
    )
  })

  it('does not steal a live cross-Store lock and recovers a dead-owner lock', async () => {
    const root = await cache()
    const first = new ProjectTransactionState(root)
    const second = new ProjectTransactionState(root)
    await first.withProjectLock('project-1', async () => {
      await expect(second.withProjectLock('project-1', async () => undefined)).rejects.toThrow()
    })

    const locksRoot = join(root, 'proposals', 'project-locks')
    await mkdir(locksRoot, { recursive: true })
    await writeFile(
      join(locksRoot, 'project-1.lock'),
      JSON.stringify({
        schemaVersion: 1,
        pid: 2_147_483_647,
        token: 'c'.repeat(32),
        createdAt: 0,
      }),
    )
    await expect(second.withProjectLock('project-1', async () => 'recovered')).resolves.toBe(
      'recovered',
    )
  })
})
