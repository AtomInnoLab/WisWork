import { afterEach, describe, expect, it } from 'vitest'
import { mkdtemp, mkdir, readFile, rm, symlink, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { loadTabSession, saveTabSession } from '../src/main/tab-session'

const roots: string[] = []

afterEach(async () => {
  await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
})

async function fixture() {
  const root = await mkdtemp(join(tmpdir(), 'shell-tab-session-'))
  roots.push(root)
  const project = join(root, 'paper')
  await mkdir(project)
  await writeFile(join(project, 'main.tex'), '\\documentclass{article}')
  return { root, project, sessionPath: join(root, 'open-tabs.json') }
}

describe('LaTeX tab session persistence', () => {
  it('atomically restores canonical, valid, deduplicated directories and active project', async () => {
    const { root, project, sessionPath } = await fixture()
    const alias = join(root, 'alias')
    await symlink(project, alias, 'dir')
    await saveTabSession(sessionPath, {
      projectPaths: [project, alias],
      activeProjectPath: alias,
    })

    const restored = await loadTabSession(sessionPath)

    expect(restored).toEqual({ projectPaths: [project], activeProjectPath: project })
    expect(JSON.parse(await readFile(sessionPath, 'utf8'))).toMatchObject({ version: 1 })
  })

  it('drops missing and invalid directories and clears an invalid active project', async () => {
    const { root, project, sessionPath } = await fixture()
    const invalid = join(root, 'invalid')
    await mkdir(invalid)
    await writeFile(
      sessionPath,
      JSON.stringify({
        version: 1,
        projectPaths: [project, invalid, join(root, 'missing')],
        activeProjectPath: invalid,
      }),
    )

    await expect(loadTabSession(sessionPath)).resolves.toEqual({
      projectPaths: [project],
      activeProjectPath: null,
    })
  })

  it('degrades corrupted, unknown-version, and oversized files to an empty session', async () => {
    const { sessionPath } = await fixture()
    for (const content of [
      '{broken',
      JSON.stringify({ version: 2, projectPaths: [], activeProjectPath: null }),
      'x'.repeat(70 * 1024),
    ]) {
      await writeFile(sessionPath, content)
      await expect(loadTabSession(sessionPath)).resolves.toEqual({
        projectPaths: [],
        activeProjectPath: null,
      })
    }
  })

  it('never accepts non-string paths or an unbounded tab count', async () => {
    const { project, sessionPath } = await fixture()
    await writeFile(
      sessionPath,
      JSON.stringify({
        version: 1,
        projectPaths: [project, ...Array.from({ length: 80 }, () => 42)],
        activeProjectPath: project,
      }),
    )
    await expect(loadTabSession(sessionPath)).resolves.toEqual({
      projectPaths: [],
      activeProjectPath: null,
    })
  })
})
