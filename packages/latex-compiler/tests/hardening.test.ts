import { EventEmitter } from 'node:events'
import {
  access,
  mkdtemp,
  mkdir,
  readFile,
  readdir,
  rename,
  rm,
  symlink,
  writeFile,
} from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { CompileQueue } from '../src/queue.js'
import {
  commitCompileGeneration,
  compileIsolated,
  runTectonic,
  type SpawnTectonic,
  type TectonicRunError,
} from '../src/runner.js'
import { createCompileWorkspace } from '../src/workspace.js'

class DelayedCloseChild extends EventEmitter {
  stdout = new EventEmitter()
  stderr = new EventEmitter()
  pid = 9999
  kill = vi.fn()
}

describe('compiler hardening', () => {
  const roots: string[] = []
  afterEach(async () =>
    Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true }))),
  )

  async function root() {
    const value = await mkdtemp(join(tmpdir(), 'latex-hardening-'))
    roots.push(value)
    return value
  }

  it('rejects a directory replaced by an intermediate symlink after enumeration', async () => {
    const base = await root()
    const project = join(base, 'project')
    const moved = join(base, 'moved')
    await mkdir(join(project, 'chapter'), { recursive: true })
    await writeFile(join(project, 'main.tex'), 'main')
    await writeFile(join(project, 'chapter/a.tex'), 'safe')
    await expect(
      createCompileWorkspace(project, join(base, 'tmp'), {
        hooks: {
          afterDirectoryRead: async (path) => {
            if (path.endsWith('/chapter')) return
            await rename(join(project, 'chapter'), moved)
            await symlink(moved, join(project, 'chapter'))
          },
        },
      }),
    ).rejects.toThrow(/changed|link|identity/i)
  })

  it('counts bytes actually read and rejects post-read mutation or a mixed revision', async () => {
    const base = await root()
    const project = join(base, 'project')
    await mkdir(project)
    await writeFile(join(project, 'main.tex'), '1234')
    await expect(
      createCompileWorkspace(project, join(base, 'grow'), {
        maxFileBytes: 4,
        hooks: { afterFileOpen: async (path) => writeFile(path, '12345') },
      }),
    ).rejects.toThrow(/limit|changed/i)
    await writeFile(join(project, 'main.tex'), 'v1')
    await writeFile(join(project, 'other.tex'), 'v1')
    await expect(
      createCompileWorkspace(project, join(base, 'mixed'), {
        hooks: { beforeFinalValidation: async () => writeFile(join(project, 'main.tex'), 'v2') },
      }),
    ).rejects.toThrow(/changed|revision/i)
  })

  it('waits for child close after a kill request and preserves failure logs', async () => {
    const child = new DelayedCloseChild()
    const killTree = vi.fn(() => {
      throw new Error('kill raced')
    })
    const pending = runTectonic({
      executable: '/app/tectonic',
      bundlePath: '/cache/bundle.ttb',
      mainFile: 'main.tex',
      workspace: {
        root: '/tmp/job',
        inputDirectory: '/tmp/job/input',
        outputDirectory: '/tmp/job/output',
        mainFile: 'main.tex',
        cleanup: async () => undefined,
      },
      spawn: (() => child) as SpawnTectonic,
      killTree,
      killGraceMs: 50,
    })
    child.stderr.emit('data', Buffer.from('error: main.tex:4: Missing }'))
    child.emit('close', 2, null)
    const error = (await pending.then(
      () => {
        throw new Error('expected compile failure')
      },
      (value: TectonicRunError) => value,
    )) as TectonicRunError
    expect(error.code).toBe('TECTONIC_EXIT_NONZERO')
    expect(error.log).toContain('main.tex:4')
    expect(killTree).toHaveBeenCalled()
  })

  it('stages bounded artifacts and atomically switches complete generations', async () => {
    const base = await root()
    const project = join(base, 'project')
    const cache = join(base, 'cache')
    await mkdir(project)
    await writeFile(join(project, 'main.tex'), 'source')
    const stage = async (pdf: string) =>
      compileIsolated({
        projectDirectory: project,
        temporaryRoot: join(base, 'tmp'),
        cacheDirectory: cache,
        mainFile: 'main.tex',
        executable: '/app/tectonic',
        bundlePath: '/cache/bundle.ttb',
        maxArtifactBytes: 8,
        maxPublishedBytes: 16,
        run: async ({ workspace }) => {
          await writeFile(join(workspace.outputDirectory, 'main.pdf'), pdf)
          await writeFile(join(workspace.outputDirectory, 'main.synctex.gz'), 'sync')
          return { exitCode: 0, signal: null, log: 'log' }
        },
      })
    const first = await stage('old')
    await expect(access(join(cache, 'current.json'))).rejects.toMatchObject({ code: 'ENOENT' })
    const committed = await commitCompileGeneration(first, cache)
    expect(await readFile(committed.pdfPath!, 'utf8')).toBe('old')
    const second = await stage('new')
    await expect(
      commitCompileGeneration(second, cache, {
        beforePointerCommit: async () => {
          throw new Error('injected pointer failure')
        },
      }),
    ).rejects.toThrow()
    expect(JSON.parse(await readFile(join(cache, 'current.json'), 'utf8')).generationId).toBe(
      first.generationId,
    )
    expect((await readdir(join(cache, 'generations'))).length).toBeGreaterThanOrEqual(1)
  })

  it('rejects oversized generated artifacts without replacing the prior generation', async () => {
    const base = await root()
    const project = join(base, 'project')
    const cache = join(base, 'cache')
    await mkdir(project)
    await writeFile(join(project, 'main.tex'), 'source')
    await expect(
      compileIsolated({
        projectDirectory: project,
        temporaryRoot: join(base, 'tmp'),
        cacheDirectory: cache,
        mainFile: 'main.tex',
        executable: '/app/tectonic',
        bundlePath: '/cache/bundle.ttb',
        maxArtifactBytes: 2,
        run: async ({ workspace }) => {
          await writeFile(join(workspace.outputDirectory, 'main.pdf'), 'huge')
          return { exitCode: 0, signal: null, log: '' }
        },
      }),
    ).rejects.toThrow(/output|artifact|limit/i)
  })

  it('cancels by project and never calls publish for the cancelled run', async () => {
    const queue = new CompileQueue<string>()
    let finish!: (value: string) => void
    const publish = vi.fn()
    const pending = queue.request({
      projectId: 'p',
      revision: '1',
      run: async () =>
        new Promise((resolve) => {
          finish = resolve
        }),
      publish,
    })
    await Promise.resolve()
    expect(queue.cancel('p')).toBe(true)
    finish('staged-old')
    await expect(pending).rejects.toMatchObject({ code: 'TECTONIC_STALE_RESULT' })
    expect(publish).not.toHaveBeenCalled()
  })
})
