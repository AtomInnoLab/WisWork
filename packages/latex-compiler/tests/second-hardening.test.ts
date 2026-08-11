import { EventEmitter } from 'node:events'
import { access, mkdtemp, mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { CompileQueue } from '../src/queue.js'
import {
  commitCompileGeneration,
  compileIsolated,
  killProcessTree,
  runTectonic,
  type SpawnTectonic,
  type TectonicRunError,
} from '../src/runner.js'

class FakeChild extends EventEmitter {
  stdout = new EventEmitter()
  stderr = new EventEmitter()
  pid = 5151
  kill = vi.fn()
}

describe('second compiler hardening pass', () => {
  const roots: string[] = []
  afterEach(async () =>
    Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true }))),
  )

  async function sandbox() {
    const root = await mkdtemp(join(tmpdir(), 'latex-second-hardening-'))
    roots.push(root)
    return root
  }

  it('uses an absolute trusted taskkill path and rejects a failed taskkill', async () => {
    const child = new FakeChild()
    const killer = new EventEmitter()
    const spawn = vi.fn(() => killer)
    const pending = killProcessTree(child, child.pid, {
      platform: 'win32',
      systemRoot: 'C:\\Windows',
      spawn,
    })
    expect(spawn).toHaveBeenCalledWith(
      'C:\\Windows\\System32\\taskkill.exe',
      ['/pid', '5151', '/t', '/f'],
      expect.objectContaining({ shell: false }),
    )
    killer.emit('close', 1)
    await expect(pending).rejects.toThrow(/taskkill|exit/i)
  })

  it('bounds a hanging kill attempt and reports unconfirmed termination', async () => {
    const child = new FakeChild()
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
      killTree: async () => new Promise(() => undefined),
      totalTimeoutMs: 2,
      idleTimeoutMs: 1_000,
      killAttemptTimeoutMs: 3,
      killGraceMs: 3,
    })
    const error = (await pending.catch((value: TectonicRunError) => value)) as TectonicRunError
    expect(error.code).toBe('TECTONIC_TOTAL_TIMEOUT')
    expect(error.terminationConfirmed).toBe(false)
  })

  it('quarantines a workspace when process-tree exit cannot be confirmed', async () => {
    const root = await sandbox()
    const project = join(root, 'project')
    await mkdir(project)
    await writeFile(join(project, 'main.tex'), 'source')
    const child = new FakeChild()
    const pending = compileIsolated({
      projectDirectory: project,
      temporaryRoot: join(root, 'job'),
      cacheDirectory: join(root, 'cache'),
      mainFile: 'main.tex',
      executable: '/app/tectonic',
      bundlePath: '/cache/bundle.ttb',
      spawn: (() => child) as SpawnTectonic,
      killTree: async () => new Promise(() => undefined),
      totalTimeoutMs: 2,
      idleTimeoutMs: 1_000,
      killAttemptTimeoutMs: 3,
      killGraceMs: 3,
    })
    const error = (await pending.catch((value: TectonicRunError) => value)) as TectonicRunError
    expect(error.terminationConfirmed).toBe(false)
    expect(error.quarantinedWorkspace).toMatch(/job-/)
    await expect(access(error.quarantinedWorkspace!)).resolves.toBeUndefined()
  })

  it('kills a compile that exceeds output-directory limits while still running', async () => {
    const root = await sandbox()
    const outputDirectory = join(root, 'output')
    await mkdir(outputDirectory)
    const child = new FakeChild()
    const pending = runTectonic({
      executable: '/app/tectonic',
      bundlePath: '/cache/bundle.ttb',
      mainFile: 'main.tex',
      workspace: {
        root,
        inputDirectory: root,
        outputDirectory,
        mainFile: 'main.tex',
        cleanup: async () => undefined,
      },
      spawn: (() => child) as SpawnTectonic,
      killTree: async () => undefined,
      maxOutputArtifactBytes: 3,
      maxOutputDirectoryBytes: 4,
      maxOutputEntries: 2,
      outputPollIntervalMs: 2,
      killGraceMs: 20,
    })
    await writeFile(join(outputDirectory, 'large.pdf'), '12345')
    setTimeout(() => child.emit('close', null, 'SIGKILL'), 10)
    await expect(pending).rejects.toMatchObject({ code: 'TECTONIC_OUTPUT_LIMIT' })
  })

  it('prunes old non-current generations and readbacks a committed pointer after fsync failure', async () => {
    const root = await sandbox()
    const project = join(root, 'project')
    const cache = join(root, 'cache')
    await mkdir(project)
    await writeFile(join(project, 'main.tex'), 'source')
    const make = async (text: string) =>
      compileIsolated({
        projectDirectory: project,
        temporaryRoot: join(root, 'job'),
        cacheDirectory: cache,
        mainFile: 'main.tex',
        executable: '/app/tectonic',
        bundlePath: '/cache/bundle.ttb',
        run: async ({ workspace }) => {
          await writeFile(join(workspace.outputDirectory, 'main.pdf'), text)
          return { exitCode: 0, signal: null, log: text }
        },
      })
    const first = await make('one')
    await commitCompileGeneration(first, cache, { maxGenerations: 2 })
    const second = await make('two')
    await commitCompileGeneration(second, cache, { maxGenerations: 2 })
    const third = await make('three')
    await expect(
      commitCompileGeneration(third, cache, {
        maxGenerations: 2,
        syncDirectory: async (path) => {
          if (path === cache) throw new Error('injected post-pointer fsync failure')
        },
      }),
    ).resolves.toMatchObject({ generationId: third.generationId })
    const generations = await readdir(join(cache, 'generations'))
    expect(generations).toHaveLength(2)
    expect(generations).toContain(third.generationId)
    expect(JSON.parse(await readFile(join(cache, 'current.json'), 'utf8')).generationId).toBe(
      third.generationId,
    )
  })

  it('does not start a newer run until the aborted old run really settles', async () => {
    const queue = new CompileQueue<string>()
    let settleOld!: (value: string) => void
    const events: string[] = []
    const old = queue.request({
      projectId: 'p',
      revision: 'r1',
      run: async ({ signal }) =>
        new Promise((resolve) => {
          events.push('old-start')
          signal.addEventListener('abort', () => events.push('old-abort'))
          settleOld = resolve
        }),
    })
    await Promise.resolve()
    const fresh = queue.request({
      projectId: 'p',
      revision: 'r2',
      run: async () => {
        events.push('new-start')
        return 'new'
      },
    })
    await Promise.resolve()
    expect(events).toEqual(['old-start', 'old-abort'])
    settleOld('old')
    await expect(old).rejects.toMatchObject({ code: 'TECTONIC_STALE_RESULT' })
    await expect(fresh).resolves.toBe('new')
    expect(events).toEqual(['old-start', 'old-abort', 'new-start'])
  })

  it.each([
    ['a throwing tree killer', async () => Promise.reject(new Error('kill failed'))],
    [
      'a nonzero Windows taskkill',
      async (child: FakeChild) => {
        const killer = new EventEmitter()
        const pending = killProcessTree(child, child.pid, {
          platform: 'win32',
          systemRoot: 'C:\\Windows',
          spawn: () => killer,
        })
        killer.emit('close', 1)
        await pending
      },
    ],
  ])('does not confirm root close after %s', async (_label, kill) => {
    const child = new FakeChild()
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
      killTree: () => kill(child),
      totalTimeoutMs: 2,
      idleTimeoutMs: 1_000,
      killGraceMs: 20,
    })
    setTimeout(() => child.emit('close', null, 'SIGKILL'), 5)
    const error = (await pending.catch((value: TectonicRunError) => value)) as TectonicRunError
    expect(error.terminationConfirmed).toBe(false)
  })

  it('confirms termination only after a successful tree kill and root close', async () => {
    const child = new FakeChild()
    const killTree = vi.fn(async () => undefined)
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
      totalTimeoutMs: 2,
      idleTimeoutMs: 1_000,
      killGraceMs: 20,
    })
    setTimeout(() => child.emit('close', null, 'SIGKILL'), 5)
    const error = (await pending.catch((value: TectonicRunError) => value)) as TectonicRunError
    expect(error.terminationConfirmed).toBe(true)
    expect(killTree).toHaveBeenCalled()
  })

  it('quarantines after tree-kill failure even when the root closes', async () => {
    const root = await sandbox()
    const project = join(root, 'project')
    await mkdir(project)
    await writeFile(join(project, 'main.tex'), 'source')
    const child = new FakeChild()
    const pending = compileIsolated({
      projectDirectory: project,
      temporaryRoot: join(root, 'job'),
      cacheDirectory: join(root, 'cache'),
      mainFile: 'main.tex',
      executable: '/app/tectonic',
      bundlePath: '/cache/bundle.ttb',
      spawn: (() => child) as SpawnTectonic,
      killTree: async () => {
        throw new Error('kill failed')
      },
      totalTimeoutMs: 2,
      idleTimeoutMs: 1_000,
      killGraceMs: 20,
    })
    setTimeout(() => child.emit('close', null, 'SIGKILL'), 5)
    const error = (await pending.catch((value: TectonicRunError) => value)) as TectonicRunError
    expect(error.terminationConfirmed).toBe(false)
    await expect(access(error.quarantinedWorkspace!)).resolves.toBeUndefined()
  })
})
