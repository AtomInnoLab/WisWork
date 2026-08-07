import { EventEmitter } from 'node:events'
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  commitCompileGeneration,
  killProcessTree,
  compileIsolated,
  runTectonic,
  type SpawnTectonic,
} from '../src/runner.js'

class FakeChild extends EventEmitter {
  stdout = new EventEmitter()
  stderr = new EventEmitter()
  pid = 4242
  kill = vi.fn()
}

describe('controlled Tectonic runner', () => {
  const roots: string[] = []
  afterEach(async () =>
    Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true }))),
  )

  async function workspace() {
    const root = await mkdtemp(join(tmpdir(), 'latex-runner-'))
    roots.push(root)
    const inputDirectory = join(root, 'input')
    const outputDirectory = join(root, 'output')
    await mkdir(inputDirectory)
    await mkdir(outputDirectory)
    await writeFile(join(inputDirectory, 'main.tex'), 'source')
    return {
      mainFile: 'main.tex',
      root,
      inputDirectory,
      outputDirectory,
      cleanup: async () => rm(root, { recursive: true, force: true }),
    }
  }

  it('uses fixed args, no shell, an isolated cwd, and a minimal untrusted environment', async () => {
    const child = new FakeChild()
    const spawn = vi.fn(() => child) as unknown as SpawnTectonic
    const ws = await workspace()
    const pending = runTectonic({
      executable: '/app/tectonic',
      bundlePath: '/cache/bundle.tar',
      mainFile: 'main.tex',
      workspace: ws,
      spawn,
    })
    child.emit('close', 0, null)
    await expect(pending).resolves.toMatchObject({ exitCode: 0 })
    expect(spawn).toHaveBeenCalledWith(
      '/app/tectonic',
      [
        'main.tex',
        '--untrusted',
        '--only-cached',
        '--synctex',
        '--bundle',
        '/cache/bundle.tar',
        '--outdir',
        ws.outputDirectory,
      ],
      expect.objectContaining({
        cwd: ws.inputDirectory,
        shell: false,
        detached: true,
        env: { LANG: 'C.UTF-8', TECTONIC_UNTRUSTED_MODE: '1' },
      }),
    )
  })

  interface FailureBehavior {
    code?: number
    totalTimeoutMs?: number
    idleTimeoutMs?: number
    maxOutputBytes?: number
    output?: string
    cancel?: boolean
  }
  it.each<[string, FailureBehavior]>([
    ['nonzero', { code: 2 }],
    ['total timeout', { totalTimeoutMs: 5 }],
    ['idle timeout', { idleTimeoutMs: 5 }],
    ['output limit', { maxOutputBytes: 3, output: '1234' }],
    ['cancel', { cancel: true }],
  ])('rejects %s and terminates the full process group', async (_label, behavior) => {
    const child = new FakeChild()
    const killTree = vi.fn()
    const ws = await workspace()
    const controller = new AbortController()
    const pending = runTectonic({
      executable: '/app/tectonic',
      bundlePath: '/cache/bundle.tar',
      mainFile: 'main.tex',
      workspace: ws,
      spawn: (() => child) as SpawnTectonic,
      killTree,
      signal: controller.signal,
      totalTimeoutMs: behavior.totalTimeoutMs ?? 1000,
      idleTimeoutMs: behavior.idleTimeoutMs ?? 1000,
      maxOutputBytes: behavior.maxOutputBytes ?? 1000,
      killGraceMs: 5,
    })
    if (behavior.output) child.stdout.emit('data', Buffer.from(behavior.output))
    if (behavior.cancel) controller.abort()
    if (behavior.code) child.emit('close', behavior.code, null)
    await expect(pending).rejects.toMatchObject({ code: expect.stringMatching(/^TECTONIC_/) })
    expect(killTree).toHaveBeenCalledWith(4242)
  })

  it('publishes only PDF, SyncTeX and clipped logs, then cleans the workspace', async () => {
    const root = await mkdtemp(join(tmpdir(), 'latex-compile-'))
    roots.push(root)
    const project = join(root, 'project')
    const cache = join(root, 'cache')
    await mkdir(project)
    await writeFile(join(project, 'main.tex'), 'source')
    const staged = await compileIsolated({
      projectDirectory: project,
      temporaryRoot: join(root, 'tmp'),
      cacheDirectory: cache,
      mainFile: 'main.tex',
      executable: '/app/tectonic',
      bundlePath: '/cache/bundle.tar',
      maxLogBytes: 4,
      run: async ({ workspace }) => {
        await writeFile(join(workspace.outputDirectory, 'main.pdf'), 'pdf')
        await writeFile(join(workspace.outputDirectory, 'main.synctex.gz'), 'sync')
        await writeFile(join(workspace.outputDirectory, 'unexpected.txt'), 'no')
        return { exitCode: 0, signal: null, log: '123456' }
      },
    })
    const result = await commitCompileGeneration(staged, cache)
    expect(await readFile(result.pdfPath!, 'utf8')).toBe('pdf')
    expect(await readFile(result.synctexPath!, 'utf8')).toBe('sync')
    expect(result.log).toBe('1234')
    expect(await readFile(result.logPath, 'utf8')).toBe('1234')
    expect(result.workspaceCleaned).toBe(true)
    expect(result.published).toHaveLength(3)
  })
  it('uses taskkill /T /F without a shell for a Windows process tree', async () => {
    const child = new FakeChild()
    const killer = new EventEmitter()
    const spawn = vi.fn(() => killer)
    const pending = killProcessTree(child, 4242, {
      platform: 'win32',
      spawn,
      systemRoot: 'C:\\Windows',
    })
    expect(spawn).toHaveBeenCalledWith(
      'C:\\Windows\\System32\\taskkill.exe',
      ['/pid', '4242', '/t', '/f'],
      expect.objectContaining({ shell: false, windowsHide: true }),
    )
    killer.emit('close', 0)
    await expect(pending).resolves.toBeUndefined()
  })
})
