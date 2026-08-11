import { access, mkdtemp, mkdir, open, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { gzipSync } from 'node:zlib'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { ProjectSessionRegistry, UnsavedBuffersError } from '../src/main/project-session.js'

describe('LaTeX project sessions', () => {
  const roots: string[] = []

  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  async function setup() {
    const root = await mkdtemp(join(tmpdir(), 'latex-session-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'disk-v1')
    return { root, projectRoot }
  }

  async function setupBundleSession(bundleUrl: string) {
    const { root, projectRoot } = await setup()
    const compiler = vi.fn(async () => ({
      generationId: 'bundle-compile',
      stagingDirectory: join(root, 'bundle-stage'),
      files: [],
      log: '',
      workspaceCleaned: true as const,
    }))
    const commitGeneration = vi.fn(async () => ({
      generationId: 'bundle-compile',
      pdfPath: null,
      synctexPath: null,
      logPath: join(root, 'bundle.log'),
      log: '',
      published: [],
      workspaceCleaned: true as const,
    }))
    const session = await new ProjectSessionRegistry({
      watch: () => ({ close() {} }),
      compiler: compiler as never,
      commitGeneration: commitGeneration as never,
      compilerRuntime: {
        tectonicPath: '/app/tectonic',
        userDataPath: root,
        bundleAsset: {
          id: 'tectonic-default-bundle-v33',
          url: bundleUrl,
          bytes: 100,
          sha256: 'a'.repeat(64),
          license: { spdx: 'MIT', sourceUrl: 'https://tug.org/texlive/copying.html' },
        },
      },
    } as never).attach(11, projectRoot)
    return { root, session, compiler, commitGeneration }
  }

  it('uses the validated remote indexed tar directly instead of downloading it', async () => {
    const bundleUrl = 'https://relay.fullyjustified.net/default_bundle_v33.tar'
    const { root, session, compiler } = await setupBundleSession(bundleUrl)
    await session.compile(1, 'main.tex')
    expect(session.getBundleStatus()).toEqual({ state: 'remote' })
    expect(compiler).toHaveBeenCalledWith(
      expect.objectContaining({
        bundlePath: bundleUrl,
        tectonicCacheDirectory: join(root, 'latex', 'tectonic-cache'),
      }),
    )
  })

  it('fails closed before compile for any non-pinned bundle asset', async () => {
    const { session, compiler } = await setupBundleSession(
      'https://relay.fullyjustified.net/default_bundle_v999.tar',
    )
    expect(session.getBundleStatus()).toEqual({
      state: 'error',
      code: 'BUNDLE_NOT_CONFIGURED',
    })
    await expect(session.compile(1, 'main.tex')).rejects.toThrow(/not configured/i)
    expect(compiler).not.toHaveBeenCalled()
  })

  it('maps one project session to one WebContents and disposes owned resources on destroy', async () => {
    const { projectRoot } = await setup()
    const watcherClose = vi.fn()
    const cancelCompile = vi.fn()
    const cancelDownload = vi.fn()
    const registry = new ProjectSessionRegistry({
      watch: () => ({ close: watcherClose }),
    })
    const first = await registry.attach(11, projectRoot)
    first.trackCompile(cancelCompile)
    first.trackDownload(cancelDownload)
    await expect(registry.attach(11, projectRoot)).rejects.toThrow(/already owns/i)
    expect(registry.getOwned(11, first.projectId)).toBe(first)
    registry.destroy(11)
    expect(watcherClose).toHaveBeenCalledOnce()
    expect(cancelCompile).toHaveBeenCalledOnce()
    expect(cancelDownload).toHaveBeenCalledOnce()
    expect(registry.getByWebContents(11)).toBeUndefined()
  })

  it('refreshes clean buffers but preserves dirty text and reports external conflicts', async () => {
    const { projectRoot } = await setup()
    const registry = new ProjectSessionRegistry({ watch: () => ({ close() {} }) })
    const session = await registry.attach(11, projectRoot)
    await session.readText('main.tex')
    await writeFile(join(projectRoot, 'main.tex'), 'disk-v2')
    await session.handleExternalChange('main.tex')
    expect((await session.getBuffer('main.tex'))?.text).toBe('disk-v2')

    session.updateBuffer('main.tex', 'local-v3')
    await writeFile(join(projectRoot, 'main.tex'), 'disk-v3')
    await session.handleExternalChange('main.tex')
    expect(await session.getBuffer('main.tex')).toMatchObject({
      text: 'local-v3',
      dirty: true,
      conflict: { diskText: 'disk-v3' },
    })
  })

  it('returns authoritative baseline text for an existing dirty buffer', async () => {
    const { projectRoot } = await setup()
    const session = await new ProjectSessionRegistry({ watch: () => ({ close() {} }) }).attach(
      11,
      projectRoot,
    )
    await session.readText('main.tex')
    session.updateBuffer('main.tex', 'local-v2')
    expect(await session.readText('main.tex')).toMatchObject({
      text: 'local-v2',
      diskText: 'disk-v1',
      dirty: true,
    })
  })

  it('uses baseline hashes so save never overwrites a newly changed disk file', async () => {
    const { projectRoot } = await setup()
    const registry = new ProjectSessionRegistry({ watch: () => ({ close() {} }) })
    const session = await registry.attach(11, projectRoot)
    await session.readText('main.tex')
    session.updateBuffer('main.tex', 'local')
    await writeFile(join(projectRoot, 'main.tex'), 'external')
    await expect(session.saveText('main.tex')).rejects.toThrow(/changed|conflict/i)
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('external')
  })

  it('refuses to rename the configured main file', async () => {
    const { projectRoot } = await setup()
    const session = await new ProjectSessionRegistry({ watch: () => ({ close() {} }) }).attach(
      11,
      projectRoot,
    )
    await expect(session.renameText('main.tex', 'renamed.tex')).rejects.toThrow(/main file/i)
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('disk-v1')
  })

  it('keeps edits typed while a save is awaiting and reconciles its watcher event', async () => {
    const { projectRoot } = await setup()
    const onExternalChange = vi.fn()
    const registry = new ProjectSessionRegistry({
      watch: () => ({ close() {} }),
      onExternalChange,
    })
    const session = await registry.attach(11, projectRoot)
    await session.readText('main.tex')
    session.updateBuffer('main.tex', 'save-v2')
    const original = session.project.saveText.bind(session.project)
    let release!: () => void
    const gate = new Promise<void>((resolve) => (release = resolve))
    vi.spyOn(session.project, 'saveText').mockImplementation(async (...args) => {
      const saved = await original(...args)
      await gate
      return saved
    })
    const saving = session.saveText('main.tex')
    await vi.waitFor(async () =>
      expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('save-v2'),
    )
    session.updateBuffer('main.tex', 'typed-v3')
    await session.handleExternalChange('main.tex')
    expect(onExternalChange).toHaveBeenLastCalledWith(
      11,
      expect.objectContaining({
        text: 'save-v2',
        diskText: 'save-v2',
        dirty: false,
        conflict: null,
      }),
    )
    release()
    await saving
    expect(session.getBuffer('main.tex')).toMatchObject({
      text: 'typed-v3',
      dirty: true,
      conflict: null,
    })
    await session.saveText('main.tex')
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('typed-v3')
  })

  it('serializes same-path saves and the queued save uses latest state', async () => {
    const { projectRoot } = await setup()
    const session = await new ProjectSessionRegistry({ watch: () => ({ close() {} }) }).attach(
      11,
      projectRoot,
    )
    await session.readText('main.tex')
    const original = session.project.saveText.bind(session.project)
    let release!: () => void
    const gate = new Promise<void>((resolve) => (release = resolve))
    let active = 0
    let maxActive = 0
    let calls = 0
    vi.spyOn(session.project, 'saveText').mockImplementation(async (...args) => {
      calls += 1
      active += 1
      maxActive = Math.max(maxActive, active)
      if (calls === 1) await gate
      const result = await original(...args)
      active -= 1
      return result
    })
    session.updateBuffer('main.tex', 'v2')
    const v2 = session.saveText('main.tex')
    await vi.waitFor(() => expect(calls).toBe(1))
    session.updateBuffer('main.tex', 'v3')
    const v3 = session.saveText('main.tex')
    expect(calls).toBe(1)
    release()
    await Promise.all([v2, v3])
    expect(maxActive).toBe(1)
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('v3')
  })

  it('retains an external conflict observed after the save write but before save settles', async () => {
    const { projectRoot } = await setup()
    const session = await new ProjectSessionRegistry({ watch: () => ({ close() {} }) }).attach(
      11,
      projectRoot,
    )
    await session.readText('main.tex')
    session.updateBuffer('main.tex', 'saved-v2')
    const original = session.project.saveText.bind(session.project)
    let release!: () => void
    const gate = new Promise<void>((resolve) => (release = resolve))
    vi.spyOn(session.project, 'saveText').mockImplementation(async (...args) => {
      const saved = await original(...args)
      await gate
      return saved
    })
    const saving = session.saveText('main.tex')
    await vi.waitFor(async () =>
      expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('saved-v2'),
    )
    await writeFile(join(projectRoot, 'main.tex'), 'external-v3')
    await session.handleExternalChange('main.tex')
    release()
    await saving
    expect(session.getBuffer('main.tex')).toMatchObject({ conflict: { diskText: 'external-v3' } })
  })

  it('discard reloads the current disk text even when no watcher conflict was recorded', async () => {
    const { projectRoot } = await setup()
    const registry = new ProjectSessionRegistry({ watch: () => ({ close() {} }) })
    const session = await registry.attach(11, projectRoot)
    await session.readText('main.tex')
    session.updateBuffer('main.tex', 'local')
    await writeFile(join(projectRoot, 'main.tex'), 'disk-v2')
    await session.discardAll()
    expect(session.getBuffer('main.tex')).toMatchObject({
      text: 'disk-v2',
      dirty: false,
      conflict: null,
    })
  })

  it('rejects compile when a hidden non-main buffer is dirty without invoking compiler', async () => {
    const { root, projectRoot } = await setup()
    await writeFile(join(projectRoot, 'hidden.tex'), 'disk')
    const compiler = vi.fn()
    const session = await new ProjectSessionRegistry({
      watch: () => ({ close() {} }),
      compiler: compiler as never,
      compilerRuntime: { tectonicPath: 'fixed-tectonic', userDataPath: root },
    }).attach(11, projectRoot)
    await session.readText('hidden.tex')
    session.updateBuffer('hidden.tex', 'local')
    await expect(session.compile(1, 'main.tex')).rejects.toThrow(/unsaved/i)
    expect(compiler).not.toHaveBeenCalled()
  })

  it('rechecks hidden buffers after a queued compile request starts running', async () => {
    const { root, projectRoot } = await setup()
    await writeFile(join(projectRoot, 'hidden.tex'), 'disk')
    let calls = 0
    const compiler = vi.fn((request: { signal?: AbortSignal }) => {
      calls += 1
      if (calls === 1) {
        return new Promise((resolve, reject) => {
          request.signal!.addEventListener('abort', () => reject(new Error('cancelled')), {
            once: true,
          })
        })
      }
      return Promise.resolve({
        generationId: 'unexpected',
        stagingDirectory: join(root, 'unexpected'),
        files: [],
        log: '',
        workspaceCleaned: true,
      })
    }) as never
    const commit = vi.fn(async () => ({
      generationId: 'unexpected',
      pdfPath: null,
      synctexPath: null,
      logPath: join(root, 'log'),
      log: '',
      published: [],
      workspaceCleaned: true as const,
    }))
    const session = await new ProjectSessionRegistry({
      watch: () => ({ close() {} }),
      compiler,
      commitGeneration: commit as never,
      compilerRuntime: { tectonicPath: 'fixed-tectonic', userDataPath: root },
    }).attach(11, projectRoot)
    await session.readText('hidden.tex')
    const first = session.compile(1, 'main.tex').catch((error) => error)
    await vi.waitFor(() => expect(compiler).toHaveBeenCalledTimes(1))
    const queued = session.compile(2, 'main.tex')
    session.updateBuffer('hidden.tex', 'local')
    await expect(queued).rejects.toBeInstanceOf(UnsavedBuffersError)
    await first
    expect(compiler).toHaveBeenCalledTimes(1)
    expect(commit).not.toHaveBeenCalled()
  })

  it('serializes revision ABA compiles and publishes only the newest token', async () => {
    const { root, projectRoot } = await setup()
    const runs: {
      signal: AbortSignal
      resolve: (value: any) => void
      reject: (error: Error) => void
    }[] = []
    const compiler = vi.fn(
      (request: { signal?: AbortSignal }) =>
        new Promise<any>((resolve, reject) => {
          if (request.signal!.aborted) {
            reject(new Error('cancelled'))
            return
          }
          const run = { signal: request.signal!, resolve, reject }
          runs.push(run)
          request.signal!.addEventListener('abort', () => reject(new Error('cancelled')), {
            once: true,
          })
        }),
    ) as never
    const commit = vi.fn(async (staged: any) => ({
      generationId: staged.generationId,
      pdfPath: join(root, `${staged.generationId}.pdf`),
      synctexPath: null,
      logPath: join(root, 'log'),
      log: '',
      published: [],
      workspaceCleaned: true as const,
    }))
    const registry = new ProjectSessionRegistry({
      watch: () => ({ close() {} }),
      compiler,
      commitGeneration: commit as never,
      compilerRuntime: { tectonicPath: '/fixed/tectonic', userDataPath: root },
    })
    const session = await registry.attach(11, projectRoot)
    const a1 = session.compile(1, 'main.tex').catch((error) => error)
    await vi.waitFor(() => expect(runs).toHaveLength(1))
    const b = session.compile(2, 'main.tex').catch((error) => error)
    const a2 = session.compile(1, 'main.tex')
    // The middle revision is cancelled before its sidecar starts; first and newest A run.
    await vi.waitFor(() => expect(runs).toHaveLength(2))
    runs.at(-1)!.resolve({
      generationId: 'new-a',
      stagingDirectory: join(root, 'new-a'),
      files: [],
      log: '',
      workspaceCleaned: true,
    })
    await a2
    await Promise.all([a1, b])
    expect(commit).toHaveBeenCalledTimes(1)
    expect(commit.mock.calls[0]![0].generationId).toBe('new-a')
    expect(session.pdfPath(1)).toContain('new-a')
  })

  it('removes stale staged output and parses only the committed SyncTeX artifact', async () => {
    const { root, projectRoot } = await setup()
    const stale = join(root, 'stale-stage')
    const fresh = join(root, 'fresh-stage')
    await Promise.all([mkdir(stale), mkdir(fresh)])
    let release!: (value: any) => void
    const first = new Promise<any>((resolve) => (release = resolve))
    let calls = 0
    const compiler = vi.fn(async () => {
      calls += 1
      if (calls === 1) return first
      return {
        generationId: 'fresh',
        stagingDirectory: fresh,
        files: [],
        log: '',
        workspaceCleaned: true,
      }
    }) as never
    const synctex = join(root, 'fresh.synctex.gz')
    await writeFile(
      synctex,
      gzipSync(
        'Input:1:/tmp/job/input/main.tex\nUnit:1\nMagnification:1000\n{1\nx1,9:65782,131564\n}\n',
      ),
    )
    const commit = vi.fn(async () => ({
      generationId: 'fresh',
      pdfPath: null,
      synctexPath: synctex,
      synctexInputRoot: '/tmp/job/input',
      logPath: join(root, 'log'),
      log: '',
      published: [synctex],
      workspaceCleaned: true as const,
    }))
    const session = await new ProjectSessionRegistry({
      watch: () => ({ close() {} }),
      compiler,
      commitGeneration: commit as never,
      compilerRuntime: { tectonicPath: '/fixed/tectonic', userDataPath: root },
    }).attach(11, projectRoot)
    const old = session.compile(1, 'main.tex').catch((error) => error)
    await vi.waitFor(() => expect(calls).toBe(1))
    const current = session.compile(2, 'main.tex')
    release({
      generationId: 'stale',
      stagingDirectory: stale,
      files: [],
      log: '',
      workspaceCleaned: true,
    })
    await current
    await old
    await expect(access(stale)).rejects.toMatchObject({ code: 'ENOENT' })
    expect(commit).toHaveBeenCalledTimes(1)
    expect(session.syncTexForward(2, 'main.tex', 9)).toMatchObject({ page: 1 })
    expect(session.syncTexReverse(2, 1, 1, 2)).toMatchObject({ path: 'main.tex', line: 9 })
  })

  it('rejects oversized SyncTeX before reading it and evicts dead revision URLs', async () => {
    const { root, projectRoot } = await setup()
    const oversized = join(root, 'oversized.synctex.gz')
    const handle = await open(oversized, 'w')
    await handle.truncate(32 * 1024 * 1024 + 1)
    await handle.close()
    let generation = 0
    const compiler = vi.fn(async () => ({
      generationId: `g-${++generation}`,
      stagingDirectory: join(root, `stage-${generation}`),
      files: [],
      log: '',
      workspaceCleaned: true as const,
    })) as never
    const commit = vi.fn(async (staged: any) => ({
      generationId: staged.generationId,
      pdfPath: join(root, `${staged.generationId}.pdf`),
      synctexPath: oversized,
      synctexInputRoot: '/tmp/job/input',
      logPath: join(root, 'log'),
      log: '',
      published: [],
      workspaceCleaned: true as const,
    }))
    const session = await new ProjectSessionRegistry({
      watch: () => ({ close() {} }),
      compiler,
      commitGeneration: commit as never,
      maxCompileResults: 3,
      compilerRuntime: { tectonicPath: '/fixed/tectonic', userDataPath: root },
    }).attach(11, projectRoot)
    for (let revision = 1; revision <= 5; revision += 1) await session.compile(revision, 'main.tex')
    expect(session.syncTexForward(5, 'main.tex', 1)).toBeNull()
    expect(session.pdfPath(1)).toBeUndefined()
    expect(session.pdfPath(2)).toBeUndefined()
    expect(session.pdfPath(3)).toContain('g-3')
  })

  it('does not report cancellation or clear active state once publishing begins', async () => {
    const { root, projectRoot } = await setup()
    const staged = {
      generationId: 'publishing',
      stagingDirectory: join(root, 'publishing-stage'),
      files: [],
      log: '',
      workspaceCleaned: true as const,
    }
    const compiler = vi.fn(async () => staged) as never
    let release!: () => void
    const gate = new Promise<void>((resolve) => (release = resolve))
    const commit = vi.fn(async () => {
      await gate
      return {
        generationId: 'publishing',
        pdfPath: join(root, 'publishing.pdf'),
        synctexPath: null,
        synctexInputRoot: '/tmp/job/input',
        logPath: join(root, 'log'),
        log: '',
        published: [],
        workspaceCleaned: true as const,
      }
    })
    const session = await new ProjectSessionRegistry({
      watch: () => ({ close() {} }),
      compiler,
      commitGeneration: commit as never,
      compilerRuntime: { tectonicPath: '/fixed/tectonic', userDataPath: root },
    }).attach(11, projectRoot)
    const compiling = session.compile(8, 'main.tex')
    await vi.waitFor(() => expect(commit).toHaveBeenCalledOnce())
    expect(session.cancelCompile()).toBe(false)
    release()
    await expect(compiling).resolves.toMatchObject({ revision: 8 })
    expect(session.pdfPath(8)).toContain('publishing.pdf')
  })

  it('cancels the pending latest revision before it can enter the compiler', async () => {
    const { root, projectRoot } = await setup()
    let releaseA!: (value: any) => void
    const deferredA = new Promise<any>((resolve) => (releaseA = resolve))
    let compilerCalls = 0
    const compiler = vi.fn(async () => {
      compilerCalls += 1
      if (compilerCalls === 1) return deferredA
      return {
        generationId: 'should-not-run',
        stagingDirectory: join(root, 'should-not-run'),
        files: [],
        log: '',
        synctexInputRoot: '/tmp/job/input',
        workspaceCleaned: true as const,
      }
    }) as never
    const commit = vi.fn(async (staged: any) => ({
      generationId: staged.generationId,
      pdfPath: null,
      synctexPath: null,
      synctexInputRoot: staged.synctexInputRoot,
      logPath: join(root, 'log'),
      log: '',
      published: [],
      workspaceCleaned: true as const,
    }))
    const session = await new ProjectSessionRegistry({
      watch: () => ({ close() {} }),
      compiler,
      commitGeneration: commit as never,
      compilerRuntime: { tectonicPath: '/fixed/tectonic', userDataPath: root },
    }).attach(11, projectRoot)
    const a = session.compile(1, 'main.tex').catch((error) => error)
    await vi.waitFor(() => expect(compiler).toHaveBeenCalledOnce())
    const b = session.compile(2, 'main.tex').catch((error) => error as Error)
    expect(session.cancelCompile()).toBe(true)
    releaseA({
      generationId: 'old-a',
      stagingDirectory: join(root, 'old-a'),
      files: [],
      log: '',
      synctexInputRoot: '/tmp/job/input',
      workspaceCleaned: true,
    })
    await a
    const cancelled = await b
    expect(cancelled).toBeInstanceOf(Error)
    if (!(cancelled instanceof Error)) throw new Error('Expected cancellation error')
    expect(cancelled.message).toMatch(/cancel/i)
    expect(compiler).toHaveBeenCalledTimes(1)
  })

  it('always clears the active token when staging cleanup fails without replacing success', async () => {
    const { root, projectRoot } = await setup()
    let generation = 0
    const compiler = vi.fn(async () => ({
      generationId: `cleanup-${++generation}`,
      stagingDirectory: join(root, `cleanup-${generation}`),
      files: [],
      log: '',
      synctexInputRoot: '/tmp/job/input',
      workspaceCleaned: true as const,
    })) as never
    const commit = vi.fn(async (staged: any) => ({
      generationId: staged.generationId,
      pdfPath: null,
      synctexPath: null,
      synctexInputRoot: staged.synctexInputRoot,
      logPath: join(root, 'log'),
      log: '',
      published: [],
      workspaceCleaned: true as const,
    }))
    const cleanupStaging = vi.fn(async () => {
      throw new Error('injected cleanup failure')
    })
    const session = await new ProjectSessionRegistry({
      watch: () => ({ close() {} }),
      compiler,
      commitGeneration: commit as never,
      cleanupStaging,
      compilerRuntime: { tectonicPath: '/fixed/tectonic', userDataPath: root },
    }).attach(11, projectRoot)
    await expect(session.compile(4, 'main.tex')).resolves.toMatchObject({ revision: 4 })
    await expect(session.compile(4, 'main.tex')).resolves.toMatchObject({ revision: 4 })
    expect(compiler).toHaveBeenCalledTimes(2)
    expect(cleanupStaging).toHaveBeenCalledTimes(2)
  })
})
