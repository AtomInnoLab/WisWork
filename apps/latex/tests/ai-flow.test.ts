import { createHash } from 'node:crypto'
import { access, mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { compileIsolated, TectonicRunError } from '@wiswork/latex-compiler'
import { ProjectSessionRegistry } from '../src/main/project-session.js'

const sha = (text: string) => createHash('sha256').update(text).digest('hex')
const testFreeze = async () => () => undefined

function deferred<T>() {
  let resolve!: (value: T) => void
  const promise = new Promise<T>((done) => (resolve = done))
  return { promise, resolve }
}

describe('confirmed LaTeX AI edit flow', () => {
  const roots: string[] = []

  afterEach(async () => {
    vi.restoreAllMocks()
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  async function setup() {
    const root = await mkdtemp(join(tmpdir(), 'latex-ai-flow-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'before')
    const session = await new ProjectSessionRegistry({
      watch: () => ({ close() {} }),
      acquireRendererFreeze: testFreeze,
      compilerRuntime: { tectonicPath: '/fixed/tectonic', userDataPath: root },
    }).attach(41, projectRoot)
    await session.readText('main.tex')
    await session.registerProposal({
      id: 'proposal-1',
      expiresAt: Date.now() + 60_000,
      files: [{ path: 'main.tex', beforeSha256: sha('before'), afterText: 'after' }],
    })
    return { projectRoot, session }
  }

  async function setupVerification(
    files: Array<{ path: string; beforeSha256: string | null; afterText: string }> = [
      { path: 'main.tex', beforeSha256: sha('before'), afterText: 'after' },
    ],
    compilerFailure?: Error,
  ) {
    const root = await mkdtemp(join(tmpdir(), 'latex-proposal-verify-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    const stagingDirectory = join(root, 'verify-stage')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'before')
    const compiler = vi.fn(async () => {
      if (compilerFailure) throw compilerFailure
      await mkdir(stagingDirectory, { recursive: true })
      await writeFile(join(stagingDirectory, 'main.log'), 'staged evidence')
      return {
        generationId: 'verification-generation',
        stagingDirectory,
        files: [],
        log: 'main.tex:2:3: warning: isolated warning',
        synctexInputRoot: join(root, 'isolated-input'),
        workspaceCleaned: true as const,
      }
    })
    const commitGeneration = vi.fn()
    const cleanupStaging = vi.fn((path: string) => rm(path, { recursive: true, force: true }))
    const session = await new ProjectSessionRegistry({
      watch: () => ({ close() {} }),
      compiler: compiler as never,
      commitGeneration: commitGeneration as never,
      cleanupStaging,
      acquireRendererFreeze: testFreeze,
      compilerRuntime: { tectonicPath: '/fixed/tectonic', userDataPath: root },
    }).attach(51, projectRoot)
    await session.readText('main.tex')
    await session.registerProposal({
      id: 'proposal-verify',
      expiresAt: Date.now() + 60_000,
      files,
    })
    return {
      root,
      projectRoot,
      stagingDirectory,
      session,
      compiler,
      commitGeneration,
      cleanupStaging,
    }
  }

  it('verifies a replacement in isolation without changing source, publishing, or consuming it', async () => {
    const { projectRoot, stagingDirectory, session, compiler, commitGeneration, cleanupStaging } =
      await setupVerification()
    const sourceBefore = sha(await readFile(join(projectRoot, 'main.tex'), 'utf8'))

    await expect(session.verifyProposal('proposal-verify')).resolves.toMatchObject({
      proposalId: 'proposal-verify',
      state: 'verified',
      diagnostics: [{ path: 'main.tex', severity: 'warning' }],
    })
    expect(compiler).toHaveBeenCalledWith(
      expect.objectContaining({
        projectDirectory: projectRoot,
        mainFile: 'main.tex',
        overlay: [{ path: 'main.tex', text: 'after' }],
        expectedSourceHashes: { 'main.tex': sha('before') },
      }),
    )
    expect(commitGeneration).not.toHaveBeenCalled()
    expect(cleanupStaging).toHaveBeenCalledWith(stagingDirectory)
    await expect(access(stagingDirectory)).rejects.toThrow()
    expect(sha(await readFile(join(projectRoot, 'main.tex'), 'utf8'))).toBe(sourceBefore)

    vi.spyOn(session, 'compile').mockResolvedValue({
      revision: 1,
      pdfUrl: null,
      diagnostics: [],
      log: '',
    })
    await expect(session.applyConfirmedProposal('proposal-verify')).resolves.toMatchObject({
      proposalId: 'proposal-verify',
    })
    await expect(session.applyConfirmedProposal('proposal-verify')).rejects.toThrow(/not found/i)
  })

  it('returns typed unverifiable evidence for a new file without calling the compiler', async () => {
    const { session, compiler } = await setupVerification([
      { path: 'new.tex', beforeSha256: null, afterText: 'new file' },
    ])
    await expect(session.verifyProposal('proposal-verify')).resolves.toMatchObject({
      proposalId: 'proposal-verify',
      state: 'unverifiable',
      diagnostics: [],
      reason: expect.stringMatching(/new file/i),
    })
    expect(compiler).not.toHaveBeenCalled()
  })

  it('checks every existing baseline before returning mixed proposals as unverifiable', async () => {
    const { projectRoot, session, compiler } = await setupVerification([
      { path: 'main.tex', beforeSha256: sha('before'), afterText: 'after' },
      { path: 'new.tex', beforeSha256: null, afterText: 'new file' },
    ])
    await writeFile(join(projectRoot, 'main.tex'), 'stale')
    await expect(session.verifyProposal('proposal-verify')).rejects.toThrow(/baseline|changed/i)
    expect(compiler).not.toHaveBeenCalled()
  })

  it('waits for the renderer freeze before checking persisted buffers and always releases it', async () => {
    const freeze = deferred<() => void>()
    const release = vi.fn()
    const root = await mkdtemp(join(tmpdir(), 'latex-freeze-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'before')
    const compiler = vi.fn(async () => ({
      generationId: 'verified',
      stagingDirectory: join(root, 'stage'),
      files: [],
      log: '',
      synctexInputRoot: projectRoot,
      workspaceCleaned: true as const,
    }))
    const session = await new ProjectSessionRegistry({
      watch: () => ({ close() {} }),
      compiler: compiler as never,
      compilerRuntime: { tectonicPath: '/fixed/tectonic', userDataPath: root },
      acquireRendererFreeze: () => freeze.promise,
    }).attach(91, projectRoot)
    await session.readText('main.tex')
    session.updateBuffer('main.tex', 'pending')
    await session.registerProposal({
      id: 'fenced',
      expiresAt: Date.now() + 60_000,
      files: [{ path: 'main.tex', beforeSha256: sha('before'), afterText: 'after' }],
    })
    const verifying = session.verifyProposal('fenced')
    await Promise.resolve()
    expect(compiler).not.toHaveBeenCalled()
    await session.saveText('main.tex', 'pending')
    freeze.resolve(release)
    await expect(verifying).rejects.toThrow(/baseline|changed/i)
    expect(release).toHaveBeenCalledOnce()
  })

  it('rejects when disposed while waiting for the renderer freeze', async () => {
    const freeze = deferred<() => void>()
    const { session } = await setupVerification()
    ;(
      session as unknown as { acquireRendererFreeze: () => Promise<() => void> }
    ).acquireRendererFreeze = () => freeze.promise
    const release = vi.fn()
    const verifying = session.verifyProposal('proposal-verify')
    session.dispose()
    freeze.resolve(release)
    await expect(verifying).rejects.toThrow(/closed/i)
    expect(release).toHaveBeenCalledOnce()
  })

  it('refuses a confirmed transaction when the renderer cannot be frozen', async () => {
    const root = await mkdtemp(join(tmpdir(), 'latex-freeze-fail-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'before')
    const compiler = vi.fn()
    const session = await new ProjectSessionRegistry({
      watch: () => ({ close() {} }),
      compiler: compiler as never,
      compilerRuntime: { tectonicPath: '/fixed/tectonic', userDataPath: root },
      acquireRendererFreeze: async () => {
        throw new Error('freeze failed')
      },
    }).attach(92, projectRoot)
    await session.registerProposal({
      id: 'fenced',
      expiresAt: Date.now() + 60_000,
      files: [{ path: 'main.tex', beforeSha256: sha('before'), afterText: 'after' }],
    })
    await expect(session.verifyProposal('fenced')).rejects.toThrow(/freeze/i)
    expect(compiler).not.toHaveBeenCalled()
  })

  it('rejects dirty buffers and changed proposal baselines before compiling', async () => {
    const dirty = await setupVerification()
    dirty.session.updateBuffer('main.tex', 'local dirty')
    await expect(dirty.session.verifyProposal('proposal-verify')).rejects.toThrow(/unsaved/i)
    expect(dirty.compiler).not.toHaveBeenCalled()

    const changed = await setupVerification()
    await writeFile(join(changed.projectRoot, 'main.tex'), 'external change')
    await expect(changed.session.verifyProposal('proposal-verify')).rejects.toThrow(
      /baseline|changed/i,
    )
    expect(changed.compiler).not.toHaveBeenCalled()
  })

  it('rejects a baseline changed between session precheck and compiler copy before run', async () => {
    const root = await mkdtemp(join(tmpdir(), 'latex-proposal-toctou-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'before')
    const run = vi.fn()
    let changed = false
    const compiler = vi.fn((request: Parameters<typeof compileIsolated>[0]) =>
      compileIsolated({
        ...request,
        hooks: {
          afterDirectoryRead: async (path) => {
            if (path !== projectRoot || changed) return
            changed = true
            await writeFile(join(projectRoot, 'main.tex'), 'changed')
          },
        },
        run,
      }),
    )
    const session = await new ProjectSessionRegistry({
      watch: () => ({ close() {} }),
      compiler,
      acquireRendererFreeze: testFreeze,
      compilerRuntime: { tectonicPath: '/fixed/tectonic', userDataPath: root },
    }).attach(52, projectRoot)
    await session.registerProposal({
      id: 'toctou-proposal',
      expiresAt: Date.now() + 60_000,
      files: [{ path: 'main.tex', beforeSha256: sha('before'), afterText: 'after' }],
    })

    await expect(session.verifyProposal('toctou-proposal')).rejects.toThrow(/safety|rejected/i)
    expect(run).not.toHaveBeenCalled()
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('changed')
  })

  it('returns bounded failed evidence for a Tectonic run error', async () => {
    const log = `${'😀'.repeat(5_000)}\n${Array.from(
      { length: 150 },
      (_, index) => `main.tex:${index + 1}:1: error: ${'x'.repeat(200)}`,
    ).join('\n')}`
    const failure = new TectonicRunError('TECTONIC_EXIT_NONZERO', 'compile failed', log, 1)
    failure.terminationConfirmed = true
    const { session, compiler, cleanupStaging } = await setupVerification(undefined, failure)

    const result = await session.verifyProposal('proposal-verify')
    expect(result).toMatchObject({ proposalId: 'proposal-verify', state: 'failed' })
    expect(result.diagnostics).toHaveLength(100)
    expect(Buffer.byteLength(result.logSummary, 'utf8')).toBeLessThanOrEqual(16_000)
    expect(compiler).toHaveBeenCalledOnce()
    expect(cleanupStaging).not.toHaveBeenCalled()
  })

  it.each([
    ['unconfirmed exit', 'TECTONIC_EXIT_NONZERO', false, 1],
    ['spawn failure', 'TECTONIC_EXIT_NONZERO', true, null],
    ['total timeout', 'TECTONIC_TOTAL_TIMEOUT', true, null],
    ['idle timeout', 'TECTONIC_IDLE_TIMEOUT', true, null],
    ['output limit', 'TECTONIC_OUTPUT_LIMIT', true, null],
    ['cancelled', 'TECTONIC_CANCELLED', true, null],
  ] as const)(
    'rejects unsafe Tectonic failure classification: %s',
    async (_label, code, terminationConfirmed, exitCode) => {
      const failure = new TectonicRunError(code, 'unsafe failure', 'unsafe log', exitCode)
      failure.terminationConfirmed = terminationConfirmed
      const { session } = await setupVerification(undefined, failure)
      await expect(session.verifyProposal('proposal-verify')).rejects.toThrow(/rejected|safety/i)
    },
  )

  it('serializes ordinary compile and proposal verification through one queue', async () => {
    const root = await mkdtemp(join(tmpdir(), 'latex-proposal-queue-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'before')
    let calls = 0
    let active = 0
    let maxActive = 0
    let markFirstStarted!: () => void
    let releaseFirst!: () => void
    const firstStarted = new Promise<void>((resolve) => (markFirstStarted = resolve))
    const firstGate = new Promise<void>((resolve) => (releaseFirst = resolve))
    const compiler = vi.fn(async (request: { signal?: AbortSignal }) => {
      calls += 1
      active += 1
      maxActive = Math.max(maxActive, active)
      try {
        if (calls === 1) {
          markFirstStarted()
          await Promise.race([
            firstGate.then(() => {
              throw new Error('test release')
            }),
            new Promise<never>((_resolve, reject) =>
              request.signal?.addEventListener('abort', () => reject(new Error('aborted')), {
                once: true,
              }),
            ),
          ])
        }
        return {
          generationId: 'verified',
          stagingDirectory: join(root, 'verified-stage'),
          files: [],
          log: '',
          synctexInputRoot: '/isolated/input',
          workspaceCleaned: true as const,
        }
      } finally {
        active -= 1
      }
    })
    const session = await new ProjectSessionRegistry({
      watch: () => ({ close() {} }),
      compiler: compiler as never,
      commitGeneration: vi.fn() as never,
      acquireRendererFreeze: testFreeze,
      compilerRuntime: { tectonicPath: '/fixed/tectonic', userDataPath: root },
    }).attach(61, projectRoot)
    await session.registerProposal({
      id: 'queued-proposal',
      expiresAt: Date.now() + 60_000,
      files: [{ path: 'main.tex', beforeSha256: sha('before'), afterText: 'after' }],
    })

    const compiling = session.compile(1, 'main.tex').catch((error) => error)
    await firstStarted
    const verifying = session.verifyProposal('queued-proposal')
    await vi.waitFor(() => expect(compiler).toHaveBeenCalledTimes(2))
    releaseFirst()
    await expect(verifying).resolves.toMatchObject({ state: 'verified' })
    await compiling
    expect(maxActive).toBe(1)
  })

  it('dispose aborts and rejects an active proposal verification', async () => {
    const root = await mkdtemp(join(tmpdir(), 'latex-proposal-cancel-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'before')
    let markStarted!: () => void
    let release!: () => void
    let observedSignal: AbortSignal | undefined
    const started = new Promise<void>((resolve) => (markStarted = resolve))
    const gate = new Promise<void>((resolve) => (release = resolve))
    const compiler = vi.fn((request: { signal?: AbortSignal }) => {
      observedSignal = request.signal
      markStarted()
      return new Promise((resolve, reject) => {
        request.signal?.addEventListener('abort', () => reject(new Error('aborted')), {
          once: true,
        })
        void gate.then(() => reject(new Error('test release')))
      })
    })
    const session = await new ProjectSessionRegistry({
      watch: () => ({ close() {} }),
      compiler: compiler as never,
      acquireRendererFreeze: testFreeze,
      compilerRuntime: { tectonicPath: '/fixed/tectonic', userDataPath: root },
    }).attach(71, projectRoot)
    await session.registerProposal({
      id: 'cancel-proposal',
      expiresAt: Date.now() + 60_000,
      files: [{ path: 'main.tex', beforeSha256: sha('before'), afterText: 'after' }],
    })

    const verifying = session.verifyProposal('cancel-proposal')
    await started
    expect(session.cancelCompile()).toBe(false)
    await expect(session.compile(99, 'main.tex')).rejects.toThrow(/transaction/i)
    session.dispose()
    const cancelled = true
    const aborted = observedSignal?.aborted === true
    if (!aborted) release()
    await expect(verifying).rejects.toThrow()
    expect(cancelled).toBe(true)
    expect(aborted).toBe(true)
  })

  it('rejects expired proposals before compiling', async () => {
    const { session, compiler } = await setupVerification()
    vi.spyOn(Date, 'now').mockReturnValue(Number.MAX_SAFE_INTEGER)
    await expect(session.verifyProposal('proposal-verify')).rejects.toThrow(/expired/i)
    expect(compiler).not.toHaveBeenCalled()
  })

  it('rejects renderer-dirty state before consuming authorization', async () => {
    const { session } = await setup()
    session.updateBuffer('main.tex', 'local dirty')
    const compile = vi.spyOn(session, 'compile')
    await expect(session.applyConfirmedProposal('proposal-1')).rejects.toThrow(/unsaved|conflict/i)
    expect(compile).not.toHaveBeenCalled()
    await session.discardAll()
    compile.mockResolvedValue({ revision: 1, pdfUrl: null, diagnostics: [], log: '' })
    await expect(session.applyConfirmedProposal('proposal-1')).resolves.toMatchObject({
      proposalId: 'proposal-1',
      compile: { ok: true },
    })
  })

  it('applies once, compiles once, then undoes the whole batch and compiles once', async () => {
    const { projectRoot, session } = await setup()
    const compile = vi
      .spyOn(session, 'compile')
      .mockResolvedValue({ revision: 1, pdfUrl: null, diagnostics: [], log: '' })
    const applied = await session.applyConfirmedProposal('proposal-1')
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('after')
    expect(compile).toHaveBeenCalledTimes(1)
    await expect(session.applyConfirmedProposal('proposal-1')).rejects.toThrow()
    expect(compile).toHaveBeenCalledTimes(1)
    await expect(session.undoConfirmedProposal(applied.snapshotId)).resolves.toMatchObject({
      restored: true,
      compile: { ok: true },
    })
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('before')
    expect(compile).toHaveBeenCalledTimes(2)
  })

  it('keeps the confirmed edit when automatic compile fails without another agent turn', async () => {
    const { projectRoot, session } = await setup()
    vi.spyOn(session, 'compile').mockRejectedValue(new Error('compile error'))
    await expect(session.applyConfirmedProposal('proposal-1')).resolves.toMatchObject({
      compile: { ok: false, error: 'compile error' },
    })
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('after')
  })

  it('rejects external compile/cancel during apply while allowing its formal internal compile', async () => {
    const { projectRoot, session } = await setup()
    let started!: () => void
    let release!: () => void
    const startedPromise = new Promise<void>((resolve) => (started = resolve))
    const gate = new Promise<void>((resolve) => (release = resolve))
    const staged = {
      generationId: 'formal',
      stagingDirectory: join(projectRoot, '..', 'formal-stage'),
      files: [],
      log: '',
      workspaceCleaned: true as const,
    }
    ;(session as unknown as { compiler: () => Promise<typeof staged> }).compiler = async () => {
      started()
      await gate
      return staged
    }
    ;(session as unknown as { commitGeneration: () => Promise<unknown> }).commitGeneration =
      async () => ({
        ...staged,
        pdfPath: null,
        synctexPath: null,
        synctexInputRoot: projectRoot,
        logPath: join(projectRoot, '..', 'formal.log'),
        published: [],
      })
    const applying = session.applyConfirmedProposal('proposal-1')
    await startedPromise
    await expect(session.compile(999, 'main.tex')).rejects.toThrow(/transaction/i)
    expect(session.cancelCompile()).toBe(false)
    release()
    await expect(applying).resolves.toMatchObject({ compile: { ok: true } })
  })

  it('bounds main-process AI reads and rejects binary targets', async () => {
    const { projectRoot, session } = await setup()
    await writeFile(join(projectRoot, 'large.tex'), 'x'.repeat(256 * 1024 + 1))
    await writeFile(join(projectRoot, 'binary.tex'), Buffer.from([0xff, 0xfe, 0x00]))
    await expect(session.readProjectTextForAi('large.tex', 0, 100)).rejects.toThrow(/size/i)
    await expect(session.readProjectTextForAi('binary.tex', 0, 100)).rejects.toThrow(
      /UTF-8|binary/i,
    )
    await expect(session.readProjectTextForAi('../outside.tex', 0, 100)).rejects.toThrow()
    await expect(session.searchProjectTextForAi('x', 51)).rejects.toThrow(/limit/i)
  })

  it('normalizes content and owns proposal id, TTL, and baseline hashes in main', async () => {
    const { projectRoot, session } = await setup()
    await mkdir(join(projectRoot, 'chapters'))
    const before = Date.now()
    const first = await session.createEditProposal([
      { path: './chapters\\new.tex', afterText: 'complete file' },
    ])
    const second = await session.createEditProposal([
      { path: 'main.tex', afterText: 'replacement' },
    ])
    expect(first.id).not.toBe(second.id)
    expect(first.expiresAt).toBeGreaterThan(before)
    expect(first.files[0]).toMatchObject({
      path: 'chapters/new.tex',
      beforeText: null,
      beforeSha256: null,
      afterText: 'complete file',
    })
    expect(second.files[0].beforeSha256).toBe(sha('before'))
  })

  it('keeps ordinary editor reads available for sensitive-looking LaTeX files while AI access refuses them', async () => {
    const { projectRoot, session } = await setup()
    await writeFile(join(projectRoot, 'secret.tex'), 'editor secret')
    await writeFile(join(projectRoot, 'credentials.tex'), 'editor credentials')

    await expect(session.readText('secret.tex')).resolves.toMatchObject({ text: 'editor secret' })
    await expect(session.readText('credentials.tex')).resolves.toMatchObject({
      text: 'editor credentials',
    })
    await expect(session.listProjectFilesForAi()).resolves.not.toEqual(
      expect.objectContaining({ files: expect.arrayContaining(['secret.tex', 'credentials.tex']) }),
    )
    await expect(session.readProjectTextForAi('secret.tex', 0, 100)).rejects.toThrow(/sensitive/i)
  })

  it('rejects renderer file mutations for the whole confirmed apply and undo transaction', async () => {
    const { session } = await setup()
    await session.createText('notes.tex', 'notes')
    let releaseApply!: () => void
    let markApplyStarted!: () => void
    const applyGate = new Promise<void>((resolve) => (releaseApply = resolve))
    const applyStarted = new Promise<void>((resolve) => (markApplyStarted = resolve))
    vi.spyOn(session, 'compile').mockImplementationOnce(async () => {
      markApplyStarted()
      await applyGate
      return { revision: 1, pdfUrl: null, diagnostics: [], log: '' }
    })

    const applying = session.applyConfirmedProposal('proposal-1')
    await applyStarted
    expect(() => session.updateBuffer('main.tex', 'racing edit')).toThrow(/transaction/i)
    await expect(session.saveText('main.tex', 'racing save')).rejects.toThrow(/transaction/i)
    await expect(session.createText('new.tex', 'racing create')).rejects.toThrow(/transaction/i)
    await expect(session.renameText('notes.tex', 'renamed.tex')).rejects.toThrow(/transaction/i)
    releaseApply()
    const applied = await applying

    let releaseUndo!: () => void
    let markUndoStarted!: () => void
    const undoGate = new Promise<void>((resolve) => (releaseUndo = resolve))
    const undoStarted = new Promise<void>((resolve) => (markUndoStarted = resolve))
    vi.spyOn(session, 'compile').mockImplementationOnce(async () => {
      markUndoStarted()
      await undoGate
      return { revision: 2, pdfUrl: null, diagnostics: [], log: '' }
    })
    const undoing = session.undoConfirmedProposal(applied.snapshotId)
    await undoStarted
    expect(() => session.updateBuffer('main.tex', 'racing undo edit')).toThrow(/transaction/i)
    await expect(session.saveText('main.tex', 'racing undo save')).rejects.toThrow(/transaction/i)
    releaseUndo()
    await undoing
    expect(() => session.updateBuffer('main.tex', 'allowed afterwards')).not.toThrow()
  })

  it('waits for an already-running create before applying and rejects new mutations meanwhile', async () => {
    const { session } = await setup()
    const originalSave = session.project.saveText.bind(session.project)
    let releaseCreate!: () => void
    let markCreateStarted!: () => void
    const createGate = new Promise<void>((resolve) => (releaseCreate = resolve))
    const createStarted = new Promise<void>((resolve) => (markCreateStarted = resolve))
    let createInFlight = false
    let applyOverlappedCreate = false
    vi.spyOn(session.project, 'saveText').mockImplementation(async (path, text, options) => {
      if (path === 'slow.tex') {
        createInFlight = true
        markCreateStarted()
        await createGate
        createInFlight = false
      }
      if (path === 'main.tex' && createInFlight) applyOverlappedCreate = true
      return originalSave(path, text, options)
    })
    vi.spyOn(session, 'compile').mockResolvedValue({
      revision: 1,
      pdfUrl: null,
      diagnostics: [],
      log: '',
    })

    const creating = session.createText('slow.tex', 'slow')
    await createStarted
    let applyFinished = false
    const applying = session.applyConfirmedProposal('proposal-1').then((value) => {
      applyFinished = true
      return value
    })
    await Promise.resolve()
    expect(applyFinished).toBe(false)
    expect(applyOverlappedCreate).toBe(false)
    expect(() => session.updateBuffer('main.tex', 'new mutation')).toThrow(/transaction/i)
    releaseCreate()
    await creating
    await applying
    expect(applyFinished).toBe(true)
  })

  it('waits for already-running save and rename before undoing without overlap', async () => {
    const { session } = await setup()
    await session.createText('notes.tex', 'notes')
    await session.createText('rename-me.tex', 'rename')
    vi.spyOn(session, 'compile').mockResolvedValue({
      revision: 1,
      pdfUrl: null,
      diagnostics: [],
      log: '',
    })
    const applied = await session.applyConfirmedProposal('proposal-1')
    session.updateBuffer('notes.tex', 'saved notes')

    const originalSave = session.project.saveText.bind(session.project)
    const originalRead = session.project.readText.bind(session.project)
    let releaseSave!: () => void
    let markSaveStarted!: () => void
    let releaseRename!: () => void
    let markRenameStarted!: () => void
    const saveGate = new Promise<void>((resolve) => (releaseSave = resolve))
    const saveStarted = new Promise<void>((resolve) => (markSaveStarted = resolve))
    const renameGate = new Promise<void>((resolve) => (releaseRename = resolve))
    const renameStarted = new Promise<void>((resolve) => (markRenameStarted = resolve))
    vi.spyOn(session.project, 'saveText').mockImplementation(async (path, text, options) => {
      if (path === 'notes.tex') {
        markSaveStarted()
        await saveGate
      }
      return originalSave(path, text, options)
    })
    vi.spyOn(session.project, 'readText').mockImplementation(async (path) => {
      if (path === 'rename-me.tex') {
        markRenameStarted()
        await renameGate
      }
      return originalRead(path)
    })

    const saving = session.saveText('notes.tex')
    const renaming = session.renameText('rename-me.tex', 'renamed.tex')
    await Promise.all([saveStarted, renameStarted])
    let undoFinished = false
    const undoing = session.undoConfirmedProposal(applied.snapshotId).then(
      (value) => {
        undoFinished = true
        return { ok: true as const, value }
      },
      (error: unknown) => ({ ok: false as const, error }),
    )
    await Promise.resolve()
    expect(undoFinished).toBe(false)
    releaseSave()
    await saving
    await Promise.resolve()
    expect(undoFinished).toBe(false)
    releaseRename()
    await renaming
    const undoResult = await undoing
    expect(undoResult).toMatchObject({ ok: true })
    expect(undoFinished).toBe(true)
  })

  it('bounds both each baseline and aggregate before/after proposal review data', async () => {
    const { projectRoot, session } = await setup()
    await writeFile(join(projectRoot, 'large.tex'), 'x'.repeat(256 * 1024 + 1))
    await expect(
      session.createEditProposal([{ path: 'large.tex', afterText: 'replacement' }]),
    ).rejects.toThrow(/baseline|size/i)

    const files = Array.from({ length: 17 }, (_, index) => ({
      path: `part-${index}.tex`,
      afterText: 'replacement',
    }))
    await Promise.all(
      files.map((file) => writeFile(join(projectRoot, file.path), 'x'.repeat(250 * 1024))),
    )
    await expect(session.createEditProposal(files)).rejects.toThrow(/total|size/i)
  })
})
