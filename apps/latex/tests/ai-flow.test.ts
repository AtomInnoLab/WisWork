import { createHash } from 'node:crypto'
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { ProjectSessionRegistry } from '../src/main/project-session.js'

const sha = (text: string) => createHash('sha256').update(text).digest('hex')

describe('confirmed LaTeX AI edit flow', () => {
  const roots: string[] = []

  afterEach(async () => {
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
