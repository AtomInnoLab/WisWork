import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { LATEX_CHANNELS, type LatexIpcResult } from '../src/shared/ipc.js'
import { registerLatexIpc, type IpcMainLike } from '../src/main/ipc.js'
import { registerLatexPdfProtocol } from '../src/main/latex-main.js'
import {
  ProjectSessionRegistry,
  ProposalVerificationRejectedError,
  UnsavedBuffersError,
} from '../src/main/project-session.js'

describe('LaTeX typed IPC boundary', () => {
  const roots: string[] = []
  const registries: ProjectSessionRegistry[] = []

  afterEach(async () => {
    for (const registry of registries.splice(0)) registry.disposeAll()
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  async function setup() {
    const root = await mkdtemp(join(tmpdir(), 'latex-ipc-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'before')
    const registry = new ProjectSessionRegistry()
    registries.push(registry)
    const session = await registry.attach(11, projectRoot)
    const handlers = new Map<
      string,
      (event: { sender: { id: number } }, payload: unknown) => unknown
    >()
    const ipcMain: IpcMainLike = {
      handle: (channel, handler) => handlers.set(channel, handler),
      removeHandler: vi.fn(),
    }
    registerLatexIpc({ ipcMain, registry })
    const call = async <T>(channel: string, senderId: number, payload: unknown) =>
      (await handlers.get(channel)!({ sender: { id: senderId } }, payload)) as LatexIpcResult<T>
    return { root, projectRoot, session, registry, call }
  }

  it('rejects forged senders and project IDs with stable error codes', async () => {
    const { session, call } = await setup()
    await expect(
      call(LATEX_CHANNELS.projectList, 99, { projectId: session.projectId }),
    ).resolves.toMatchObject({
      ok: false,
      error: { code: 'LATEX_FORBIDDEN_SENDER' },
    })
    await expect(
      call(LATEX_CHANNELS.projectList, 11, { projectId: 'forged-project' }),
    ).resolves.toMatchObject({
      ok: false,
      error: { code: 'LATEX_PROJECT_SESSION_MISMATCH' },
    })
  })

  it('rejects roots, traversal, unknown fields, and oversized text before touching disk', async () => {
    const { root, projectRoot, session, call } = await setup()
    const outside = join(root, 'outside.tex')
    await writeFile(outside, 'outside')
    for (const payload of [
      { projectId: session.projectId, rootPath: root, path: 'main.tex', text: 'owned' },
      { projectId: session.projectId, path: '../outside.tex', text: 'owned' },
      { projectId: session.projectId, path: 'main.tex', text: 'x'.repeat(2 * 1024 * 1024 + 1) },
    ]) {
      await expect(call(LATEX_CHANNELS.fileSave, 11, payload)).resolves.toMatchObject({
        ok: false,
        error: { code: 'LATEX_INVALID_PAYLOAD' },
      })
    }
    expect(await readFile(outside, 'utf8')).toBe('outside')
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('before')
  })

  it('atomically saves the supplied text and rejects main-file rename', async () => {
    const { projectRoot, session, call } = await setup()
    await session.readText('main.tex')
    await expect(
      call(LATEX_CHANNELS.fileSave, 11, {
        projectId: session.projectId,
        path: 'main.tex',
        text: 'saved',
        editRevision: 4,
      }),
    ).resolves.toMatchObject({
      ok: true,
      value: { savedText: 'saved', buffer: { text: 'saved', dirty: false } },
    })
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('saved')
    await expect(
      call(LATEX_CHANNELS.fileRename, 11, {
        projectId: session.projectId,
        from: 'main.tex',
        to: 'other.tex',
      }),
    ).resolves.toMatchObject({ ok: false, error: { code: 'LATEX_INVALID_PAYLOAD' } })
  })

  it('does not accept executable, bundle URL, root, headers, env, or arguments from compile payloads', async () => {
    const { session, call } = await setup()
    for (const field of ['executable', 'bundleUrl', 'rootPath', 'headers', 'env', 'args']) {
      await expect(
        call(LATEX_CHANNELS.compileStart, 11, {
          projectId: session.projectId,
          revision: 1,
          mainFile: 'main.tex',
          [field]: field === 'args' ? ['--shell-escape'] : 'attacker-controlled',
        }),
      ).resolves.toMatchObject({ ok: false, error: { code: 'LATEX_INVALID_PAYLOAD' } })
    }
  })

  it('accepts only projectId and proposalId for isolated proposal verification', async () => {
    const { session, call } = await setup()
    const verifyProposal = vi.fn().mockResolvedValue({
      proposalId: 'proposal-1',
      state: 'verified',
      diagnostics: [],
      logSummary: '',
      verifiedAt: 1,
    })
    ;(session as unknown as { verifyProposal: typeof verifyProposal }).verifyProposal =
      verifyProposal

    await expect(
      call(LATEX_CHANNELS.proposalVerify, 11, {
        projectId: session.projectId,
        proposalId: 'proposal-1',
      }),
    ).resolves.toMatchObject({ ok: true, value: { state: 'verified' } })
    expect(verifyProposal).toHaveBeenCalledWith('proposal-1')

    for (const [field, value] of [
      ['overlay', [{ path: 'main.tex', text: 'attacker' }]],
      ['args', ['--shell-escape']],
      ['mainFile', '../outside.tex'],
      ['projectDirectory', '/tmp/attacker-project'],
    ] as const) {
      await expect(
        call(LATEX_CHANNELS.proposalVerify, 11, {
          projectId: session.projectId,
          proposalId: 'proposal-1',
          [field]: value,
        }),
      ).resolves.toMatchObject({ ok: false, error: { code: 'LATEX_INVALID_PAYLOAD' } })
    }
    expect(verifyProposal).toHaveBeenCalledTimes(1)
  })

  it('exposes bundle status only through the owning project session', async () => {
    const { session, call } = await setup()
    const payload = { projectId: session.projectId }
    await expect(call(LATEX_CHANNELS.bundleStatus, 11, payload)).resolves.toEqual({
      ok: true,
      value: { state: 'error', code: 'BUNDLE_NOT_CONFIGURED' },
    })
    await expect(call(LATEX_CHANNELS.bundleStatus, 99, payload)).resolves.toMatchObject({
      ok: false,
      error: { code: 'LATEX_FORBIDDEN_SENDER' },
    })
    await expect(
      call(LATEX_CHANNELS.bundleStatus, 11, {
        ...payload,
        bundleUrl: 'https://attacker.invalid/bundle.tar',
      }),
    ).resolves.toMatchObject({
      ok: false,
      error: { code: 'LATEX_INVALID_PAYLOAD' },
    })
  })

  it('maps unsaved-buffer compile rejection to a stable conflict error', async () => {
    const { session, call } = await setup()
    vi.spyOn(session, 'compile').mockRejectedValue(new UnsavedBuffersError())
    await expect(
      call(LATEX_CHANNELS.compileStart, 11, {
        projectId: session.projectId,
        revision: 1,
        mainFile: 'main.tex',
      }),
    ).resolves.toEqual({
      ok: false,
      error: { code: 'LATEX_CONFLICT', message: 'Project has unsaved LaTeX changes' },
    })
  })

  it('maps proposal verification rejections to stable public error codes', async () => {
    const { session, call } = await setup()
    const request = { projectId: session.projectId, proposalId: 'proposal-1' }
    const verify = vi.spyOn(session, 'verifyProposal')

    verify.mockRejectedValueOnce(
      new ProposalVerificationRejectedError('expired', 'Proposal has expired'),
    )
    await expect(call(LATEX_CHANNELS.proposalVerify, 11, request)).resolves.toEqual({
      ok: false,
      error: { code: 'LATEX_NOT_FOUND', message: 'Proposal has expired' },
    })

    verify.mockRejectedValueOnce(
      new ProposalVerificationRejectedError('baseline', 'Proposal baseline changed on disk'),
    )
    await expect(call(LATEX_CHANNELS.proposalVerify, 11, request)).resolves.toEqual({
      ok: false,
      error: { code: 'LATEX_CONFLICT', message: 'Proposal baseline changed on disk' },
    })

    verify.mockRejectedValueOnce(
      new ProposalVerificationRejectedError(
        'safety',
        'Proposal verification was rejected by the compiler safety policy',
      ),
    )
    await expect(call(LATEX_CHANNELS.proposalVerify, 11, request)).resolves.toEqual({
      ok: false,
      error: {
        code: 'LATEX_VERIFICATION_REJECTED',
        message: 'Proposal verification was rejected by the compiler safety policy',
      },
    })
  })

  it('serves only session-owned compiled PDFs through the controlled protocol', async () => {
    const { root } = await setup()
    const pdf = join(root, 'controlled.pdf')
    await writeFile(pdf, '%PDF-controlled')
    let handler: ((request: { url: string }) => Promise<Response>) | undefined
    registerLatexPdfProtocol({ handle: (_scheme, next) => (handler = next) }, {
      resolvePdf: (projectId: string, revision: number) =>
        projectId === 'a'.repeat(32) && revision === 7 ? pdf : undefined,
    } as never)
    const allowed = await handler!({ url: `wiswork-latex-pdf://${'a'.repeat(32)}/7` })
    expect(allowed.status).toBe(200)
    expect(allowed.headers.get('access-control-allow-origin')).toBe('*')
    expect(await allowed.text()).toBe('%PDF-controlled')
    await expect(
      handler!({ url: `wiswork-latex-pdf://${'a'.repeat(32)}/../../outside.txt` }),
    ).resolves.toMatchObject({ status: 400 })
    await expect(handler!({ url: 'file:///etc/passwd' })).resolves.toMatchObject({ status: 400 })
  })

  it('returns a fixed internal error without leaking absolute paths', async () => {
    const { session, call } = await setup()
    vi.spyOn(session, 'listTextFiles').mockRejectedValue(
      new Error('failed at /private/secret/project'),
    )
    await expect(
      call(LATEX_CHANNELS.projectList, 11, { projectId: session.projectId }),
    ).resolves.toEqual({
      ok: false,
      error: { code: 'LATEX_INTERNAL', message: 'LaTeX operation failed' },
    })
  })

  it('validates and owns SyncTeX forward and reverse requests', async () => {
    const { session, call } = await setup()
    vi.spyOn(session, 'syncTexForward').mockReturnValue({ page: 1, x: 2, y: 3 })
    vi.spyOn(session, 'syncTexReverse').mockReturnValue({ path: 'main.tex', line: 9 })
    await expect(
      call(LATEX_CHANNELS.syncTexForward, 11, {
        projectId: session.projectId,
        revision: 1,
        path: 'main.tex',
        line: 9,
      }),
    ).resolves.toMatchObject({ ok: true, value: { page: 1 } })
    await expect(
      call(LATEX_CHANNELS.syncTexReverse, 11, {
        projectId: session.projectId,
        revision: 1,
        page: 1,
        x: 2,
        y: 3,
      }),
    ).resolves.toMatchObject({ ok: true, value: { path: 'main.tex' } })
    await expect(
      call(LATEX_CHANNELS.syncTexForward, 99, {
        projectId: session.projectId,
        revision: 1,
        path: 'main.tex',
        line: 9,
      }),
    ).resolves.toMatchObject({ ok: false, error: { code: 'LATEX_FORBIDDEN_SENDER' } })
  })
})
