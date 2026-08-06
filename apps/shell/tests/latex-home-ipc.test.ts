import { describe, expect, it, vi } from 'vitest'
import { registerLatexHomeIpc } from '../src/main/latex-home-ipc'
import { HOME_CHANNELS } from '../src/shared/home-api'

describe('LaTeX Home IPC ownership', () => {
  it('rejects non-Shell senders and wrong argument counts before side effects', async () => {
    const handlers = new Map<string, (...args: unknown[]) => unknown>()
    const shellSender = {}
    const recents = vi.fn()
    const create = vi.fn()
    const importProject = vi.fn()
    const open = vi.fn()
    registerLatexHomeIpc({
      ipcMain: { handle: (channel, handler) => handlers.set(channel, handler) },
      shellSender: () => shellSender,
      recents,
      create,
      importProject,
      open,
    })
    for (const channel of [
      HOME_CHANNELS.latexRecents,
      HOME_CHANNELS.newLatexProject,
      HOME_CHANNELS.importLatexProject,
    ]) {
      await expect(handlers.get(channel)!({ sender: {} })).rejects.toThrow(/Untrusted/)
      await expect(handlers.get(channel)!({ sender: shellSender }, 'extra')).rejects.toThrow(
        /payload/,
      )
    }
    await expect(
      handlers.get(HOME_CHANNELS.openLatexProject)!({ sender: {} }, '/valid'),
    ).rejects.toThrow(/Untrusted/)
    await expect(
      handlers.get(HOME_CHANNELS.openLatexProject)!({ sender: shellSender }),
    ).rejects.toThrow(/payload/)
    await expect(
      handlers.get(HOME_CHANNELS.openLatexProject)!({ sender: shellSender }, '/one', '/two'),
    ).rejects.toThrow(/payload/)
    expect(recents).not.toHaveBeenCalled()
    expect(create).not.toHaveBeenCalled()
    expect(importProject).not.toHaveBeenCalled()
    expect(open).not.toHaveBeenCalled()
  })

  it('allows only exact owned calls', async () => {
    const handlers = new Map<string, (...args: unknown[]) => unknown>()
    const sender = {}
    const deps = {
      ipcMain: {
        handle: (channel: string, handler: (...args: unknown[]) => unknown) =>
          handlers.set(channel, handler),
      },
      shellSender: () => sender,
      recents: vi.fn().mockResolvedValue([]),
      create: vi.fn().mockResolvedValue(null),
      importProject: vi.fn().mockResolvedValue(null),
      open: vi.fn().mockResolvedValue(undefined),
    }
    registerLatexHomeIpc(deps)
    await handlers.get(HOME_CHANNELS.latexRecents)!({ sender })
    await handlers.get(HOME_CHANNELS.newLatexProject)!({ sender })
    await handlers.get(HOME_CHANNELS.importLatexProject)!({ sender })
    await handlers.get(HOME_CHANNELS.openLatexProject)!({ sender }, '/project')
    expect(deps.recents).toHaveBeenCalledOnce()
    expect(deps.create).toHaveBeenCalledOnce()
    expect(deps.importProject).toHaveBeenCalledOnce()
    expect(deps.open).toHaveBeenCalledWith('/project')
  })
})
