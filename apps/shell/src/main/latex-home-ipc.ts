import { HOME_CHANNELS, type LatexRecentProjectEntry } from '../shared/home-api'

interface IpcMainLike {
  handle(channel: string, handler: (event: { sender: object }, ...args: unknown[]) => unknown): void
}

interface LatexHomeIpcDependencies {
  ipcMain: IpcMainLike
  shellSender(): object | null
  recents(): Promise<LatexRecentProjectEntry[]>
  create(): Promise<LatexRecentProjectEntry | null>
  importProject(): Promise<LatexRecentProjectEntry | null>
  open(path: string): Promise<unknown>
}

function assertOwnedCall(
  event: { sender: object },
  args: readonly unknown[],
  expectedArgs: number,
  shellSender: () => object | null,
): void {
  const owner = shellSender()
  if (!owner || event.sender !== owner) throw new Error('Untrusted IPC sender.')
  if (args.length !== expectedArgs) throw new Error('Invalid LaTeX Home IPC payload.')
}

export function registerLatexHomeIpc(deps: LatexHomeIpcDependencies): void {
  deps.ipcMain.handle(HOME_CHANNELS.latexRecents, async (event, ...args) => {
    assertOwnedCall(event, args, 0, deps.shellSender)
    return deps.recents()
  })
  deps.ipcMain.handle(HOME_CHANNELS.newLatexProject, async (event, ...args) => {
    assertOwnedCall(event, args, 0, deps.shellSender)
    return deps.create()
  })
  deps.ipcMain.handle(HOME_CHANNELS.importLatexProject, async (event, ...args) => {
    assertOwnedCall(event, args, 0, deps.shellSender)
    return deps.importProject()
  })
  deps.ipcMain.handle(HOME_CHANNELS.openLatexProject, async (event, ...args) => {
    assertOwnedCall(event, args, 1, deps.shellSender)
    const [path] = args
    if (typeof path !== 'string' || !path || path.length > 4096) {
      throw new Error('Invalid LaTeX Home IPC payload.')
    }
    return deps.open(path)
  })
}
