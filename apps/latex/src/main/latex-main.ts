import { createRequire } from 'node:module'
import { readFile } from 'node:fs/promises'
import type { BrowserWindow, WebContents, WebContentsView } from 'electron'
import { LATEX_CHANNELS } from '../shared/ipc.js'
import { LatexEditFlushCoordinator, type FlushWebContents } from './edit-flush.js'
import { registerLatexIpc } from './ipc.js'
import { ProjectSessionRegistry } from './project-session.js'
import type { TectonicBundleAsset } from '@wiswork/latex-compiler'

export interface LatexRuntimeConfig {
  preloadPath: string
  rendererUrl?: string
  rendererFile?: string
  tectonicPath: string
  userDataPath: string
  bundleAsset?: TectonicBundleAsset
}

export interface DirtyWebContents {
  id: number
  isDestroyed(): boolean
}

interface DialogLike {
  showMessageBox(parentOrOptions: unknown, maybeOptions?: unknown): Promise<{ response: number }>
}

let runtime: LatexRuntimeConfig | undefined
let registry = new ProjectSessionRegistry()
let ipcRegistered = false
let unregisterIpc: (() => void) | undefined
let editFlushCoordinator: LatexEditFlushCoordinator | undefined

export function configureLatexRuntime(config: LatexRuntimeConfig): void {
  if (runtime) throw new Error('LaTeX runtime is already configured')
  if (!config.preloadPath || !config.tectonicPath || !config.userDataPath) {
    throw new Error('Incomplete LaTeX runtime configuration')
  }
  if (!config.rendererUrl && !config.rendererFile) {
    throw new Error('LaTeX runtime requires a renderer URL or file')
  }
  runtime = { ...config }
  registry = new ProjectSessionRegistry({
    compilerRuntime: {
      tectonicPath: config.tectonicPath,
      userDataPath: config.userDataPath,
      bundleAsset: config.bundleAsset,
    },
    onExternalChange: (id, buffer) => {
      const electron = electronRuntime()
      electron.webContents.fromId(id)?.send(LATEX_CHANNELS.externalChanged, buffer)
    },
  })
}

export function createLatexView(projectPath: string): WebContentsView {
  if (!runtime) throw new Error('LaTeX runtime is not configured')
  const electron = electronRuntime()
  editFlushCoordinator ??= new LatexEditFlushCoordinator(electron.ipcMain)
  const ownedRuntime = runtime
  const ownedRegistry = registry
  if (!ipcRegistered) {
    unregisterIpc = registerLatexIpc({ ipcMain: electron.ipcMain, registry: ownedRegistry })
    ipcRegistered = true
  }
  const view = new electron.WebContentsView({
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  }) as WebContentsView
  view.webContents.once('destroyed', () => ownedRegistry.destroy(view.webContents.id))
  const load = ownedRuntime.rendererUrl
    ? view.webContents.loadURL(ownedRuntime.rendererUrl)
    : view.webContents.loadFile(ownedRuntime.rendererFile!)
  void Promise.all([load, ownedRegistry.attach(view.webContents.id, projectPath)])
    .then(([, session]) => {
      if (view.webContents.isDestroyed()) {
        ownedRegistry.destroy(view.webContents.id)
        return
      }
      view.webContents.send(LATEX_CHANNELS.projectOpened, {
        projectId: session.projectId,
        name: projectPath.split(/[\\/]/).at(-1) ?? 'LaTeX Project',
      })
    })
    .catch(async () => {
      await load.catch(() => undefined)
      if (!view.webContents.isDestroyed()) {
        view.webContents.send('latex:host-error', {
          code: 'LATEX_PROJECT_OPEN_FAILED',
          message: 'Unable to open LaTeX project',
        })
      }
    })
  return view
}

export function requestLatexEditFlush(
  contents: FlushWebContents,
  coordinator: LatexEditFlushCoordinator | undefined = editFlushCoordinator,
): Promise<boolean> {
  return coordinator ? coordinator.request(contents) : Promise.resolve(false)
}

export function releaseLatexEditFlush(
  contents: FlushWebContents,
  coordinator: LatexEditFlushCoordinator | undefined = editFlushCoordinator,
): void {
  coordinator?.release(contents)
}

export async function latexQueryDirty(
  contents: DirtyWebContents,
  sessions: ProjectSessionRegistry = registry,
): Promise<boolean> {
  if (contents.isDestroyed()) return false
  return sessions.getByWebContents(contents.id)?.isDirty() ?? false
}

export async function requestLatexClose(
  contents: DirtyWebContents,
  parent: BrowserWindow | null = null,
  sessions: ProjectSessionRegistry = registry,
  dialogOverride?: DialogLike,
): Promise<boolean> {
  const session = sessions.getByWebContents(contents.id)
  if (!session || contents.isDestroyed()) return true
  try {
    await session.settleSaves()
  } catch {
    return false
  }
  if (contents.isDestroyed()) return true
  if (!session.isDirty()) return true
  const dialog = dialogOverride ?? electronRuntime().dialog
  const options = {
    type: 'warning',
    message: 'This LaTeX project has unsaved changes.',
    detail: 'Save changes before closing?',
    buttons: ['Save', "Don't Save", 'Cancel'],
    defaultId: 0,
    cancelId: 2,
    noLink: true,
  }
  const { response } = parent
    ? await dialog.showMessageBox(parent, options)
    : await dialog.showMessageBox(options)
  if (response === 2) return false
  if (response === 1) {
    try {
      await session.settleSaves()
      await session.discardAll()
      return !session.isDirty()
    } catch {
      return false
    }
  }
  try {
    await session.saveAll()
    await session.settleSaves()
    return !session.isDirty()
  } catch {
    return false
  }
}

export interface ProtocolLike {
  handle(scheme: string, handler: (request: { url: string }) => Promise<Response>): void
}

/** Shell owns calling this once after app.ready; standalone development may call it itself. */
export function registerLatexPdfProtocol(protocol: ProtocolLike, sessions = registry): void {
  protocol.handle('wiswork-latex-pdf', async ({ url }) => {
    let parsed: URL
    try {
      parsed = new URL(url)
    } catch {
      return new Response('Invalid PDF URL', { status: 400 })
    }
    const projectId = parsed.hostname
    const revisionText = parsed.pathname.replace(/^\//, '')
    if (!/^[a-f0-9]{32}$/.test(projectId) || !/^(0|[1-9]\d*)$/.test(revisionText)) {
      return new Response('Invalid PDF URL', { status: 400 })
    }
    const path = sessions.resolvePdf(projectId, Number(revisionText))
    if (!path) return new Response('PDF not found', { status: 404 })
    try {
      return new Response(await readFile(path), {
        headers: { 'content-type': 'application/pdf', 'cache-control': 'no-store' },
      })
    } catch {
      return new Response('PDF not found', { status: 404 })
    }
  })
}

export function getLatexSessionRegistry(): ProjectSessionRegistry {
  return registry
}

/** @internal Tests only; production configuration is single-assignment. */
export function resetLatexRuntimeForTests(): void {
  registry.disposeAll()
  editFlushCoordinator?.dispose()
  editFlushCoordinator = undefined
  unregisterIpc?.()
  unregisterIpc = undefined
  runtime = undefined
  registry = new ProjectSessionRegistry()
  ipcRegistered = false
}

export function notifyLatexProjectRenamed(contents: WebContents, name: string): void {
  if (contents.isDestroyed() || !name || name.length > 255 || /[\\/\0]/.test(name)) {
    throw new Error('Invalid LaTeX project name')
  }
  contents.send(LATEX_CHANNELS.projectRenamed, { name })
}

interface ElectronRuntime {
  WebContentsView: new (options: unknown) => WebContentsView
  ipcMain: Parameters<typeof registerLatexIpc>[0]['ipcMain'] &
    ConstructorParameters<typeof LatexEditFlushCoordinator>[0]
  dialog: DialogLike
  webContents: { fromId(id: number): WebContents | undefined }
}

function electronRuntime(): ElectronRuntime {
  const required = createRequire(import.meta.url)('electron') as ElectronRuntime | string
  if (typeof required === 'string') throw new Error('Electron runtime is unavailable')
  return required
}
