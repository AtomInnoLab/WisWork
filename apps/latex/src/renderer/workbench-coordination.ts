import type { LatexApi, LatexBufferDto, LatexIpcResult, LatexSaveDto } from '../shared/ipc.js'
import {
  beginSave,
  completeSave,
  type EditorState,
  type SaveSnapshot,
} from './editor/editor-state.js'

type SaveFile = (
  request: Parameters<LatexApi['saveFile']>[0],
) => Promise<LatexIpcResult<LatexSaveDto>>

export class PendingSaveRegistry {
  private readonly entries = new Map<string, SaveSnapshot>()

  register(snapshot: SaveSnapshot): void {
    this.entries.set(snapshot.path, snapshot)
  }

  get(path: string): SaveSnapshot | undefined {
    return this.entries.get(path)
  }

  match(incoming: LatexBufferDto): SaveSnapshot | undefined {
    const snapshot = this.entries.get(incoming.path)
    if (
      !snapshot ||
      incoming.dirty ||
      incoming.conflict ||
      incoming.text !== snapshot.text ||
      incoming.diskText !== snapshot.text ||
      !incoming.diskSha256
    )
      return undefined
    return snapshot
  }

  clear(path: string, snapshot: SaveSnapshot): void {
    if (this.entries.get(path) === snapshot) this.entries.delete(path)
  }
}

export async function persistBuffer(options: {
  projectId: string
  path: string
  getState: () => EditorState
  setState: (state: EditorState) => void
  saveFile: SaveFile
  pendingSaves?: PendingSaveRegistry
}): Promise<boolean> {
  const before = options.getState()
  const buffer = before.buffers[options.path]
  if (!buffer) return false
  if (!buffer.dirty) return buffer.conflict === null
  if (buffer.conflict) return false
  const snapshot = beginSave(before, options.path)
  options.pendingSaves?.register(snapshot)
  try {
    const result = await options.saveFile({
      projectId: options.projectId,
      path: options.path,
      text: snapshot.text,
      editRevision: snapshot.editRevision,
    })
    if (!result.ok) return false
    const current = options.getState().buffers[options.path]
    const expectedDirty = current?.text !== result.value.savedText
    if (
      !current ||
      result.value.savedText !== snapshot.text ||
      result.value.buffer.path !== options.path ||
      result.value.buffer.text !== current.text ||
      result.value.buffer.dirty !== expectedDirty ||
      result.value.buffer.conflict !== null ||
      !result.value.diskSha256
    ) {
      return false
    }
    options.setState(completeSave(options.getState(), snapshot, result.value.diskSha256))
    return true
  } finally {
    options.pendingSaves?.clear(options.path, snapshot)
  }
}

export async function saveAllForCompile(
  paths: readonly string[],
  save: (path: string) => Promise<boolean>,
  getState: () => EditorState,
): Promise<boolean> {
  const initial = getState()
  const initialRevision = initial.workspaceRevision
  const initialKeys = Object.keys(initial.buffers).sort().join('\0')
  for (const path of paths) if (!(await save(path))) return false
  const current = getState()
  if (
    current.workspaceRevision !== initialRevision ||
    Object.keys(current.buffers).sort().join('\0') !== initialKeys
  )
    return false
  return Object.values(current.buffers).every((buffer) => !buffer.dirty && !buffer.conflict)
}

export class CompileRequestQueue {
  private running = false
  private pending: (() => Promise<void>) | null = null

  request(run: () => Promise<void>): void {
    if (this.running) {
      this.pending = run
      return
    }
    void this.drain(run)
  }

  cancelPending(): void {
    this.pending = null
  }

  private async drain(run: () => Promise<void>): Promise<void> {
    this.running = true
    try {
      await run()
    } finally {
      this.running = false
      const next = this.pending
      this.pending = null
      if (next) void this.drain(next)
    }
  }
}

export interface SyncRequestToken {
  id: number
  revision: number
  path: string
}

export class SyncRequestGate {
  private next = 0
  private current = 0

  begin(revision: number, path: string): SyncRequestToken {
    const token = { id: ++this.next, revision, path }
    this.current = token.id
    return token
  }

  invalidate(): void {
    this.current = ++this.next
  }

  accept(token: SyncRequestToken, revision: number, path: string): boolean {
    return token.id === this.current && token.revision === revision && token.path === path
  }
}

export function canRenameFile(path: string, mainFile: string | null): boolean {
  return path !== mainFile
}

export function remapWorkspacePaths(
  files: readonly string[],
  openPaths: readonly string[],
  activePath: string | null,
  from: string,
  to: string,
) {
  const remap = (path: string) => (path === from ? to : path)
  return {
    files: files.map(remap).sort(),
    openPaths: openPaths.map(remap),
    activePath: activePath === null ? null : remap(activePath),
  }
}

export async function completeReverseSync(options: {
  token: SyncRequestToken
  gate: SyncRequestGate
  current: () => { revision: number; path: string } | null
  load: () => Promise<unknown>
  activate: () => void
}): Promise<boolean> {
  const accepted = () => {
    const current = options.current()
    return Boolean(current && options.gate.accept(options.token, current.revision, current.path))
  }
  if (!accepted()) return false
  const loaded = await options.load()
  if (loaded === false || !accepted()) return false
  options.activate()
  return true
}

export async function runRenameFlow(options: {
  prompt: () => string | null
  from: string
  cancelTimer: () => void
  scheduleTimer: () => void
  save: (path: string) => Promise<boolean>
  isClean: () => boolean
  rename: (from: string, to: string) => Promise<boolean>
  apply: (from: string, to: string) => void
}): Promise<boolean> {
  const to = options.prompt()?.trim()
  if (!to || to === options.from) return false
  options.cancelTimer()
  try {
    if (!(await options.save(options.from)) || !options.isClean()) {
      options.scheduleTimer()
      return false
    }
    if (!(await options.rename(options.from, to))) {
      options.scheduleTimer()
      return false
    }
    options.apply(options.from, to)
    return true
  } catch {
    options.scheduleTimer()
    return false
  }
}
