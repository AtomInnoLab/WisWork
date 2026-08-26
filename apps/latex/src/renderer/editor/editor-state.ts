export interface EditorConflict {
  diskText: string
  diskSha256: string | null
}

export interface EditorBuffer {
  path: string
  text: string
  diskText: string
  diskSha256: string
  editRevision: number
  dirty: boolean
  conflict: EditorConflict | null
}

export interface EditorPreview {
  revision: number
  pdfUrl: string
  compiledWorkspaceRevision: number
}

export interface EditorState {
  buffers: Record<string, EditorBuffer>
  workspaceRevision: number
  nextCompileRevision: number
  latestCompileRevision: number | null
  preview: EditorPreview | null
  previewStale: boolean
}

export interface SaveSnapshot {
  path: string
  text: string
  editRevision: number
  expectedSha256: string
}

export function createEditorState(
  files: readonly {
    path: string
    text: string
    diskText?: string
    diskSha256: string
    dirty?: boolean
    conflict?: EditorConflict | null
  }[],
): EditorState {
  return {
    buffers: Object.fromEntries(
      files.map((file) => [
        file.path,
        {
          path: file.path,
          text: file.text,
          diskText: file.diskText ?? file.text,
          diskSha256: file.diskSha256,
          editRevision: 0,
          dirty: file.dirty ?? false,
          conflict: file.conflict ?? null,
        },
      ]),
    ),
    workspaceRevision: 0,
    nextCompileRevision: 1,
    latestCompileRevision: null,
    preview: null,
    previewStale: false,
  }
}

export function restoreEditorPreview(
  state: EditorState,
  result: { revision: number; pdfUrl: string | null },
): EditorState {
  if (!result.pdfUrl) return state
  return {
    ...state,
    preview: {
      revision: result.revision,
      pdfUrl: result.pdfUrl,
      compiledWorkspaceRevision: state.workspaceRevision,
    },
    previewStale: false,
  }
}

export function editBuffer(state: EditorState, path: string, text: string): EditorState {
  const buffer = state.buffers[path]
  if (!buffer || buffer.text === text) return state
  const revision = state.workspaceRevision + 1
  return {
    ...state,
    workspaceRevision: revision,
    previewStale: state.preview !== null,
    buffers: {
      ...state.buffers,
      [path]: { ...buffer, text, editRevision: revision, dirty: text !== buffer.diskText },
    },
  }
}

export function beginSave(state: EditorState, path: string): SaveSnapshot {
  const buffer = state.buffers[path]
  if (!buffer) throw new Error(`Unknown editor buffer: ${path}`)
  return {
    path,
    text: buffer.text,
    editRevision: buffer.editRevision,
    expectedSha256: buffer.diskSha256,
  }
}

export function completeSave(
  state: EditorState,
  snapshot: SaveSnapshot,
  diskSha256: string,
): EditorState {
  const buffer = state.buffers[snapshot.path]
  if (!buffer) return state
  const savedCurrentRevision =
    buffer.editRevision === snapshot.editRevision &&
    buffer.text === snapshot.text &&
    buffer.conflict === null
  const dirty = !savedCurrentRevision
  if (
    buffer.diskText === snapshot.text &&
    buffer.diskSha256 === diskSha256 &&
    buffer.dirty === dirty
  )
    return state
  return {
    ...state,
    buffers: {
      ...state.buffers,
      [snapshot.path]: {
        ...buffer,
        diskText: snapshot.text,
        diskSha256,
        dirty,
        conflict: buffer.conflict,
      },
    },
  }
}

export function recordExternalChange(
  state: EditorState,
  path: string,
  diskText: string,
  diskSha256: string,
): EditorState {
  const buffer = state.buffers[path]
  if (!buffer) return state
  if (
    buffer.conflict?.diskSha256 === diskSha256 ||
    (buffer.diskSha256 === diskSha256 && buffer.diskText === diskText)
  )
    return state
  const workspaceRevision = state.workspaceRevision + 1
  const next = buffer.dirty
    ? { ...buffer, conflict: { diskText, diskSha256 } }
    : {
        ...buffer,
        text: diskText,
        diskText,
        diskSha256,
        conflict: null,
      }
  return {
    ...state,
    workspaceRevision,
    previewStale: state.preview !== null,
    buffers: { ...state.buffers, [path]: next },
  }
}

export function addEditorBuffer(
  state: EditorState,
  file: {
    path: string
    text: string
    diskText?: string
    diskSha256: string
    dirty?: boolean
    conflict?: EditorConflict | null
  },
  mode: 'initial' | 'tracked' | 'created',
): EditorState {
  if (state.buffers[file.path]) return state
  const changesWorkspace =
    mode === 'created' || (mode === 'tracked' && Boolean(file.dirty || file.conflict))
  const workspaceRevision = changesWorkspace ? state.workspaceRevision + 1 : state.workspaceRevision
  return {
    ...state,
    workspaceRevision,
    previewStale: changesWorkspace ? state.preview !== null : state.previewStale,
    buffers: {
      ...state.buffers,
      [file.path]: {
        path: file.path,
        text: file.text,
        diskSha256: file.diskSha256,
        diskText: file.diskText ?? file.text,
        editRevision: workspaceRevision,
        dirty: file.dirty ?? false,
        conflict: file.conflict ?? null,
      },
    },
  }
}

export function reconcileExternalBuffer(
  state: EditorState,
  incoming: {
    path: string
    text: string
    diskText?: string
    diskSha256: string
    dirty: boolean
    conflict: { diskText: string | null; diskSha256: string | null } | null
  },
  pendingSnapshot?: SaveSnapshot,
): EditorState {
  const current = state.buffers[incoming.path]
  const incomingDiskText = incoming.diskText ?? incoming.text
  if (!current) return state
  if (
    pendingSnapshot &&
    !incoming.dirty &&
    !incoming.conflict &&
    incoming.text === pendingSnapshot.text &&
    incomingDiskText === pendingSnapshot.text
  ) {
    return completeSave(state, pendingSnapshot, incoming.diskSha256)
  }
  if (incoming.conflict) {
    if (incoming.conflict.diskSha256 !== null) {
      return recordExternalChange(
        state,
        incoming.path,
        incoming.conflict.diskText ?? '',
        incoming.conflict.diskSha256,
      )
    }
    if (current.conflict?.diskSha256 === null) return state
    return {
      ...state,
      workspaceRevision: state.workspaceRevision + 1,
      previewStale: state.preview !== null,
      buffers: {
        ...state.buffers,
        [incoming.path]: {
          ...current,
          conflict: { diskText: incoming.conflict.diskText ?? '', diskSha256: null },
        },
      },
    }
  }
  if (incoming.dirty) return state
  if (current.dirty && incoming.text === current.text) {
    return {
      ...state,
      buffers: {
        ...state.buffers,
        [incoming.path]: {
          ...current,
          diskText: incoming.text,
          diskSha256: incoming.diskSha256,
          dirty: false,
          conflict: null,
        },
      },
    }
  }
  return recordExternalChange(state, incoming.path, incomingDiskText, incoming.diskSha256)
}

export function renameEditorBuffer(
  state: EditorState,
  from: string,
  to: string,
  confirmedClean = false,
): EditorState {
  const buffer = state.buffers[from]
  if (!buffer || state.buffers[to]) return state
  const buffers = { ...state.buffers }
  delete buffers[from]
  buffers[to] = {
    ...buffer,
    path: to,
    ...(confirmedClean ? { diskText: buffer.text, dirty: false, conflict: null } : {}),
  }
  return {
    ...state,
    workspaceRevision: state.workspaceRevision + 1,
    previewStale: state.preview !== null,
    buffers,
  }
}

export function removeEditorBuffer(state: EditorState, path: string): EditorState {
  if (!state.buffers[path]) return state
  const buffers = { ...state.buffers }
  delete buffers[path]
  return {
    ...state,
    buffers,
    workspaceRevision: state.workspaceRevision + 1,
    previewStale: state.preview !== null,
  }
}

export function beginCompile(state: EditorState): {
  state: EditorState
  request: { revision: number; workspaceRevision: number }
} {
  const revision = state.nextCompileRevision
  return {
    state: {
      ...state,
      nextCompileRevision: revision + 1,
      latestCompileRevision: revision,
    },
    request: { revision, workspaceRevision: state.workspaceRevision },
  }
}

export function acceptCompileResult(
  state: EditorState,
  result: { revision: number; pdfUrl: string | null; workspaceRevision?: number },
): EditorState {
  if (result.revision !== state.latestCompileRevision || !result.pdfUrl) return state
  const compiledWorkspaceRevision = result.workspaceRevision ?? state.workspaceRevision
  return {
    ...state,
    preview: { revision: result.revision, pdfUrl: result.pdfUrl, compiledWorkspaceRevision },
    previewStale: state.workspaceRevision !== compiledWorkspaceRevision,
  }
}
