import { contextBridge, ipcRenderer } from 'electron'
import { LATEX_CHANNELS, type LatexApi } from '../shared/ipc.js'

const api: LatexApi = {
  getSession: () => ipcRenderer.invoke(LATEX_CHANNELS.sessionGet),
  listFiles: (request) => ipcRenderer.invoke(LATEX_CHANNELS.projectList, request),
  readFile: (request) => ipcRenderer.invoke(LATEX_CHANNELS.fileRead, request),
  updateFile: (request) => ipcRenderer.invoke(LATEX_CHANNELS.fileUpdate, request),
  saveFile: (request) => ipcRenderer.invoke(LATEX_CHANNELS.fileSave, request),
  createFile: (request) => ipcRenderer.invoke(LATEX_CHANNELS.fileCreate, request),
  renameFile: (request) => ipcRenderer.invoke(LATEX_CHANNELS.fileRename, request),
  compile: (request) => ipcRenderer.invoke(LATEX_CHANNELS.compileStart, request),
  cancelCompile: (request) => ipcRenderer.invoke(LATEX_CHANNELS.compileCancel, request),
  syncTexForward: (request) => ipcRenderer.invoke(LATEX_CHANNELS.syncTexForward, request),
  syncTexReverse: (request) => ipcRenderer.invoke(LATEX_CHANNELS.syncTexReverse, request),
  getProposal: (request) => ipcRenderer.invoke(LATEX_CHANNELS.proposalGet, request),
  proposeProjectEdits: (request) => ipcRenderer.invoke(LATEX_CHANNELS.proposalCreate, request),
  applyProposal: (request) => ipcRenderer.invoke(LATEX_CHANNELS.proposalApply, request),
  undoProposal: (request) => ipcRenderer.invoke(LATEX_CHANNELS.proposalUndo, request),
  listProjectFiles: (request) => ipcRenderer.invoke(LATEX_CHANNELS.aiProjectList, request),
  searchProjectText: (request) => ipcRenderer.invoke(LATEX_CHANNELS.aiProjectSearch, request),
  readProjectText: (request) => ipcRenderer.invoke(LATEX_CHANNELS.aiProjectRead, request),
  getCompileDiagnostics: (request) => ipcRenderer.invoke(LATEX_CHANNELS.aiDiagnosticsGet, request),
  compileProjectForAi: (request) => ipcRenderer.invoke(LATEX_CHANNELS.aiCompile, request),
  resolveDirectoryChat: (request) => ipcRenderer.invoke(LATEX_CHANNELS.aiChatResolve, request),
  appendDirectoryChat: (request) => ipcRenderer.invoke(LATEX_CHANNELS.aiChatAppend, request),
  loadDirectoryChat: (request) => ipcRenderer.invoke(LATEX_CHANNELS.aiChatLoad, request),
  getAiSettings: () => ipcRenderer.invoke('ai:get-settings'),
  aiStream: (request) => ipcRenderer.invoke('ai:stream', request),
  aiStreamCancel: (requestId) => ipcRenderer.invoke('ai:stream-cancel', requestId),
  onAiStream: (handler) => {
    const listener = (_event: Electron.IpcRendererEvent, chunk: Parameters<typeof handler>[0]) =>
      handler(chunk)
    ipcRenderer.on('ai:stream-chunk', listener)
    return () => ipcRenderer.removeListener('ai:stream-chunk', listener)
  },
  onExternalChange: (handler) => {
    const listener = (_event: Electron.IpcRendererEvent, buffer: Parameters<typeof handler>[0]) =>
      handler(buffer)
    ipcRenderer.on(LATEX_CHANNELS.externalChanged, listener)
    return () => ipcRenderer.removeListener(LATEX_CHANNELS.externalChanged, listener)
  },
  onEditFlushRequest: (handler) => {
    const listener = (_event: Electron.IpcRendererEvent, payload: unknown) => {
      if (!payload || typeof payload !== 'object') return
      const record = payload as Record<string, unknown>
      if (Object.keys(record).join(',') !== 'requestId' || typeof record.requestId !== 'string') {
        return
      }
      void Promise.resolve()
        .then(() => handler(record.requestId as string))
        .then((ok) =>
          ipcRenderer.send(LATEX_CHANNELS.editFlushAck, { requestId: record.requestId, ok }),
        )
        .catch(() =>
          ipcRenderer.send(LATEX_CHANNELS.editFlushAck, { requestId: record.requestId, ok: false }),
        )
    }
    ipcRenderer.on(LATEX_CHANNELS.editFlushRequest, listener)
    return () => ipcRenderer.removeListener(LATEX_CHANNELS.editFlushRequest, listener)
  },
  onEditFlushRelease: (handler) => {
    const listener = (_event: Electron.IpcRendererEvent, payload: unknown) => {
      if (!payload || typeof payload !== 'object') return
      const record = payload as Record<string, unknown>
      if (Object.keys(record).join(',') === 'requestId' && typeof record.requestId === 'string')
        handler(record.requestId)
    }
    ipcRenderer.on(LATEX_CHANNELS.editFlushRelease, listener)
    return () => ipcRenderer.removeListener(LATEX_CHANNELS.editFlushRelease, listener)
  },
}

contextBridge.exposeInMainWorld('latexApi', api)
