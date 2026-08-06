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
  applyProposal: (request) => ipcRenderer.invoke(LATEX_CHANNELS.proposalApply, request),
  undoProposal: (request) => ipcRenderer.invoke(LATEX_CHANNELS.proposalUndo, request),
  onExternalChange: (handler) => {
    const listener = (_event: Electron.IpcRendererEvent, buffer: Parameters<typeof handler>[0]) =>
      handler(buffer)
    ipcRenderer.on(LATEX_CHANNELS.externalChanged, listener)
    return () => ipcRenderer.removeListener(LATEX_CHANNELS.externalChanged, listener)
  },
}

contextBridge.exposeInMainWorld('latexApi', api)
