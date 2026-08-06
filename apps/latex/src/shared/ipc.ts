export const MAX_IPC_TEXT_BYTES = 2 * 1024 * 1024
export const MAX_IPC_PATH_LENGTH = 1024
export const MAX_IPC_FILES = 1_000

export const LATEX_CHANNELS = {
  sessionGet: 'latex:session:get',
  projectList: 'latex:project:list',
  fileRead: 'latex:file:read',
  fileUpdate: 'latex:file:update',
  fileSave: 'latex:file:save',
  fileCreate: 'latex:file:create',
  fileRename: 'latex:file:rename',
  compileStatus: 'latex:compile:status',
  compileStart: 'latex:compile:start',
  compileCancel: 'latex:compile:cancel',
  syncTexForward: 'latex:synctex:forward',
  syncTexReverse: 'latex:synctex:reverse',
  proposalGet: 'latex:proposal:get',
  proposalApply: 'latex:proposal:apply',
  proposalUndo: 'latex:proposal:undo',
  dirtyChanged: 'latex:dirty:changed',
  externalChanged: 'latex:external:changed',
  projectOpened: 'latex:project:opened',
  projectRenamed: 'latex:project:renamed',
} as const

export type LatexIpcErrorCode =
  | 'LATEX_FORBIDDEN_SENDER'
  | 'LATEX_PROJECT_SESSION_MISMATCH'
  | 'LATEX_INVALID_PAYLOAD'
  | 'LATEX_CONFLICT'
  | 'LATEX_NOT_FOUND'
  | 'LATEX_INTERNAL'

export type LatexIpcResult<T> =
  { ok: true; value: T } | { ok: false; error: { code: LatexIpcErrorCode; message: string } }

export interface SessionRequest {
  projectId: string
}

export interface FileRequest extends SessionRequest {
  path: string
}

export interface UpdateFileRequest extends FileRequest {
  text: string
}

export interface RenameFileRequest extends SessionRequest {
  from: string
  to: string
}

export interface CompileRequest extends SessionRequest {
  revision: number
  mainFile: string
}

export interface CompileRevisionRequest extends SessionRequest {
  revision: number
}

export interface SyncTexForwardRequest extends CompileRevisionRequest {
  path: string
  line: number
}

export interface SyncTexReverseRequest extends CompileRevisionRequest {
  page: number
  x: number
  y: number
}

export interface ProposalRequest extends SessionRequest {
  proposalId: string
}

export interface UndoProposalRequest extends SessionRequest {
  snapshotId: string
}

export interface LatexBufferDto {
  path: string
  text: string
  dirty: boolean
  conflict: { diskText: string | null } | null
}

export interface LatexSessionDto {
  projectId: string
  mainFile: string | null
  dirty: boolean
}

export interface CompileResultDto {
  revision: number
  pdfUrl: string | null
  diagnostics: readonly unknown[]
  log: string
}

export interface LatexApi {
  getSession(): Promise<LatexIpcResult<LatexSessionDto>>
  listFiles(request: SessionRequest): Promise<LatexIpcResult<string[]>>
  readFile(request: FileRequest): Promise<LatexIpcResult<LatexBufferDto>>
  updateFile(request: UpdateFileRequest): Promise<LatexIpcResult<LatexBufferDto>>
  saveFile(request: FileRequest): Promise<LatexIpcResult<LatexBufferDto>>
  createFile(request: UpdateFileRequest): Promise<LatexIpcResult<LatexBufferDto>>
  renameFile(request: RenameFileRequest): Promise<LatexIpcResult<void>>
  compile(request: CompileRequest): Promise<LatexIpcResult<CompileResultDto>>
  cancelCompile(request: SessionRequest): Promise<LatexIpcResult<{ cancelled: boolean }>>
  syncTexForward(
    request: SyncTexForwardRequest,
  ): Promise<LatexIpcResult<{ page: number; x: number; y: number } | null>>
  syncTexReverse(
    request: SyncTexReverseRequest,
  ): Promise<LatexIpcResult<{ path: string; line: number } | null>>
  getProposal(request: ProposalRequest): Promise<LatexIpcResult<unknown>>
  applyProposal(
    request: ProposalRequest,
  ): Promise<LatexIpcResult<{ proposalId: string; snapshotId: string }>>
  undoProposal(
    request: UndoProposalRequest,
  ): Promise<LatexIpcResult<{ snapshotId: string; restored: boolean }>>
  onExternalChange(handler: (buffer: LatexBufferDto) => void): () => void
}
