import type { AiSettings, AiStreamChunk, AiStreamRequest } from '@wiswork/ai-provider'

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
  bundleStatus: 'latex:bundle:status',
  compileStart: 'latex:compile:start',
  compileCancel: 'latex:compile:cancel',
  syncTexForward: 'latex:synctex:forward',
  syncTexReverse: 'latex:synctex:reverse',
  proposalGet: 'latex:proposal:get',
  proposalCreate: 'latex:proposal:create',
  proposalDiscard: 'latex:proposal:discard',
  proposalApply: 'latex:proposal:apply',
  proposalUndo: 'latex:proposal:undo',
  codexMutationRevision: 'latex:codex:mutation:revision',
  codexProposalPrepare: 'latex:codex:proposal:prepare',
  codexProposalExecute: 'latex:codex:proposal:execute',
  codexProposalDiscard: 'latex:codex:proposal:discard',
  aiProjectList: 'latex:ai:project:list',
  aiProjectSearch: 'latex:ai:project:search',
  aiProjectRead: 'latex:ai:project:read',
  aiDiagnosticsGet: 'latex:ai:diagnostics:get',
  aiCompile: 'latex:ai:compile',
  aiChatResolve: 'latex:ai:chat:resolve',
  aiChatAppend: 'latex:ai:chat:append',
  aiChatLoad: 'latex:ai:chat:load',
  dirtyChanged: 'latex:dirty:changed',
  externalChanged: 'latex:external:changed',
  projectOpened: 'latex:project:opened',
  projectRenamed: 'latex:project:renamed',
  editFlushRequest: 'latex:edit-flush:request',
  editFlushAck: 'latex:edit-flush:ack',
  editFlushRelease: 'latex:edit-flush:release',
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

export interface SaveFileRequest extends UpdateFileRequest {
  editRevision: number
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

export interface CreateProposalRequest extends SessionRequest {
  files: Array<{ path: string; afterText: string }>
}

export interface PrepareCodexProposalMutationRequest extends SessionRequest {
  documentId: string
  callId: string
  proposalId: string
  expectedRevision: string
}

export interface ExecuteCodexProposalMutationRequest extends SessionRequest {
  documentId: string
  callId: string
  preparationId: string
  snapshotId: string
  expectedRevision: string
}

export interface DiscardCodexProposalMutationRequest extends SessionRequest {
  documentId: string
  callId: string
  preparationId: string
  snapshotId: string
}

export interface PreparedCodexProposalMutationDto {
  preparationId: string
  snapshotId: string
}

export interface CodexProposalMutationResultDto {
  proposalId: string
  snapshotId: string
  compile: { ok: true; result: CompileResultDto } | { ok: false; error: string }
}

export interface AiProjectReadRequest extends FileRequest {
  offset: number
  maxChars: number
}

export interface AiProjectSearchRequest extends SessionRequest {
  query: string
  maxResults: number
}

export interface AiChatAppendRequest extends SessionRequest {
  storeProjectId: string
  chatId: string
  role: 'user' | 'assistant'
  text: string
}

export interface AiChatLoadRequest extends SessionRequest {
  storeProjectId: string
  chatId: string
  limit: number
}

export interface LatexProposalDto {
  id: string
  projectId: string
  expiresAt: number
  files: Array<{
    path: string
    beforeText: string | null
    beforeSha256: string | null
    afterText: string
  }>
}

export interface LatexBufferDto {
  path: string
  text: string
  diskText: string
  diskSha256: string
  dirty: boolean
  conflict: { diskText: string | null; diskSha256: string | null } | null
}

export interface LatexSessionDto {
  projectId: string
  mainFile: string | null
  dirty: boolean
}

export interface LatexSaveDto {
  savedText: string
  diskSha256: string
  buffer: LatexBufferDto
}

export interface CompileResultDto {
  revision: number
  pdfUrl: string | null
  diagnostics: readonly unknown[]
  log: string
}

export type LatexBundleStatusDto =
  | { state: 'missing' }
  | { state: 'downloading'; receivedBytes: number; totalBytes: number }
  | { state: 'ready'; bytes: number }
  | { state: 'remote' }
  | { state: 'error'; code: string }

export interface LatexApi {
  getSession(): Promise<LatexIpcResult<LatexSessionDto>>
  listFiles(request: SessionRequest): Promise<LatexIpcResult<string[]>>
  readFile(request: FileRequest): Promise<LatexIpcResult<LatexBufferDto>>
  updateFile(request: UpdateFileRequest): Promise<LatexIpcResult<LatexBufferDto>>
  saveFile(request: SaveFileRequest): Promise<LatexIpcResult<LatexSaveDto>>
  createFile(request: UpdateFileRequest): Promise<LatexIpcResult<LatexBufferDto>>
  renameFile(request: RenameFileRequest): Promise<LatexIpcResult<void>>
  compile(request: CompileRequest): Promise<LatexIpcResult<CompileResultDto>>
  cancelCompile(request: SessionRequest): Promise<LatexIpcResult<{ cancelled: boolean }>>
  getBundleStatus(request: SessionRequest): Promise<LatexIpcResult<LatexBundleStatusDto>>
  syncTexForward(
    request: SyncTexForwardRequest,
  ): Promise<LatexIpcResult<{ page: number; x: number; y: number } | null>>
  syncTexReverse(
    request: SyncTexReverseRequest,
  ): Promise<LatexIpcResult<{ path: string; line: number } | null>>
  getProposal(request: ProposalRequest): Promise<LatexIpcResult<unknown>>
  proposeProjectEdits(request: CreateProposalRequest): Promise<LatexIpcResult<LatexProposalDto>>
  discardProposal(request: ProposalRequest): Promise<LatexIpcResult<void>>
  applyProposal(request: ProposalRequest): Promise<LatexIpcResult<unknown>>
  undoProposal(request: UndoProposalRequest): Promise<LatexIpcResult<unknown>>
  getCodexMutationRevision(request: SessionRequest): Promise<LatexIpcResult<{ revision: string }>>
  prepareCodexProposalMutation(
    request: PrepareCodexProposalMutationRequest,
  ): Promise<LatexIpcResult<PreparedCodexProposalMutationDto>>
  executeCodexProposalMutation(
    request: ExecuteCodexProposalMutationRequest,
  ): Promise<LatexIpcResult<CodexProposalMutationResultDto>>
  discardCodexProposalMutation(
    request: DiscardCodexProposalMutationRequest,
  ): Promise<LatexIpcResult<void>>
  listProjectFiles(request: SessionRequest): Promise<LatexIpcResult<unknown>>
  searchProjectText(request: AiProjectSearchRequest): Promise<LatexIpcResult<unknown>>
  readProjectText(request: AiProjectReadRequest): Promise<LatexIpcResult<unknown>>
  getCompileDiagnostics(request: SessionRequest): Promise<LatexIpcResult<unknown>>
  compileProjectForAi(request: SessionRequest): Promise<LatexIpcResult<unknown>>
  resolveDirectoryChat(
    request: SessionRequest,
  ): Promise<LatexIpcResult<{ projectId: string; chatId: string }>>
  appendDirectoryChat(request: AiChatAppendRequest): Promise<LatexIpcResult<void>>
  loadDirectoryChat(
    request: AiChatLoadRequest,
  ): Promise<LatexIpcResult<Array<{ role: 'user' | 'assistant'; text: string }>>>
  getAiSettings(): Promise<AiSettings>
  aiStream(request: AiStreamRequest): Promise<void>
  aiStreamCancel(requestId: string): Promise<void>
  onAiStream(handler: (chunk: AiStreamChunk) => void): () => void
  onExternalChange(handler: (buffer: LatexBufferDto) => void): () => void
  onEditFlushRequest(handler: (requestId: string) => Promise<boolean>): () => void
  onEditFlushRelease(handler: (requestId: string) => void): () => void
}
