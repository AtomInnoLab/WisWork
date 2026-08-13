import { ProjectWriteConflictError } from '@wiswork/latex-project'
import {
  LATEX_CHANNELS,
  MAX_IPC_PATH_LENGTH,
  MAX_IPC_TEXT_BYTES,
  type CompileRequest,
  type FileRequest,
  type LatexIpcErrorCode,
  type LatexIpcResult,
  type LatexAccountStatus,
  type RenameFileRequest,
  type SaveFileRequest,
  type SessionRequest,
  type UpdateFileRequest,
} from '../shared/ipc.js'
import {
  MainFileRenameError,
  ProposalVerificationRejectedError,
  UnsavedBuffersError,
  type ProjectSession,
  type ProjectSessionRegistry,
} from './project-session.js'

export interface IpcEventLike {
  sender: { id: number }
}

export interface IpcMainLike {
  handle(channel: string, handler: (event: IpcEventLike, payload?: unknown) => unknown): void
  removeHandler(channel: string): void
}

export interface RegisterLatexIpcOptions {
  ipcMain: IpcMainLike
  registry: ProjectSessionRegistry
  getAccountStatus?: () => Promise<LatexAccountStatus> | LatexAccountStatus
}

export function registerLatexIpc(options: RegisterLatexIpcOptions): () => void {
  const { ipcMain, registry } = options
  const channels: string[] = []
  const handle = (
    channel: string,
    action: (event: IpcEventLike, payload: unknown) => Promise<unknown> | unknown,
  ) => {
    channels.push(channel)
    ipcMain.handle(channel, async (event, payload) => {
      try {
        return ok(await action(event, payload))
      } catch (error) {
        return fail(error)
      }
    })
  }

  handle(LATEX_CHANNELS.sessionGet, async (event, payload) => {
    assertNoPayload(payload)
    const session = registry.getByWebContents(event.sender.id)
    if (!session) throw coded('LATEX_FORBIDDEN_SENDER', 'Sender has no LaTeX project session')
    let account: LatexAccountStatus = { loggedIn: false }
    try {
      account = normalizeAccountStatus(await options.getAccountStatus?.())
    } catch {
      // Account refresh must never prevent a project from opening.
    }
    return {
      projectId: session.projectId,
      mainFile: session.mainFile ?? null,
      dirty: session.isDirty(),
      latestCompile: session.latestCompile(),
      account,
    }
  })
  handle(LATEX_CHANNELS.projectList, (event, payload) =>
    withSession(event, payload, registry, parseSessionRequest, (session) =>
      session.listTextFiles(),
    ),
  )
  handle(LATEX_CHANNELS.fileRead, (event, payload) =>
    withSession(event, payload, registry, parseFileRequest, (session, request) =>
      session.readText(request.path),
    ),
  )
  handle(LATEX_CHANNELS.fileUpdate, (event, payload) =>
    withSession(event, payload, registry, parseUpdateRequest, (session, request) =>
      session.updateBuffer(request.path, request.text),
    ),
  )
  handle(LATEX_CHANNELS.fileSave, (event, payload) =>
    withSession(event, payload, registry, parseSaveRequest, (session, request) =>
      session.saveText(request.path, request.text, request.editRevision),
    ),
  )
  handle(LATEX_CHANNELS.fileCreate, (event, payload) =>
    withSession(event, payload, registry, parseUpdateRequest, (session, request) =>
      session.createText(request.path, request.text),
    ),
  )
  handle(LATEX_CHANNELS.fileRename, (event, payload) =>
    withSession(event, payload, registry, parseRenameRequest, (session, request) =>
      session.renameText(request.from, request.to),
    ),
  )
  handle(LATEX_CHANNELS.fileDelete, (event, payload) =>
    withSession(event, payload, registry, parseFileRequest, (session, request) =>
      session.deleteText(request.path),
    ),
  )
  handle(LATEX_CHANNELS.compileStart, (event, payload) =>
    withSession(event, payload, registry, parseCompileRequest, (session, request) =>
      session.compile(request.revision, request.mainFile),
    ),
  )
  handle(LATEX_CHANNELS.compileCancel, (event, payload) =>
    withSession(event, payload, registry, parseSessionRequest, (session) => ({
      cancelled: session.cancelCompile(),
    })),
  )
  handle(LATEX_CHANNELS.bundleStatus, (event, payload) =>
    withSession(event, payload, registry, parseSessionRequest, (session) =>
      session.getBundleStatus(),
    ),
  )
  handle(LATEX_CHANNELS.compileStatus, (event, payload) =>
    withSession(event, payload, registry, parseRevisionRequest, (session, request) => ({
      revision: request.revision,
      pdfUrl: session.pdfPath(request.revision)
        ? `wiswork-latex-pdf://${session.projectId}/${request.revision}`
        : null,
    })),
  )
  handle(LATEX_CHANNELS.syncTexForward, (event, payload) =>
    withSession(event, payload, registry, parseSyncTexForward, (session, request) =>
      session.syncTexForward(request.revision, request.path, request.line),
    ),
  )
  handle(LATEX_CHANNELS.syncTexReverse, (event, payload) =>
    withSession(event, payload, registry, parseSyncTexReverse, (session, request) =>
      session.syncTexReverse(request.revision, request.page, request.x, request.y),
    ),
  )
  handle(LATEX_CHANNELS.aiProjectList, (event, payload) =>
    withSession(event, payload, registry, parseSessionRequest, (session) =>
      session.listProjectFilesForAi(),
    ),
  )
  handle(LATEX_CHANNELS.aiProjectRead, (event, payload) =>
    withSession(event, payload, registry, parseAiReadRequest, (session, request) =>
      session.readProjectTextForAi(request.path, request.offset, request.maxChars),
    ),
  )
  handle(LATEX_CHANNELS.aiProjectSearch, (event, payload) =>
    withSession(event, payload, registry, parseAiSearchRequest, (session, request) =>
      session.searchProjectTextForAi(request.query, request.maxResults),
    ),
  )
  handle(LATEX_CHANNELS.aiDiagnosticsGet, (event, payload) =>
    withSession(event, payload, registry, parseSessionRequest, (session) =>
      session.getCompileDiagnosticsForAi(),
    ),
  )
  handle(LATEX_CHANNELS.aiCompile, (event, payload) =>
    withSession(event, payload, registry, parseSessionRequest, (session) => session.compileForAi()),
  )
  handle(LATEX_CHANNELS.aiChatResolve, (event, payload) =>
    withSession(event, payload, registry, parseSessionRequest, (session) =>
      session.resolveDirectoryChat(),
    ),
  )
  handle(LATEX_CHANNELS.aiChatAppend, (event, payload) =>
    withSession(event, payload, registry, parseAiChatAppend, (session, request) =>
      session.appendDirectoryChat(
        request.storeProjectId,
        request.chatId,
        request.role,
        request.text,
      ),
    ),
  )
  handle(LATEX_CHANNELS.aiChatLoad, (event, payload) =>
    withSession(event, payload, registry, parseAiChatLoad, (session, request) =>
      session.loadDirectoryChat(request.storeProjectId, request.chatId, request.limit),
    ),
  )
  handle(LATEX_CHANNELS.proposalCreate, (event, payload) =>
    withSession(event, payload, registry, parseCreateProposalRequest, (session, request) =>
      session.createEditProposal(request.files),
    ),
  )
  handle(LATEX_CHANNELS.proposalGet, (event, payload) =>
    withSession(event, payload, registry, parseProposalRequest, (session, request) => {
      const proposal = session.getProposal(request.proposalId)
      if (!proposal) throw coded('LATEX_NOT_FOUND', 'Proposal not found')
      return proposal
    }),
  )
  handle(LATEX_CHANNELS.proposalVerify, (event, payload) =>
    withSession(event, payload, registry, parseProposalRequest, (session, request) =>
      session.verifyProposal(request.proposalId),
    ),
  )
  handle(LATEX_CHANNELS.proposalApply, (event, payload) =>
    withSession(event, payload, registry, parseProposalRequest, (session, request) =>
      session.applyConfirmedProposal(request.proposalId),
    ),
  )
  handle(LATEX_CHANNELS.proposalUndo, (event, payload) =>
    withSession(event, payload, registry, parseUndoRequest, (session, request) =>
      session.undoConfirmedProposal(request.snapshotId),
    ),
  )

  return () => {
    for (const channel of channels) ipcMain.removeHandler(channel)
  }
}

function normalizeAccountStatus(value: unknown): LatexAccountStatus {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return { loggedIn: false }
  const item = value as Record<string, unknown>
  if (item.loggedIn !== true) return { loggedIn: false }
  const email = typeof item.email === 'string' && item.email.length <= 320 ? item.email : undefined
  const userId =
    typeof item.userId === 'string' && item.userId.length <= 512 ? item.userId : undefined
  return { loggedIn: true, ...(email ? { email } : {}), ...(userId ? { userId } : {}) }
}

async function withSession<TRequest extends SessionRequest, TResult>(
  event: IpcEventLike,
  payload: unknown,
  registry: ProjectSessionRegistry,
  parse: (value: unknown) => TRequest,
  action: (session: ProjectSession, request: TRequest) => Promise<TResult> | TResult,
): Promise<TResult> {
  const senderSession = registry.getByWebContents(event.sender.id)
  if (!senderSession) {
    throw coded('LATEX_FORBIDDEN_SENDER', 'Sender has no LaTeX project session')
  }
  const request = parse(payload)
  if (senderSession.projectId !== request.projectId) {
    throw coded('LATEX_PROJECT_SESSION_MISMATCH', 'Project does not belong to sender')
  }
  return action(senderSession, request)
}

function parseSessionRequest(value: unknown): SessionRequest {
  const item = exactObject(value, ['projectId'])
  return { projectId: boundedString(item.projectId, 'projectId', 128) }
}

function parseFileRequest(value: unknown): FileRequest {
  const item = exactObject(value, ['projectId', 'path'])
  return {
    projectId: boundedString(item.projectId, 'projectId', 128),
    path: relativePath(item.path),
  }
}

function parseUpdateRequest(value: unknown): UpdateFileRequest {
  const item = exactObject(value, ['projectId', 'path', 'text'])
  const text =
    typeof item.text === 'string' && item.text.length <= MAX_IPC_TEXT_BYTES
      ? item.text
      : invalid('text is invalid')
  if (Buffer.byteLength(text, 'utf8') > MAX_IPC_TEXT_BYTES) invalid('text is too large')
  return {
    projectId: boundedString(item.projectId, 'projectId', 128),
    path: relativePath(item.path),
    text,
  }
}

function parseSaveRequest(value: unknown): SaveFileRequest {
  const item = exactObject(value, ['projectId', 'path', 'text', 'editRevision'])
  const update = parseUpdateRequest({ projectId: item.projectId, path: item.path, text: item.text })
  return { ...update, editRevision: revision(item.editRevision) }
}

function parseRenameRequest(value: unknown): RenameFileRequest {
  const item = exactObject(value, ['projectId', 'from', 'to'])
  return {
    projectId: boundedString(item.projectId, 'projectId', 128),
    from: relativePath(item.from),
    to: relativePath(item.to),
  }
}

function parseCompileRequest(value: unknown): CompileRequest {
  const item = exactObject(value, ['projectId', 'revision', 'mainFile'])
  return {
    projectId: boundedString(item.projectId, 'projectId', 128),
    revision: revision(item.revision),
    mainFile: relativePath(item.mainFile),
  }
}

function parseRevisionRequest(value: unknown): SessionRequest & { revision: number } {
  const item = exactObject(value, ['projectId', 'revision'])
  return {
    projectId: boundedString(item.projectId, 'projectId', 128),
    revision: revision(item.revision),
  }
}

function parseSyncTexForward(value: unknown) {
  const item = exactObject(value, ['projectId', 'revision', 'path', 'line'])
  return {
    projectId: boundedString(item.projectId, 'projectId', 128),
    revision: revision(item.revision),
    path: relativePath(item.path),
    line: positiveInteger(item.line, 'line'),
  }
}

function parseSyncTexReverse(value: unknown) {
  const item = exactObject(value, ['projectId', 'revision', 'page', 'x', 'y'])
  return {
    projectId: boundedString(item.projectId, 'projectId', 128),
    revision: revision(item.revision),
    page: positiveInteger(item.page, 'page'),
    x: coordinate(item.x),
    y: coordinate(item.y),
  }
}

function parseAiReadRequest(value: unknown) {
  const item = exactObject(value, ['projectId', 'path', 'offset', 'maxChars'])
  return {
    projectId: boundedString(item.projectId, 'projectId', 128),
    path: relativePath(item.path),
    offset: revision(item.offset),
    maxChars: boundedPositiveInteger(item.maxChars, 'maxChars', 24_000),
  }
}

function parseAiSearchRequest(value: unknown) {
  const item = exactObject(value, ['projectId', 'query', 'maxResults'])
  return {
    projectId: boundedString(item.projectId, 'projectId', 128),
    query: boundedString(item.query, 'query', 256),
    maxResults: boundedPositiveInteger(item.maxResults, 'maxResults', 50),
  }
}

function parseAiChatAppend(value: unknown) {
  const item = exactObject(value, ['projectId', 'storeProjectId', 'chatId', 'role', 'text'])
  const role =
    item.role === 'user' || item.role === 'assistant' ? item.role : invalid('role is invalid')
  return {
    projectId: boundedString(item.projectId, 'projectId', 128),
    storeProjectId: boundedString(item.storeProjectId, 'storeProjectId', 128),
    chatId: boundedString(item.chatId, 'chatId', 128),
    role: role as 'user' | 'assistant',
    text:
      typeof item.text === 'string' && item.text.length <= 100_000
        ? item.text
        : invalid('text is invalid'),
  }
}

function parseAiChatLoad(value: unknown) {
  const item = exactObject(value, ['projectId', 'storeProjectId', 'chatId', 'limit'])
  return {
    projectId: boundedString(item.projectId, 'projectId', 128),
    storeProjectId: boundedString(item.storeProjectId, 'storeProjectId', 128),
    chatId: boundedString(item.chatId, 'chatId', 128),
    limit: boundedPositiveInteger(item.limit, 'limit', 200),
  }
}

function parseCreateProposalRequest(value: unknown) {
  const item = exactObject(value, ['projectId', 'files'])
  if (!Array.isArray(item.files) || item.files.length < 1 || item.files.length > 20)
    invalid('files are invalid')
  let totalBytes = 0
  const files = item.files.map((raw) => {
    const file = exactObject(raw, ['path', 'afterText'])
    if (typeof file.afterText !== 'string') invalid('afterText is invalid')
    const fileBytes = Buffer.byteLength(file.afterText, 'utf8')
    if (fileBytes > MAX_IPC_TEXT_BYTES) invalid('proposal file is too large')
    totalBytes += fileBytes
    if (totalBytes > 2 * MAX_IPC_TEXT_BYTES) invalid('proposal is too large')
    return { path: relativePath(file.path), afterText: file.afterText }
  })
  return { projectId: boundedString(item.projectId, 'projectId', 128), files }
}

function parseProposalRequest(value: unknown) {
  const item = exactObject(value, ['projectId', 'proposalId'])
  return {
    projectId: boundedString(item.projectId, 'projectId', 128),
    proposalId: boundedString(item.proposalId, 'proposalId', 128),
  }
}

function parseUndoRequest(value: unknown) {
  const item = exactObject(value, ['projectId', 'snapshotId'])
  return {
    projectId: boundedString(item.projectId, 'projectId', 128),
    snapshotId: boundedString(item.snapshotId, 'snapshotId', 128),
  }
}

function exactObject(value: unknown, keys: readonly string[]): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    invalid('payload must be an object')
  const item = value as Record<string, unknown>
  const actual = Object.keys(item).sort()
  const expected = [...keys].sort()
  if (JSON.stringify(actual) !== JSON.stringify(expected)) invalid('payload fields are invalid')
  return item
}

function relativePath(value: unknown): string {
  const path = boundedString(value, 'path', MAX_IPC_PATH_LENGTH)
  const normalized = path.replaceAll('\\', '/')
  if (
    normalized.includes('\0') ||
    normalized.startsWith('/') ||
    /^[A-Za-z]:\//.test(normalized) ||
    normalized.split('/').some((part) => part === '..')
  ) {
    invalid('path must remain inside the project')
  }
  const canonical = normalized
    .split('/')
    .filter((part) => part && part !== '.')
    .join('/')
  if (!canonical) invalid('path must be non-empty')
  return canonical
}

function boundedString(value: unknown, name: string, maxLength: number): string {
  if (typeof value !== 'string' || value.length === 0 || value.length > maxLength) {
    invalid(`${name} is invalid`)
  }
  return value
}

function revision(value: unknown): number {
  if (!Number.isSafeInteger(value) || (value as number) < 0) invalid('revision is invalid')
  return value as number
}

function boundedPositiveInteger(value: unknown, name: string, max: number): number {
  if (!Number.isSafeInteger(value) || (value as number) < 1 || (value as number) > max)
    invalid(name + ' is invalid')
  return value as number
}

function positiveInteger(value: unknown, name: string): number {
  if (!Number.isSafeInteger(value) || (value as number) < 1 || (value as number) > 10_000_000)
    invalid(`${name} is invalid`)
  return value as number
}

function coordinate(value: unknown): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || Math.abs(value) > 10_000_000)
    invalid('coordinate is invalid')
  return value
}

function assertNoPayload(payload: unknown): void {
  if (payload !== undefined) invalid('channel does not accept a payload')
}

function ok<T>(value: T): LatexIpcResult<T> {
  return { ok: true, value }
}

function fail(error: unknown): LatexIpcResult<never> {
  if (error instanceof CodedIpcError) {
    return { ok: false, error: { code: error.code, message: error.message } }
  }
  if (
    error instanceof MainFileRenameError ||
    (error instanceof Error && /configured main file/i.test(error.message))
  ) {
    return { ok: false, error: { code: 'LATEX_INVALID_PAYLOAD', message: error.message } }
  }
  if (error instanceof UnsavedBuffersError) {
    return {
      ok: false,
      error: { code: 'LATEX_CONFLICT', message: 'Project has unsaved LaTeX changes' },
    }
  }
  if (error instanceof ProposalVerificationRejectedError) {
    if (error.reason === 'not-found' || error.reason === 'expired') {
      return { ok: false, error: { code: 'LATEX_NOT_FOUND', message: error.message } }
    }
    if (error.reason === 'baseline') {
      return { ok: false, error: { code: 'LATEX_CONFLICT', message: error.message } }
    }
    return {
      ok: false,
      error: { code: 'LATEX_VERIFICATION_REJECTED', message: error.message },
    }
  }
  if (
    error instanceof ProjectWriteConflictError ||
    (error instanceof Error && /conflict|changed externally/i.test(error.message))
  ) {
    return { ok: false, error: { code: 'LATEX_CONFLICT', message: 'Project changed on disk' } }
  }
  return {
    ok: false,
    error: {
      code: 'LATEX_INTERNAL',
      message: 'LaTeX operation failed',
    },
  }
}

class CodedIpcError extends Error {
  constructor(
    readonly code: LatexIpcErrorCode,
    message: string,
  ) {
    super(message)
  }
}

function coded(code: LatexIpcErrorCode, message: string): CodedIpcError {
  return new CodedIpcError(code, message)
}

function invalid(message: string): never {
  throw coded('LATEX_INVALID_PAYLOAD', message)
}
