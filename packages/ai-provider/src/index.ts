export type {
  AiChatRequest,
  AiChatResponse,
  AiProviderConfig,
  AiProviderId,
  AiProviderMeta,
  AiSettings,
  AiServiceDiagnostic,
  AiStreamChunk,
  AiStreamRequest,
  LegacyAiSettings,
  WisworkFetchWithAuth,
} from './types'
export {
  AI_PROVIDERS,
  WISWORK_DEFAULT_MODEL,
  WISWORK_MESSAGES_URL,
  WISWORK_REQUEST_LOCATION,
  defaultAiSettings,
  resolveAiSettings,
} from './providers'
export { resolveWisworkMainRequest, sanitizeWisworkSettings } from './main-config'
export type { WisworkMainRequest, WisworkRequestErrorCode } from './main-config'
export {
  AI_IPC_LIMITS,
  AiIpcError,
  registerWisworkModelIpc,
  validateAiChatRequest,
  validateAiSearchArgs,
  validateAiSettings,
  validateAiStreamRequest,
} from './ipc'
export type {
  AiIpcErrorCode,
  IpcMainLike,
  RegisterWisworkModelIpcOptions,
  WisworkIpcEvent,
  WisworkIpcSender,
  WisworkModelIpcChannels,
} from './ipc'
export { AiProviderError, safeHttpProviderError } from './errors'
export type { AiProviderErrorCode } from './errors'
export { chatForProvider } from './chat'
export { AiCreditsError, sseLines, streamForProvider } from './stream'
export type { StreamCallbacks } from './stream'
export {
  AI_CHAT_RESPONSE_TIMEOUT_MS,
  AI_CONNECT_TIMEOUT_MS,
  AI_IDLE_TIMEOUT_MS,
  AiTimeoutError,
  createStreamWatchdog,
} from './watchdog'
export type { StreamWatchdog } from './watchdog'
