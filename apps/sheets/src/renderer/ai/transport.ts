import { translateServiceError } from '@wiswork/i18n'
import { createIpcTransport, type AgentTransport } from '@wiswork/agent-core'
import type { AiSettings } from '@wiswork/ai-provider'
import { getLang, t } from '../i18n/locale'

/** The shared IPC transport wired to the sheets preload bridge (window.desktopApi). */
export function createElectronTransport(getSettings: () => AiSettings): AgentTransport {
  return createIpcTransport<AiSettings>({
    onStream: (listener) => window.desktopApi.onAiStream(listener),
    start: (request) => window.desktopApi.aiStream(request),
    cancel: (requestId) => void window.desktopApi.aiStreamCancel(requestId),
    getSettings,
    unknownErrorText: () => t('aiUnknownError'),
    timeoutErrorText: () => t('aiTimeoutError'),
    creditsErrorText: () => t('aiCreditsExhausted'),
    serviceErrorText: (code) => translateServiceError(getLang(), code),
  })
}
