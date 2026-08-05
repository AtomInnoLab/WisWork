import { translateServiceError } from '@wiswork/i18n'
import { createIpcTransport, type AgentTransport } from '@wiswork/agent-core'
import type { AiSettings } from '../../shared/ipc'
import { getLang, t } from '../i18n/locale'

/** The shared IPC transport wired to the docs preload bridge (window.desktop). */
export function createElectronTransport(getSettings: () => AiSettings): AgentTransport {
  return createIpcTransport<AiSettings>({
    onStream: (listener) => window.desktop.onAiStream(listener),
    start: (request) => window.desktop.aiStream(request),
    cancel: (requestId) => void window.desktop.aiStreamCancel(requestId),
    getSettings,
    unknownErrorText: () => t('aiUnknownError'),
    timeoutErrorText: () => t('aiTimeoutError'),
    creditsErrorText: () => t('aiCreditsExhausted'),
    serviceErrorText: (code) => translateServiceError(getLang(), code),
  })
}
