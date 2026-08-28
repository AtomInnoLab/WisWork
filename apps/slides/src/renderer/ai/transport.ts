import { translateServiceError } from '@wiswork/i18n'
import { createIpcTransport, type AgentTransport } from '@wiswork/agent-core'
import type { AiSettings } from '../../shared/ipc'
import { getLang, t } from '../i18n/locale'

/** The shared IPC transport wired to the slides preload bridge (window.slidesApi). */
export function createElectronTransport(
  getSettings: () => AiSettings,
  limits?: { maxSerializedRequestBytes: number },
): AgentTransport {
  return createIpcTransport<AiSettings>({
    onStream: (listener) => window.slidesApi.onAiStream(listener),
    start: (request) => {
      if (
        limits &&
        new TextEncoder().encode(JSON.stringify(request)).byteLength >
          limits.maxSerializedRequestBytes
      ) {
        throw new Error('quality_request_too_large')
      }
      return window.slidesApi.aiStream(request)
    },
    cancel: (requestId) => void window.slidesApi.aiStreamCancel(requestId),
    getSettings,
    unknownErrorText: () => t('aiErrUnknown'),
    timeoutErrorText: () => t('aiErrStreamTimeout'),
    creditsErrorText: () => t('aiCreditsExhausted'),
    serviceErrorText: (code) => translateServiceError(getLang(), code),
  })
}
