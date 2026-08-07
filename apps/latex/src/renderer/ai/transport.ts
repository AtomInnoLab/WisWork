import { createIpcTransport, type AgentTransport } from '@wiswork/agent-core'
import { defaultAiSettings, type AiSettings } from '@wiswork/ai-provider'

/** Renderer settings are intentionally non-authoritative: main owns endpoint, headers, model and key. */
export function fixedWisworkSettings(): AiSettings {
  const settings = defaultAiSettings()
  settings.provider = 'wiswork'
  settings.providers.wiswork = {
    ...settings.providers.wiswork,
    apiKey: '',
    baseUrl: undefined,
  }
  return settings
}

export function createLatexTransport(): AgentTransport {
  return createIpcTransport<AiSettings>({
    onStream: (listener) => window.latexApi.onAiStream(listener),
    start: (request) => window.latexApi.aiStream({ ...request, settings: fixedWisworkSettings() }),
    cancel: (requestId) => void window.latexApi.aiStreamCancel(requestId),
    getSettings: fixedWisworkSettings,
    unknownErrorText: () => 'The WisWork model service is unavailable.',
    timeoutErrorText: () => 'The WisWork model request timed out.',
  })
}
