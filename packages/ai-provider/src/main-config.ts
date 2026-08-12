import { WISWORK_DEFAULT_MODEL, defaultAiSettings } from './providers'
import type { AiProviderConfig, AiProviderId, AiSettings } from './types'

export type WisworkRequestErrorCode = 'auth_required'

export type WisworkMainRequest =
  | { ok: true; provider: 'wiswork'; config: AiProviderConfig }
  | { ok: false; errorCode: WisworkRequestErrorCode }

/**
 * Main-process-only boundary for WisUsage requests. The renderer cannot supply or
 * override the login credential, endpoint, or model. The access token itself is
 * deliberately not accepted here; it remains inside the authenticated fetch callback.
 */
export function resolveWisworkMainRequest(
  hasAccessToken: boolean,
  rendererConfig: AiProviderConfig | undefined,
): WisworkMainRequest {
  if (!hasAccessToken) return { ok: false, errorCode: 'auth_required' }
  return {
    ok: true,
    provider: 'wiswork',
    config: {
      apiKey: '',
      model: WISWORK_DEFAULT_MODEL,
    },
  }
}

/**
 * Return renderer/persistence-safe settings: WisWork is authoritative and its
 * credential and endpoint fields are always blank, even if supplied by IPC.
 */
export function sanitizeWisworkSettings(settings: AiSettings): AiSettings {
  const defaults = defaultAiSettings()
  for (const id of Object.keys(defaults.providers) as AiProviderId[]) {
    defaults.providers[id] = {
      apiKey: '',
      model:
        id === 'wiswork'
          ? WISWORK_DEFAULT_MODEL
          : settings.providers?.[id]?.model || defaults.providers[id].model,
      baseUrl: undefined,
    }
  }
  return defaults
}
