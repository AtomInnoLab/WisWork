import { WISWORK_DEFAULT_MODEL, defaultAiSettings } from './providers'
import type { AiProviderConfig, AiProviderId, AiSettings } from './types'

export type WisworkRequestErrorCode = 'auth_required' | 'model_credentials_missing'

export type WisworkMainRequest =
  | { ok: true; provider: 'wiswork'; config: AiProviderConfig }
  | { ok: false; errorCode: WisworkRequestErrorCode }

/**
 * Main-process-only boundary for WisModel requests. The renderer cannot supply or
 * override the service credential, endpoint, or model.
 */
export function resolveWisworkMainRequest(
  loggedIn: boolean,
  rendererConfig: AiProviderConfig | undefined,
  env: Readonly<Record<string, string | undefined>> = process.env,
): WisworkMainRequest {
  if (!loggedIn) return { ok: false, errorCode: 'auth_required' }
  const apiKey = env.WISWORK_MODEL_API_KEY?.trim()
  if (!apiKey) return { ok: false, errorCode: 'model_credentials_missing' }
  return {
    ok: true,
    provider: 'wiswork',
    config: {
      apiKey,
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
