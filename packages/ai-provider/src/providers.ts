import type { AiProviderId, AiProviderMeta, AiSettings, LegacyAiSettings } from './types'

export const WISWORK_MESSAGES_URL = 'https://wisusage.dev.atominnolab.com/v1/messages'

export const WISWORK_DEFAULT_MODEL = 'openai/gpt-5.6-sol'

/** WisUsage requires an explicit serving region on every managed request. */
export const WISWORK_REQUEST_LOCATION = 'sg'

export const AI_PROVIDERS: AiProviderMeta[] = [
  {
    id: 'wiswork',
    label: 'WisWork',
    models: [WISWORK_DEFAULT_MODEL],
    defaultModel: WISWORK_DEFAULT_MODEL,
    keyPlaceholder: 'Uses your WisWork login',
  },
  {
    id: 'anthropic',
    label: 'Claude',
    models: [
      'claude-sonnet-5',
      'claude-opus-4-8',
      'claude-opus-4-7',
      'claude-sonnet-4-6',
      'claude-opus-4-6',
      'claude-opus-4-5-20251101',
      'claude-haiku-4-5-20251001',
      'claude-sonnet-4-5-20250929',
    ],
    defaultModel: 'claude-opus-4-7',
    keyPlaceholder: 'sk-ant-api03-...',
  },
  {
    id: 'gemini',
    label: 'Gemini',
    models: ['gemini-2.5-pro', 'gemini-2.5-flash', 'gemini-2.0-flash'],
    defaultModel: 'gemini-2.5-flash',
    keyPlaceholder: 'AIza...',
  },
  {
    id: 'deepseek',
    label: 'DeepSeek',
    models: ['deepseek-chat', 'deepseek-reasoner'],
    defaultModel: 'deepseek-chat',
    keyPlaceholder: 'sk-...',
  },
  {
    id: 'openai',
    label: 'OpenAI',
    models: ['gpt-4.1', 'gpt-4.1-mini', 'gpt-4o', 'gpt-4o-mini'],
    defaultModel: 'gpt-4.1-mini',
    keyPlaceholder: 'sk-...',
  },
  {
    id: 'custom',
    label: 'Custom',
    models: [],
    defaultModel: '',
    keyPlaceholder: 'API Key',
    needsBaseUrl: true,
  },
]

/**
 * Fresh settings with every provider's default model and an empty key,
 * except providers listed in `defaultApiKeys` (e.g. an app-specific
 * preconfigured Anthropic key). Callers own that policy; this package
 * has no hardcoded keys.
 */
export function defaultAiSettings(
  defaultApiKeys?: Partial<Record<AiProviderId, string>>,
): AiSettings {
  const providers = {} as AiSettings['providers']
  for (const meta of AI_PROVIDERS) {
    providers[meta.id] = {
      apiKey: meta.id === 'wiswork' ? '' : (defaultApiKeys?.[meta.id] ?? ''),
      model: meta.defaultModel,
      baseUrl: meta.needsBaseUrl ? '' : undefined,
    }
  }
  return { provider: 'wiswork', providers }
}

/**
 * Merge on-disk settings over freshly computed defaults, migrating the
 * pre-provider shape (a single OpenAI-compatible endpoint) into the
 * "custom" provider slot. `stored` is whatever the caller read from its
 * settings file (already JSON-parsed); this function does no file I/O.
 */
export function resolveAiSettings(
  stored: Partial<AiSettings> & LegacyAiSettings,
  defaults: AiSettings,
): AiSettings {
  if (!stored.providers) {
    if (stored.apiKey) {
      defaults.providers.custom = {
        apiKey: stored.apiKey,
        model: stored.model ?? '',
        baseUrl: stored.baseUrl ?? 'https://api.openai.com/v1',
      }
    }
    return defaults
  }
  const knownIds = new Set(AI_PROVIDERS.map((provider) => provider.id))
  const providers = { ...defaults.providers }
  for (const [id, config] of Object.entries(stored.providers)) {
    if (id === 'wiswork') continue
    if (knownIds.has(id as AiProviderId)) providers[id as AiProviderId] = config
  }
  return {
    provider:
      stored.provider && knownIds.has(stored.provider) ? stored.provider : defaults.provider,
    providers,
  }
}
