import { AiProviderError, isAuthRequiredError, safeHttpProviderError } from './errors'
import { aiFetch } from './fetch'
import { httpBodyDetail } from './http-error'
import { WISWORK_DEFAULT_MODEL, WISWORK_MESSAGES_URL, WISWORK_REQUEST_LOCATION } from './providers'
import type { AiChatResponse, AiProviderConfig, AiProviderId, WisworkFetchWithAuth } from './types'
import { AI_CHAT_RESPONSE_TIMEOUT_MS, createStreamWatchdog, type StreamWatchdog } from './watchdog'

async function chatAnthropic(
  wd: StreamWatchdog,
  config: AiProviderConfig,
  system: string,
  user: string,
  baseUrl = 'https://api.anthropic.com',
): Promise<AiChatResponse> {
  const response = await aiFetch(`${baseUrl.replace(/\/$/, '')}/v1/messages`, {
    method: 'POST',
    signal: wd.signal,
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': config.apiKey,
      'anthropic-version': '2023-06-01',
      // Fetch in the Electron main process goes through Chromium's network stack; this header avoids 403.
      'anthropic-dangerous-direct-browser-access': 'true',
    },
    body: JSON.stringify({
      model: config.model,
      max_tokens: 8192,
      system,
      messages: [{ role: 'user', content: user }],
    }),
  })
  wd.touch()
  if (!response.ok) {
    return {
      ok: false,
      error: `Claude HTTP ${response.status}: ${httpBodyDetail(await response.text())}`,
    }
  }
  const json = (await response.json()) as { content?: Array<{ type: string; text?: string }> }
  const content = json.content
    ?.filter((c) => c.type === 'text')
    .map((c) => c.text ?? '')
    .join('')
  if (!content) return { ok: false, error: 'Claude returned an empty response' }
  return { ok: true, content }
}

async function chatWiswork(
  wd: StreamWatchdog,
  system: string,
  user: string,
  fetchWithAuth?: WisworkFetchWithAuth,
): Promise<AiChatResponse> {
  if (!fetchWithAuth) return { ok: false, error: 'auth_required', errorCode: 'auth_required' }
  let response: Response
  try {
    response = await fetchWithAuth((accessToken) =>
      fetch(WISWORK_MESSAGES_URL, {
        method: 'POST',
        signal: wd.signal,
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
          'x-req-location': WISWORK_REQUEST_LOCATION,
        },
        body: JSON.stringify({
          model: WISWORK_DEFAULT_MODEL,
          max_tokens: 8192,
          system,
          messages: [{ role: 'user', content: user }],
        }),
      }),
    )
  } catch (error) {
    if (wd.signal.aborted) throw error
    if (isAuthRequiredError(error)) throw new AiProviderError('auth_required', undefined, 'auth')
    throw new AiProviderError('model_upstream_unavailable', undefined, 'request')
  }
  wd.touch()
  if (!response.ok) {
    const errorCode =
      response.status === 401 ? 'auth_required' : safeHttpProviderError(response.status).code
    return {
      ok: false,
      error: errorCode,
      errorCode,
      diagnostic: { stage: 'response', httpStatus: response.status },
    }
  }
  let json: { content?: Array<{ type: string; text?: string }> }
  try {
    json = (await response.json()) as typeof json
  } catch {
    throw new AiProviderError('model_invalid_response', response.status, 'response')
  }
  const content = json.content
    ?.filter((block) => block.type === 'text')
    .map((block) => block.text ?? '')
    .join('')
  if (!content)
    return { ok: false, error: 'model_invalid_response', errorCode: 'model_invalid_response' }
  return { ok: true, content }
}

async function chatGemini(
  wd: StreamWatchdog,
  config: AiProviderConfig,
  system: string,
  user: string,
  baseUrl = 'https://generativelanguage.googleapis.com/v1beta',
): Promise<AiChatResponse> {
  const url = `${baseUrl.replace(/\/$/, '')}/models/${config.model}:generateContent`
  const response = await aiFetch(url, {
    method: 'POST',
    signal: wd.signal,
    headers: {
      'Content-Type': 'application/json',
      'x-goog-api-key': config.apiKey,
    },
    body: JSON.stringify({
      systemInstruction: { parts: [{ text: system }] },
      contents: [{ role: 'user', parts: [{ text: user }] }],
      generationConfig: { temperature: 0.3 },
    }),
  })
  wd.touch()
  if (!response.ok) {
    return {
      ok: false,
      error: `Gemini HTTP ${response.status}: ${httpBodyDetail(await response.text())}`,
    }
  }
  const json = (await response.json()) as {
    candidates?: Array<{ content?: { parts?: Array<{ text?: string }> } }>
  }
  const content = json.candidates?.[0]?.content?.parts?.map((p) => p.text ?? '').join('')
  if (!content) return { ok: false, error: 'Gemini returned an empty response' }
  return { ok: true, content }
}

async function chatOpenAiCompatible(
  wd: StreamWatchdog,
  baseUrl: string,
  config: AiProviderConfig,
  system: string,
  user: string,
  safeErrors = false,
): Promise<AiChatResponse> {
  const response = await aiFetch(`${baseUrl.replace(/\/$/, '')}/chat/completions`, {
    method: 'POST',
    signal: wd.signal,
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${config.apiKey}`,
    },
    body: JSON.stringify({
      model: config.model,
      messages: [
        { role: 'system', content: system },
        { role: 'user', content: user },
      ],
      temperature: 0.3,
    }),
  })
  wd.touch()
  if (!response.ok) {
    if (safeErrors) {
      const error = safeHttpProviderError(response.status)
      return { ok: false, error: error.code, errorCode: error.code }
    }
    return { ok: false, error: `HTTP ${response.status}: ${httpBodyDetail(await response.text())}` }
  }
  const json = (await response.json()) as { choices?: Array<{ message?: { content?: string } }> }
  const content = json.choices?.[0]?.message?.content
  if (!content) return { ok: false, error: 'AI returned an empty response' }
  return { ok: true, content }
}

const OPENAI_COMPATIBLE_BASE_URLS: Partial<Record<AiProviderId, string>> = {
  deepseek: 'https://api.deepseek.com/v1',
  openai: 'https://api.openai.com/v1',
}

/** route a one-shot (non-streaming, non-tool-calling) chat call by provider id */
export async function chatForProvider(
  provider: AiProviderId,
  config: AiProviderConfig,
  system: string,
  user: string,
  signal?: AbortSignal,
  fetchWithAuth?: WisworkFetchWithAuth,
): Promise<AiChatResponse> {
  // non-streaming: the server generates the full answer before the headers arrive,
  // so the connect phase gets the long budget; the body read then gets the idle budget
  const wd = createStreamWatchdog(signal, AI_CHAT_RESPONSE_TIMEOUT_MS)
  return wd.guard(() => {
    switch (provider) {
      case 'wiswork':
        return chatWiswork(wd, system, user, fetchWithAuth)
      case 'anthropic':
        return chatAnthropic(wd, config, system, user)
      case 'gemini':
        return chatGemini(wd, config, system, user)
      case 'deepseek':
      case 'openai':
        return chatOpenAiCompatible(
          wd,
          OPENAI_COMPATIBLE_BASE_URLS[provider]!,
          config,
          system,
          user,
        )
      case 'custom':
        if (!config.baseUrl)
          return Promise.resolve({
            ok: false as const,
            error: 'A custom provider requires a Base URL',
          })
        return chatOpenAiCompatible(wd, config.baseUrl, config, system, user)
      default:
        return Promise.resolve({ ok: false as const, error: `Unknown provider: ${provider}` })
    }
  })
}
