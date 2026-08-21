const MAX_RESPONSE_BYTES = 512 * 1024
const MAX_QUERY_CHARS = 4_096
const MAX_FETCH_CONTENT_CHARS = 256 * 1024
const MAX_RESULTS = 20
const REQUEST_TIMEOUT_MS = 15_000
// Intentionally empty until the retrieval service owner publishes the canonical production URL.
// Adding an endpoint is a reviewed build-time change; runtime configuration cannot widen it.
export const OFFICE_RETRIEVAL_ENDPOINTS: readonly string[] = []

export type OfficeWebCapability = 'web-search.v1' | 'web-fetch.v1' | 'image-search.v1'
export type OfficeRetrievalProxy = (
  capability: string,
  body: unknown,
  signal?: AbortSignal,
) => Promise<Uint8Array>

const record = (value: unknown): Record<string, unknown> => {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new Error('retrieval_invalid_request')
  return value as Record<string, unknown>
}
const exact = (value: Record<string, unknown>, keys: readonly string[]) => {
  if (
    Object.keys(value).length !== keys.length ||
    keys.some((key) => !Object.prototype.hasOwnProperty.call(value, key))
  )
    throw new Error('retrieval_invalid_request')
}
const boundedString = (value: unknown, maximum: number): string => {
  if (typeof value !== 'string' || value.length < 1 || value.length > maximum)
    throw new Error('retrieval_invalid_request')
  return value
}
const maxResults = (value: unknown): number => {
  if (!Number.isSafeInteger(value) || Number(value) < 1 || Number(value) > MAX_RESULTS)
    throw new Error('retrieval_invalid_request')
  return Number(value)
}

export function officeRetrievalEndpointFromEnv(
  env: Record<string, string | undefined>,
  allowedEndpoints: readonly string[] = OFFICE_RETRIEVAL_ENDPOINTS,
): string | null {
  const configured = env.WISWORK_OFFICE_RETRIEVAL_URL
  if (!configured) return null
  let url: URL
  try {
    url = new URL(configured)
  } catch {
    throw new Error('invalid_office_retrieval_url')
  }
  if (
    url.protocol !== 'https:' ||
    url.username ||
    url.password ||
    url.hash ||
    url.href !== configured ||
    !allowedEndpoints.includes(configured)
  )
    throw new Error('invalid_office_retrieval_url')
  return configured
}

function safeHttpsUrl(value: unknown): string {
  const raw = boundedString(value, 2_048)
  let url: URL
  try {
    url = new URL(raw)
  } catch {
    throw new Error('retrieval_invalid_request')
  }
  if (url.protocol !== 'https:' || url.username || url.password || url.href !== raw)
    throw new Error('retrieval_invalid_request')
  const hostname = url.hostname.replace(/^\[|\]$/g, '').toLowerCase()
  if (
    hostname === 'localhost' ||
    hostname === '::1' ||
    hostname.startsWith('127.') ||
    hostname.startsWith('10.') ||
    hostname.startsWith('192.168.') ||
    /^169\.254\./.test(hostname) ||
    /^172\.(1[6-9]|2\d|3[01])\./.test(hostname) ||
    /^0\./.test(hostname) ||
    hostname === '0.0.0.0'
  )
    throw new Error('retrieval_invalid_request')
  return raw
}

function requestFor(capability: string, input: unknown) {
  const body = record(input)
  if (capability === 'web-search.v1' || capability === 'image-search.v1') {
    exact(body, ['query', 'max_results'])
    return {
      version: 1,
      operation: capability === 'web-search.v1' ? 'web-search' : 'image-search',
      input: {
        query: boundedString(body.query, MAX_QUERY_CHARS),
        max_results: maxResults(body.max_results),
      },
    }
  }
  if (capability === 'web-fetch.v1') {
    exact(body, ['url'])
    return { version: 1, operation: 'web-fetch', input: { url: safeHttpsUrl(body.url) } }
  }
  throw new Error('retrieval_invalid_request')
}

async function boundedJson(response: Response): Promise<unknown> {
  if (
    !response.ok ||
    response.redirected ||
    response.headers.get('content-type')?.split(';', 1)[0]?.trim().toLowerCase() !==
      'application/json' ||
    Number(response.headers.get('content-length') ?? 0) > MAX_RESPONSE_BYTES
  )
    throw new Error('retrieval_upstream_error')
  const reader = response.body?.getReader()
  if (!reader) throw new Error('retrieval_upstream_error')
  const chunks: Uint8Array[] = []
  let total = 0
  try {
    while (true) {
      const next = await reader.read()
      if (next.done) break
      total += next.value.byteLength
      if (total > MAX_RESPONSE_BYTES) throw new Error('retrieval_upstream_error')
      chunks.push(next.value)
    }
  } finally {
    reader.releaseLock()
  }
  const bytes = new Uint8Array(total)
  let offset = 0
  for (const chunk of chunks) {
    bytes.set(chunk, offset)
    offset += chunk.byteLength
  }
  try {
    return JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes))
  } catch {
    throw new Error('retrieval_upstream_error')
  }
}

const optionalString = (value: unknown, maximum: number): string | undefined =>
  value === undefined ? undefined : boundedString(value, maximum)

function sanitize(capability: string, raw: unknown): Record<string, unknown> {
  const value = record(raw)
  try {
    if (capability === 'web-search.v1') {
      const allowed = value.answer === undefined ? ['results'] : ['results', 'answer']
      exact(value, allowed)
      if (!Array.isArray(value.results) || value.results.length > MAX_RESULTS)
        throw new Error('invalid')
      return {
        results: value.results.map((item) => {
          const entry = record(item)
          exact(entry, ['title', 'url', 'snippet'])
          return {
            title: boundedString(entry.title, 512),
            url: safeHttpsUrl(entry.url),
            snippet: boundedString(entry.snippet, 4_096),
          }
        }),
        ...(value.answer === undefined
          ? {}
          : { answer: optionalString(value.answer, 8_192) as string }),
      }
    }
    if (capability === 'web-fetch.v1') {
      const allowed =
        value.title === undefined
          ? ['url', 'content', 'content_type']
          : ['url', 'title', 'content', 'content_type']
      exact(value, allowed)
      return {
        url: safeHttpsUrl(value.url),
        ...(value.title === undefined ? {} : { title: boundedString(value.title, 512) }),
        content: boundedString(value.content, MAX_FETCH_CONTENT_CHARS),
        content_type: boundedString(value.content_type, 128),
      }
    }
    if (capability === 'image-search.v1') {
      exact(value, ['images'])
      if (!Array.isArray(value.images) || value.images.length > MAX_RESULTS)
        throw new Error('invalid')
      return {
        images: value.images.map((item) => {
          const entry = record(item)
          exact(entry, ['title', 'image_url', 'source_url', 'source'])
          return {
            title: boundedString(entry.title, 512),
            image_url: safeHttpsUrl(entry.image_url),
            source_url: safeHttpsUrl(entry.source_url),
            source: boundedString(entry.source, 512),
          }
        }),
      }
    }
  } catch {
    throw new Error('retrieval_upstream_error')
  }
  throw new Error('retrieval_upstream_error')
}

export function createOfficeRetrievalProxy(options: {
  endpoint: string
  fetchWithAuth(request: (accessToken: string) => Promise<Response>): Promise<Response>
  fetch?: typeof fetch
  timeoutMs?: number
  maxConcurrent?: number
}): OfficeRetrievalProxy {
  const doFetch = options.fetch ?? fetch
  const maxConcurrent = options.maxConcurrent ?? 4
  if (!Number.isSafeInteger(maxConcurrent) || maxConcurrent < 1 || maxConcurrent > 16)
    throw new Error('invalid_office_retrieval_config')
  let active = 0
  return async (capability, input, signal) => {
    const request = requestFor(capability, input)
    if (active >= maxConcurrent) throw new Error('retrieval_busy')
    active += 1
    const controller = new AbortController()
    const abort = () => controller.abort()
    signal?.addEventListener('abort', abort, { once: true })
    const timer = setTimeout(abort, options.timeoutMs ?? REQUEST_TIMEOUT_MS)
    try {
      const response = await options.fetchWithAuth((accessToken) =>
        doFetch(options.endpoint, {
          method: 'POST',
          redirect: 'error',
          headers: {
            authorization: `Bearer ${accessToken}`,
            'content-type': 'application/json',
          },
          body: JSON.stringify(request),
          signal: controller.signal,
        }),
      )
      if (response.status === 401 || response.status === 403) throw new Error('auth_required')
      const output = sanitize(capability, await boundedJson(response))
      return new TextEncoder().encode(JSON.stringify(output))
    } catch (error) {
      if (error instanceof Error && error.message === 'auth_required') throw error
      if (signal?.aborted) throw new Error('retrieval_cancelled')
      throw new Error('retrieval_upstream_error')
    } finally {
      active -= 1
      clearTimeout(timer)
      signal?.removeEventListener('abort', abort)
    }
  }
}
