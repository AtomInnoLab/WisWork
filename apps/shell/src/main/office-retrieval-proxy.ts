const MAX_RESPONSE_BYTES = 512 * 1024
const MAX_QUERY_CHARS = 4_096
const MAX_FETCH_CONTENT_CHARS = 256 * 1024
const MAX_RESULTS = 20
const REQUEST_TIMEOUT_MS = 15_000
// Intentionally empty until the service owner publishes both the canonical URL and this contract.
// The service—not this client—must resolve DNS safely on every connection, reject DNS rebinding,
// and validate every redirect hop before fetching. Runtime configuration cannot widen this map.
export interface OfficeRetrievalServiceAttestation {
  contract: 'wiswork-office-retrieval-v1'
  ssrfProtection: 'dns-rebinding-and-redirect-hops-v1'
}
export const OFFICE_RETRIEVAL_SERVICES: Readonly<
  Record<string, OfficeRetrievalServiceAttestation>
> = {}

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
  allowedServices: Readonly<
    Record<string, OfficeRetrievalServiceAttestation>
  > = OFFICE_RETRIEVAL_SERVICES,
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
    allowedServices[configured]?.contract !== 'wiswork-office-retrieval-v1' ||
    allowedServices[configured]?.ssrfProtection !== 'dns-rebinding-and-redirect-hops-v1'
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
  if (hostname === 'localhost' || unsafeIpLiteral(hostname))
    throw new Error('retrieval_invalid_request')
  return raw
}

function unsafeIpLiteral(hostname: string): boolean {
  const kind = isIP(hostname)
  if (kind === 4) {
    const [a, b, c] = hostname.split('.').map(Number)
    return (
      a === 0 ||
      a === 10 ||
      a === 127 ||
      (a === 100 && b >= 64 && b <= 127) ||
      (a === 169 && b === 254) ||
      (a === 172 && b >= 16 && b <= 31) ||
      (a === 192 && b === 0 && c === 0) ||
      (a === 192 && b === 0 && c === 2) ||
      (a === 192 && b === 88 && c === 99) ||
      (a === 192 && b === 168) ||
      (a === 198 && (b === 18 || b === 19)) ||
      (a === 198 && b === 51 && c === 100) ||
      (a === 203 && b === 0 && c === 113) ||
      a >= 224
    )
  }
  if (kind !== 6) return false
  const groups = expandIpv6(hostname)
  if (!groups) return true
  const [first, second] = groups
  const mapped = groups.slice(0, 5).every((value) => value === 0) && groups[5] === 0xffff
  return (
    groups.every((value) => value === 0) ||
    (groups.slice(0, 7).every((value) => value === 0) && groups[7] === 1) ||
    first === 0 ||
    mapped ||
    (first & 0xfe00) === 0xfc00 ||
    (first & 0xffc0) === 0xfe80 ||
    (first & 0xffc0) === 0xfec0 ||
    (first & 0xff00) === 0xff00 ||
    (first === 0x2001 && (second & 0xfe00) === 0) ||
    (first === 0x2001 && second === 0x0db8) ||
    first === 0x2002 ||
    (first & 0xfff0) === 0x3ff0 ||
    (first === 0x0064 && second === 0xff9b) ||
    (first === 0x0100 && second === 0)
  )
}

function expandIpv6(value: string): number[] | null {
  const halves = value.split('::')
  if (halves.length > 2) return null
  const left = halves[0] ? halves[0].split(':') : []
  const right = halves.length === 2 && halves[1] ? halves[1].split(':') : []
  const missing = 8 - left.length - right.length
  if ((halves.length === 1 && missing !== 0) || (halves.length === 2 && missing < 1)) return null
  const raw = [...left, ...Array.from({ length: missing }, () => '0'), ...right]
  if (raw.length !== 8 || raw.some((group) => !/^[0-9a-f]{1,4}$/i.test(group))) return null
  return raw.map((group) => Number.parseInt(group, 16))
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
      if (signal?.aborted) throw new Error('retrieval_cancelled', { cause: error })
      throw new Error('retrieval_upstream_error', { cause: error })
    } finally {
      active -= 1
      clearTimeout(timer)
      signal?.removeEventListener('abort', abort)
    }
  }
}
import { isIP } from 'node:net'
