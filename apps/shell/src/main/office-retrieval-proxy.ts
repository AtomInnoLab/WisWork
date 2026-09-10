import { imageSearch, wisUsageWebSearch } from '@wiswork/ai-search'
import type { LookupAddress } from 'node:dns'
import { lookup } from 'node:dns/promises'
import { request as httpsRequest } from 'node:https'
import type { LookupFunction, TcpNetConnectOpts } from 'node:net'
import { MAX_OFFICE_IMAGE_SOURCE_BYTES } from './office-image-handoff'
const MAX_RESPONSE_BYTES = 512 * 1024
const MAX_QUERY_CHARS = 4_096
const MAX_FETCH_CONTENT_CHARS = 256 * 1024
const MAX_RESULTS = 20
const REQUEST_TIMEOUT_MS = 15_000
const OFFICE_IMAGE_FETCH_ENDPOINT = 'https://office.8-216-134-194.sslip.io/office-image-fetch'
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

export type OfficeWebCapability =
  'web-search.v1' | 'web-fetch.v1' | 'image-search.v1' | 'image-fetch.v1'
export interface OfficeRetrievalProxy {
  (capability: string, body: unknown, signal?: AbortSignal): Promise<Uint8Array>
  clear?(): void
}

export interface DownloadedImage {
  mime: 'image/png' | 'image/jpeg'
  bytes: Uint8Array
}

function supportedImageMime(value: unknown): value is DownloadedImage['mime'] {
  return value === 'image/png' || value === 'image/jpeg'
}

function boundedImageDimension(value: unknown): number | undefined {
  return Number.isSafeInteger(value) && Number(value) >= 1 && Number(value) <= 100_000
    ? Number(value)
    : undefined
}

export async function collectBoundedImageBytes(
  chunks: AsyncIterable<Uint8Array>,
  maximum = 2 * 1024 * 1024,
): Promise<Uint8Array> {
  const collected: Uint8Array[] = []
  let total = 0
  for await (const chunk of chunks) {
    total += chunk.byteLength
    if (total > maximum) throw new Error('image_limit')
    collected.push(chunk)
  }
  const bytes = new Uint8Array(total)
  let offset = 0
  for (const chunk of collected) {
    bytes.set(chunk, offset)
    offset += chunk.byteLength
  }
  return bytes
}

async function* responseChunks(stream: ReadableStream<Uint8Array>) {
  const reader = stream.getReader()
  let complete = false
  try {
    while (true) {
      const next = await reader.read()
      if (next.done) {
        complete = true
        return
      }
      yield next.value
    }
  } finally {
    if (!complete) await reader.cancel().catch(() => undefined)
    reader.releaseLock()
  }
}

export function createPinnedLookup(
  selected: LookupAddress | readonly LookupAddress[],
): LookupFunction {
  const addresses = ('address' in selected ? [selected] : selected).map(({ address, family }) => ({
    address,
    family,
  }))
  const first = addresses[0]
  if (!first) throw new Error('retrieval_upstream_error')
  return ((_hostname, options, callback) => {
    if (typeof options === 'object' && options.all) callback(null, addresses)
    else callback(null, first.address, first.family)
  }) as LookupFunction
}

type LookupAddresses = (hostname: string) => Promise<readonly LookupAddress[]>

export function resolvePublicImageRedirect(
  currentUrl: string,
  location: string | undefined,
  redirectsRemaining: number,
): string {
  if (!location || redirectsRemaining < 1) throw new Error('retrieval_upstream_error')
  try {
    return safeHttpsUrl(new URL(location, currentUrl).href)
  } catch (error) {
    throw new Error('retrieval_upstream_error', { cause: error })
  }
}

async function downloadPublicImage(
  url: string,
  signal?: AbortSignal,
  lookupAddresses: LookupAddresses = (hostname) => lookup(hostname, { all: true, verbatim: true }),
  timeoutMs = REQUEST_TIMEOUT_MS,
  redirectsRemaining = 3,
  maximumBytes = 2 * 1024 * 1024,
): Promise<DownloadedImage> {
  const startedAt = Date.now()
  const parsed = new URL(url)
  if (signal?.aborted) throw new Error('retrieval_upstream_error')
  let lookupTimer: ReturnType<typeof setTimeout> | undefined
  let abortLookup: (() => void) | undefined
  const addresses = await Promise.race([
    lookupAddresses(parsed.hostname),
    new Promise<never>((_resolve, reject) => {
      lookupTimer = setTimeout(() => reject(new Error('retrieval_upstream_error')), timeoutMs)
    }),
    new Promise<never>((_resolve, reject) => {
      abortLookup = () => reject(new Error('retrieval_upstream_error'))
      signal?.addEventListener('abort', abortLookup, { once: true })
    }),
  ]).finally(() => {
    clearTimeout(lookupTimer)
    if (abortLookup) signal?.removeEventListener('abort', abortLookup)
  })
  if (signal?.aborted) throw new Error('retrieval_upstream_error')
  if (!addresses.length || addresses.some((entry) => unsafeIpLiteral(entry.address)))
    throw new Error('retrieval_upstream_error')
  return new Promise((resolve, reject) => {
    let settled = false
    const fail = (
      code:
        | 'retrieval_upstream_error'
        | 'image_mime_unsupported'
        | 'image_limit' = 'retrieval_upstream_error',
    ) => {
      if (settled) return
      settled = true
      reject(new Error(code))
    }
    // HTTPS forwards socket options that its RequestOptions type does not declare.
    const connectionOptions: Pick<TcpNetConnectOpts, 'autoSelectFamily'> = {
      autoSelectFamily: true,
    }
    const request = httpsRequest(
      parsed,
      {
        method: 'GET',
        agent: false,
        lookup: createPinnedLookup(addresses),
        ...connectionOptions,
        headers: {
          Accept: 'image/png,image/jpeg',
          'User-Agent':
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/128 Safari/537.36',
        },
      },
      (response) => {
        if (
          response.statusCode !== undefined &&
          response.statusCode >= 300 &&
          response.statusCode < 400
        ) {
          let next: string
          try {
            next = resolvePublicImageRedirect(
              parsed.href,
              response.headers.location,
              redirectsRemaining,
            )
          } catch {
            response.destroy()
            fail()
            return
          }
          if (settled) return
          settled = true
          response.destroy()
          resolve(
            downloadPublicImage(
              next,
              signal,
              lookupAddresses,
              timeoutMs,
              redirectsRemaining - 1,
              maximumBytes,
            ),
          )
          return
        }
        const mime = response.headers['content-type']?.split(';', 1)[0]?.trim().toLowerCase()
        const declared = Number(response.headers['content-length'] ?? 0)
        if (response.statusCode !== 200) {
          response.destroy()
          fail()
          return
        }
        if (declared > maximumBytes) {
          response.destroy()
          fail('image_limit')
          return
        }
        if (!supportedImageMime(mime)) {
          response.destroy()
          fail('image_mime_unsupported')
          return
        }
        void collectBoundedImageBytes(response, maximumBytes)
          .then((bytes) => {
            if (settled) return
            settled = true
            resolve({ mime, bytes })
          })
          .catch((error) => {
            response.destroy()
            fail(
              error instanceof Error && error.message === 'image_limit'
                ? 'image_limit'
                : 'retrieval_upstream_error',
            )
          })
      },
    )
    request.once('error', () => fail())
    const timeout = setTimeout(
      () => request.destroy(new Error('retrieval_upstream_error')),
      Math.max(1, timeoutMs - (Date.now() - startedAt)),
    )
    const abort = () => request.destroy(new Error('retrieval_upstream_error'))
    signal?.addEventListener('abort', abort, { once: true })
    if (signal?.aborted) abort()
    request.once('close', () => {
      clearTimeout(timeout)
      signal?.removeEventListener('abort', abort)
    })
    request.end()
  })
}

export function createOfficeRemoteImageDownloader(options: {
  fetchWithAuth(request: (accessToken: string) => Promise<Response>): Promise<Response>
  fetch?: typeof fetch
  timeoutMs?: number
}): (url: string, signal?: AbortSignal) => Promise<DownloadedImage> {
  const doFetch = options.fetch ?? fetch
  return async (url, signal) => {
    const controller = new AbortController()
    const cancel = () => controller.abort()
    let rejectDeadline: ((reason: Error) => void) | undefined
    const deadline = new Promise<never>((_resolve, reject) => {
      rejectDeadline = reject
    })
    signal?.addEventListener('abort', cancel, { once: true })
    controller.signal.addEventListener(
      'abort',
      () => rejectDeadline?.(new Error('image_fetch_unavailable')),
      { once: true },
    )
    const timer = setTimeout(cancel, options.timeoutMs ?? REQUEST_TIMEOUT_MS)
    try {
      const response = await Promise.race([
        options.fetchWithAuth((accessToken) =>
          doFetch(OFFICE_IMAGE_FETCH_ENDPOINT, {
            method: 'POST',
            redirect: 'error',
            headers: {
              authorization: `Bearer ${accessToken}`,
              'content-type': 'application/json',
            },
            body: JSON.stringify({ url: safeHttpsUrl(url) }),
            signal: controller.signal,
          }),
        ),
        deadline,
      ])
      if (response.status === 413) throw new Error('image_limit')
      if (response.status === 415) throw new Error('image_mime_unsupported')
      if (response.status !== 200 || response.redirected) throw new Error('image_fetch_unavailable')
      const mime = response.headers.get('content-type')?.split(';', 1)[0]?.trim().toLowerCase()
      if (!supportedImageMime(mime)) throw new Error('image_mime_unsupported')
      if (Number(response.headers.get('content-length') ?? 0) > MAX_OFFICE_IMAGE_SOURCE_BYTES)
        throw new Error('image_limit')
      if (!response.body) throw new Error('image_fetch_unavailable')
      return {
        mime,
        bytes: await collectBoundedImageBytes(
          responseChunks(response.body),
          MAX_OFFICE_IMAGE_SOURCE_BYTES,
        ),
      }
    } catch (error) {
      if (signal?.aborted) throw new Error('search_cancelled', { cause: error })
      if (
        error instanceof Error &&
        (error.message === 'image_limit' || error.message === 'image_mime_unsupported')
      )
        throw error
      throw new Error('image_fetch_unavailable', { cause: error })
    } finally {
      clearTimeout(timer)
      signal?.removeEventListener('abort', cancel)
    }
  }
}

export function createOfficeLocalSearchProxy(options: {
  fetchWithAuth(request: (accessToken: string) => Promise<Response>): Promise<Response>
  webSearch?: typeof wisUsageWebSearch
  searchImages?: typeof imageSearch
  downloadImage?: (url: string, signal?: AbortSignal) => Promise<DownloadedImage>
  remoteDownloadImage?: (url: string, signal?: AbortSignal) => Promise<DownloadedImage>
  normalizeImage?: (image: DownloadedImage) => Promise<DownloadedImage>
  lookupAddresses?: LookupAddresses
  imageTimeoutMs?: number
}): OfficeRetrievalProxy {
  const searchWeb = options.webSearch ?? wisUsageWebSearch
  const searchImages = options.searchImages ?? imageSearch
  // Only a local normalizer may consume a larger source. Never send it over Relay.
  const maximumSourceBytes = options.normalizeImage
    ? MAX_OFFICE_IMAGE_SOURCE_BYTES
    : 2 * 1024 * 1024
  const downloadImage =
    options.downloadImage ??
    ((url: string, signal?: AbortSignal) =>
      downloadPublicImage(
        url,
        signal,
        options.lookupAddresses,
        options.imageTimeoutMs,
        3,
        maximumSourceBytes,
      ))
  const semanticImageError = (error: unknown) =>
    error instanceof Error &&
    (error.message === 'image_limit' ||
      error.message === 'image_mime_unsupported' ||
      error.message === 'search_cancelled')
  const allowedImages = new Map<
    string,
    { expiresAt: number; query: string; maxResults: number; fallbackImageUrl?: string }
  >()
  let generation = 0
  const proxy: OfficeRetrievalProxy = async (capability, body, signal) => {
    const requestGeneration = generation
    const checkCurrent = () => {
      if (signal?.aborted || requestGeneration !== generation) throw new Error('search_cancelled')
    }
    const request = requestFor(capability, body)
    checkCurrent()
    if (request.operation === 'web-search') {
      const input = request.input as { query: string; max_results: number }
      const result = await searchWeb(input.query, Math.min(input.max_results, 10), {
        fetchWithAuth: options.fetchWithAuth,
        signal,
      })
      checkCurrent()
      return new TextEncoder().encode(JSON.stringify({ results: result.results }))
    }
    if (request.operation === 'image-search') {
      const input = request.input as { query: string; max_results: number }
      const result = await searchImages(input.query, input.max_results)
      checkCurrent()
      const expiresAt = Date.now() + 15 * 60_000
      for (const image of result.images)
        allowedImages.set(image.imageUrl, {
          expiresAt,
          query: input.query,
          maxResults: input.max_results,
          ...(image.fallbackImageUrl
            ? { fallbackImageUrl: safeHttpsUrl(image.fallbackImageUrl) }
            : {}),
        })
      while (allowedImages.size > 100) allowedImages.delete(allowedImages.keys().next().value!)
      return new TextEncoder().encode(
        JSON.stringify({
          images: result.images.map((image) => {
            const width = boundedImageDimension(image.width)
            const height = boundedImageDimension(image.height)
            return {
              title: image.title,
              image_url: safeHttpsUrl(image.imageUrl),
              source_url: safeHttpsUrl(new URL(image.sourceUrl).href),
              source: image.source,
              ...(width === undefined ? {} : { width }),
              ...(height === undefined ? {} : { height }),
            }
          }),
        }),
      )
    }
    if (request.operation === 'image-fetch') {
      const url = (request.input as { url: string }).url
      const source = allowedImages.get(url)
      if (!source) throw new Error('retrieval_invalid_request')
      if (source.expiresAt < Date.now()) {
        // Re-run the original search; an expired candidate is not permission to
        // fetch a stale/arbitrary URL. Keep the same bounded 100-candidate ledger.
        const result = await searchImages(source.query, source.maxResults)
        checkCurrent()
        const renewed = result.images.find((image) => image.imageUrl === url)
        if (!renewed) throw new Error('retrieval_invalid_request')
        if (renewed.fallbackImageUrl)
          source.fallbackImageUrl = safeHttpsUrl(renewed.fallbackImageUrl)
        else delete source.fallbackImageUrl
        source.expiresAt = Date.now() + 15 * 60_000
      }
      const deadline = options.remoteDownloadImage ? new AbortController() : undefined
      const cancelDeadline = () => deadline?.abort()
      signal?.addEventListener('abort', cancelDeadline, { once: true })
      const deadlineTimer = deadline
        ? setTimeout(cancelDeadline, options.imageTimeoutMs ?? REQUEST_TIMEOUT_MS)
        : undefined
      const candidateSignal = deadline?.signal ?? signal
      const fetchCandidate = async (candidateUrl: string) => {
        if (!options.remoteDownloadImage) return downloadImage(candidateUrl, candidateSignal)
        try {
          return await options.remoteDownloadImage(candidateUrl, candidateSignal)
        } catch (error) {
          if (signal?.aborted) throw error
          if (deadline?.signal.aborted)
            throw new Error('retrieval_upstream_error', { cause: error })
          if (semanticImageError(error)) throw error
          if (!(error instanceof Error) || error.message !== 'image_fetch_unavailable') throw error
          return downloadImage(candidateUrl, candidateSignal)
        }
      }
      let downloaded: DownloadedImage
      try {
        downloaded = await fetchCandidate(url)
      } catch (error) {
        if (signal?.aborted || (error instanceof Error && error.message === 'search_cancelled'))
          throw error
        if (deadline?.signal.aborted) throw new Error('retrieval_upstream_error', { cause: error })
        checkCurrent()
        if (!source.fallbackImageUrl) throw error
        downloaded = await fetchCandidate(source.fallbackImageUrl)
      } finally {
        clearTimeout(deadlineTimer)
        signal?.removeEventListener('abort', cancelDeadline)
      }
      checkCurrent()
      if (!supportedImageMime(downloaded.mime)) throw new Error('image_mime_unsupported')
      if (downloaded.bytes.byteLength > maximumSourceBytes) throw new Error('image_limit')
      const { mime, bytes } = options.normalizeImage
        ? await options.normalizeImage(downloaded)
        : downloaded
      checkCurrent()
      if (!supportedImageMime(mime)) throw new Error('image_mime_unsupported')
      if (bytes.byteLength > 2 * 1024 * 1024) throw new Error('retrieval_upstream_error')
      return new TextEncoder().encode(
        JSON.stringify({ mime, data_base64: Buffer.from(bytes).toString('base64') }),
      )
    }
    throw new Error('retrieval_capability_unavailable')
  }
  proxy.clear = () => {
    generation += 1
    allowedImages.clear()
  }
  return proxy
}

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
  if (capability === 'image-fetch.v1') {
    exact(body, ['url'])
    return { version: 1, operation: 'image-fetch', input: { url: safeHttpsUrl(body.url) } }
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
