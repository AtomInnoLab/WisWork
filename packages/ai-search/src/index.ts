/**
 * Search utilities (main process) — authenticated WisUsage web search plus
 * legacy Serper/DuckDuckGo image search.
 * Runs in the main process to avoid renderer CORS.
 */

import {
  COPYRIGHT_HOSTS,
  asRecord,
  safeHost,
  type ImageSearchResult,
  type WebSearchResult,
} from './shared'

export type { ImageSearchResult, WebSearchResult } from './shared'

const SERPER_KEY = () => process.env.SERPER_API_KEY ?? ''
const WISUSAGE_SEARCH_URL = 'https://wisusage.atominnolab.com/v1/xiaosu/search'
const WISUSAGE_RESPONSE_LIMIT = 1_048_576
const WISUSAGE_QUERY_LIMIT = 1_000
const WISUSAGE_TEXT_LIMIT = 8_192
const WISUSAGE_TIMEOUT_MS = 15_000

export interface WisUsageSearchOptions {
  readonly fetchWithAuth: (request: (accessToken: string) => Promise<Response>) => Promise<Response>
  readonly fetch?: typeof fetch
  readonly signal?: AbortSignal
}

function boundedSearchText(value: unknown, maximum: number): string {
  if (typeof value !== 'string') return ''
  const text = value.trim()
  return text.length > maximum ? text.slice(0, maximum) : text
}

function publicHttpsUrl(value: unknown): string {
  try {
    const url = new URL(String(value))
    if (
      url.protocol !== 'https:' ||
      url.username ||
      url.password ||
      url.port ||
      url.hostname === 'localhost' ||
      /^(?:127\.|10\.|192\.168\.|169\.254\.|0\.)/.test(url.hostname) ||
      /^172\.(?:1[6-9]|2\d|3[01])\./.test(url.hostname) ||
      url.hostname === '::1' ||
      url.hostname.endsWith('.local')
    )
      throw new Error()
    return url.href
  } catch (error) {
    throw new Error('search_invalid_response', { cause: error })
  }
}

async function boundedResponseJson(response: Response): Promise<unknown> {
  const contentLength = response.headers.get('content-length')
  const parsedContentLength = contentLength === null ? null : Number(contentLength)
  if (
    !response.ok ||
    response.redirected ||
    response.headers.get('content-type')?.split(';', 1)[0]?.trim().toLowerCase() !==
      'application/json' ||
    (parsedContentLength !== null &&
      (!Number.isSafeInteger(parsedContentLength) ||
        parsedContentLength < 0 ||
        parsedContentLength > WISUSAGE_RESPONSE_LIMIT))
  )
    throw new Error('search_upstream_error')
  const reader = response.body?.getReader()
  if (!reader) throw new Error('search_upstream_error')
  const chunks: Uint8Array[] = []
  let total = 0
  try {
    while (true) {
      const next = await reader.read()
      if (next.done) break
      total += next.value.byteLength
      if (total > WISUSAGE_RESPONSE_LIMIT) throw new Error('search_response_too_large')
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
  } catch (error) {
    if (error instanceof Error && error.message === 'search_response_too_large') throw error
    throw new Error('search_invalid_response', { cause: error })
  }
}

/** Authenticated WisUsage Xiaosu search used by both Standard and Enhanced desktop agents. */
export async function wisUsageWebSearch(
  query: string,
  maxResults = 10,
  options: WisUsageSearchOptions,
): Promise<{ results: WebSearchResult[]; method: 'wisusage-xiaosu' }> {
  const normalized = typeof query === 'string' ? query.trim() : ''
  if (
    !normalized ||
    normalized.length > WISUSAGE_QUERY_LIMIT ||
    !Number.isSafeInteger(maxResults) ||
    maxResults < 1 ||
    maxResults > 10
  )
    throw new Error('search_invalid_request')
  const fetchImpl = options.fetch ?? fetch
  const url = new URL(WISUSAGE_SEARCH_URL)
  url.searchParams.set('q', normalized)
  url.searchParams.set('count', '10')
  url.searchParams.set('enableContent', 'true')
  url.searchParams.set('mainText', 'true')
  url.searchParams.set('contentType', 'MARKDOWN')
  const timeout = new AbortController()
  const timer = setTimeout(() => timeout.abort(), WISUSAGE_TIMEOUT_MS)
  timer.unref?.()
  const abortFromCaller = () => timeout.abort()
  options.signal?.addEventListener('abort', abortFromCaller, { once: true })
  if (options.signal?.aborted) timeout.abort()
  let response: Response
  try {
    response = await options.fetchWithAuth((accessToken) =>
      fetchImpl(url, {
        method: 'GET',
        redirect: 'error',
        signal: timeout.signal,
        headers: { Authorization: `Bearer ${accessToken}`, 'x-req-location': 'sg' },
      }),
    )
  } catch (error) {
    if (timeout.signal.aborted)
      throw new Error(options.signal?.aborted ? 'search_cancelled' : 'search_timeout', {
        cause: error,
      })
    throw error
  } finally {
    clearTimeout(timer)
    options.signal?.removeEventListener('abort', abortFromCaller)
  }
  const data = asRecord(await boundedResponseJson(response))
  const rawResults = asRecord(data.webPages).value
  if (!Array.isArray(rawResults) || rawResults.length > 50)
    throw new Error('search_invalid_response')
  const results = rawResults.slice(0, maxResults).map((item) => {
    const entry = asRecord(item)
    const title = boundedSearchText(entry.name, 512)
    const url = publicHttpsUrl(entry.url)
    const snippet = boundedSearchText(
      entry.mainText || entry.snippet || entry.content,
      WISUSAGE_TEXT_LIMIT,
    )
    if (!title) throw new Error('search_invalid_response')
    return { title, url, snippet }
  })
  return { results, method: 'wisusage-xiaosu' }
}

// ── Web search ──────────────────────────────────────────────────────

export async function webSearch(
  query: string,
  maxResults = 6,
): Promise<{
  results: WebSearchResult[]
  answer?: string
  method: string
}> {
  const key = SERPER_KEY()
  if (key) {
    try {
      const resp = await fetchWithTimeout('https://google.serper.dev/search', {
        method: 'POST',
        headers: { 'X-API-KEY': key, 'Content-Type': 'application/json' },
        body: JSON.stringify({ q: query, num: maxResults, gl: 'us', hl: 'en' }),
      })
      if (resp.ok) {
        const data = asRecord(await resp.json())
        const organic: unknown[] = Array.isArray(data.organic) ? data.organic : []
        const results: WebSearchResult[] = organic.slice(0, maxResults).map((item) => {
          const o = asRecord(item)
          return {
            title: String(o.title ?? ''),
            url: String(o.link ?? ''),
            snippet: String(o.snippet ?? ''),
          }
        })
        const answerBox = asRecord(data.answerBox)
        const answerRaw =
          answerBox.answer || answerBox.snippet || asRecord(data.knowledgeGraph).description
        const answer = typeof answerRaw === 'string' && answerRaw ? answerRaw : undefined
        if (results.length) {
          return answer !== undefined
            ? { results, answer, method: 'serper' }
            : { results, method: 'serper' }
        }
      }
    } catch {
      /* fall back to DuckDuckGo */
    }
  }
  return { ...(await duckWebSearch(query, maxResults)), method: 'duckduckgo' }
}

// ── Image search ────────────────────────────────────────────────────

export async function imageSearch(
  query: string,
  maxResults = 8,
): Promise<{
  images: ImageSearchResult[]
  method: string
}> {
  const key = SERPER_KEY()
  if (key) {
    try {
      const resp = await fetchWithTimeout('https://google.serper.dev/images', {
        method: 'POST',
        headers: { 'X-API-KEY': key, 'Content-Type': 'application/json' },
        body: JSON.stringify({ q: query, num: Math.min(maxResults, 10), gl: 'us', hl: 'en' }),
      })
      if (resp.ok) {
        const data = asRecord(await resp.json())
        const raw: unknown[] = Array.isArray(data.images) ? data.images : []
        const images: ImageSearchResult[] = []
        for (const item of raw) {
          const img = asRecord(item)
          const imageUrl = String(img.imageUrl ?? img.original ?? '')
          if (!imageUrl) continue
          if (COPYRIGHT_HOSTS.some((d) => imageUrl.toLowerCase().includes(d))) continue
          const entry: ImageSearchResult = {
            title: String(img.title ?? ''),
            imageUrl,
            sourceUrl: String(img.link ?? ''),
            source: String(img.source ?? safeHost(img.link)),
          }
          if (typeof img.imageWidth === 'number') entry.width = img.imageWidth
          if (typeof img.imageHeight === 'number') entry.height = img.imageHeight
          images.push(entry)
          if (images.length >= maxResults) break
        }
        if (images.length) return { images, method: 'serper' }
      }
    } catch {
      /* fall back to DuckDuckGo */
    }
  }
  return { images: await duckImageSearch(query, maxResults), method: 'duckduckgo' }
}

// ── DuckDuckGo fallback (no key / quota exhausted) ──────────────────

async function duckWebSearch(
  query: string,
  maxResults: number,
): Promise<{ results: WebSearchResult[] }> {
  try {
    // DuckDuckGo HTML endpoint (lightweight, no key needed)
    const resp = await fetchWithTimeout(
      `https://html.duckduckgo.com/html/?q=${encodeURIComponent(query)}`,
      { headers: { 'User-Agent': 'Mozilla/5.0' } },
    )
    const html = await resp.text()
    const results: WebSearchResult[] = []
    const re = /<a[^>]*class="result__a"[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/g
    let m: RegExpExecArray | null
    while ((m = re.exec(html)) !== null && results.length < maxResults) {
      const url = decodeDuckUrl(m[1]!)
      const title = stripTags(m[2]!)
      if (url && title) results.push({ title, url, snippet: '' })
    }
    return { results }
  } catch {
    return { results: [] }
  }
}

async function duckImageSearch(query: string, maxResults: number): Promise<ImageSearchResult[]> {
  try {
    // DuckDuckGo i.js needs a vqd token, so it takes two steps
    const tokenResp = await fetchWithTimeout(
      `https://duckduckgo.com/?q=${encodeURIComponent(query)}`,
      { headers: { 'User-Agent': 'Mozilla/5.0' } },
    )
    const tokenHtml = await tokenResp.text()
    const vqd = /vqd=["']?([\d-]+)["']?/.exec(tokenHtml)?.[1]
    if (!vqd) return []
    const resp = await fetchWithTimeout(
      `https://duckduckgo.com/i.js?l=us-en&o=json&q=${encodeURIComponent(query)}&vqd=${vqd}`,
      { headers: { 'User-Agent': 'Mozilla/5.0', Referer: 'https://duckduckgo.com/' } },
    )
    const data = asRecord(await resp.json())
    const list: unknown[] = Array.isArray(data.results) ? data.results : []
    const out: ImageSearchResult[] = []
    for (const item of list.slice(0, maxResults)) {
      const img = asRecord(item)
      const imageUrl = String(img.image ?? '')
      if (!imageUrl || COPYRIGHT_HOSTS.some((d) => imageUrl.toLowerCase().includes(d))) continue
      const entry: ImageSearchResult = {
        title: String(img.title ?? ''),
        imageUrl,
        sourceUrl: String(img.url ?? ''),
        source: safeHost(img.url),
      }
      if (typeof img.width === 'number') entry.width = img.width
      if (typeof img.height === 'number') entry.height = img.height
      out.push(entry)
    }
    return out
  } catch {
    return []
  }
}

// ── utils ───────────────────────────────────────────────────────────

async function fetchWithTimeout(
  url: string,
  init: RequestInit & { timeoutMs?: number } = {},
): Promise<Response> {
  const controller = new AbortController()
  const t = setTimeout(() => controller.abort(), init.timeoutMs ?? 15000)
  try {
    return await fetch(url, { ...init, signal: controller.signal })
  } finally {
    clearTimeout(t)
  }
}

function stripTags(s: string): string {
  return s
    .replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&')
    .replace(/&#x27;/g, "'")
    .trim()
}

function decodeDuckUrl(href: string): string {
  // DuckDuckGo result links are often /l/?uddg=<encoded>
  const m = /[?&]uddg=([^&]+)/.exec(href)
  if (m) return decodeURIComponent(m[1]!)
  return href.startsWith('http') ? href : ''
}
