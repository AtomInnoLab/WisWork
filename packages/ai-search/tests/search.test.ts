import { describe, it, expect, vi, afterEach } from 'vitest'
import * as searchModule from '../src/index'
import { webSearch, imageSearch, wisUsageWebSearch } from '../src/index'

const realFetch = globalThis.fetch
afterEach(() => {
  globalThis.fetch = realFetch
  delete process.env.SERPER_API_KEY
  delete process.env.SERPAPI_API_KEY
})

function mockFetch(
  handler: (url: string, init?: RequestInit) => { ok: boolean; json?: any; text?: string },
) {
  globalThis.fetch = vi.fn(async (url: any, init: any) => {
    const r = handler(String(url), init)
    return {
      ok: r.ok,
      status: r.ok ? 200 : 500,
      headers: new Map(),
      json: async () => r.json,
      text: async () => r.text ?? '',
    } as any
  }) as any
}

describe('webSearch (Serper)', () => {
  it('exports search only and has no WisWork runtime surface', () => {
    expect(Object.keys(searchModule).sort()).toEqual([
      'imageSearch',
      'webSearch',
      'wisUsageWebSearch',
    ])
  })

  it('parses organic results + answer box', async () => {
    process.env.SERPER_API_KEY = 'test-key'
    mockFetch((url) => {
      expect(url).toBe('https://google.serper.dev/search')
      return {
        ok: true,
        json: {
          answerBox: { answer: '42' },
          organic: [
            { title: 'A', link: 'https://a.com', snippet: 'sa' },
            { title: 'B', link: 'https://b.com', snippet: 'sb' },
          ],
        },
      }
    })
    const r = await webSearch('meaning of life', 5)
    expect(r.method).toBe('serper')
    expect(r.answer).toBe('42')
    expect(r.results).toHaveLength(2)
    expect(r.results[0]).toEqual({ title: 'A', url: 'https://a.com', snippet: 'sa' })
  })

  it('falls back to DuckDuckGo when no key', async () => {
    mockFetch((url) => {
      expect(url).toContain('duckduckgo.com')
      return {
        ok: true,
        text: '<a class="result__a" href="/l/?uddg=https%3A%2F%2Fx.com">X Title</a>',
      }
    })
    const r = await webSearch('q', 3)
    expect(r.method).toBe('duckduckgo')
    expect(r.results[0]?.url).toBe('https://x.com')
    expect(r.results[0]?.title).toBe('X Title')
  })
})

describe('wisUsageWebSearch', () => {
  it('uses the authenticated WisUsage Xiaosu endpoint and sanitizes rich results', async () => {
    const fetchWithAuth = vi.fn(async (request: (token: string) => Promise<Response>) =>
      request('jwt-token'),
    )
    const fetchImpl = vi.fn(async (input: string | URL | Request, init?: RequestInit) => {
      const url = new URL(String(input))
      expect(url.origin + url.pathname).toBe('https://wisusage.atominnolab.com/v1/xiaosu/search')
      expect(Object.fromEntries(url.searchParams)).toEqual({
        q: '人工智能最新进展',
        count: '10',
        enableContent: 'true',
        mainText: 'true',
        contentType: 'MARKDOWN',
      })
      expect(init).toMatchObject({
        method: 'GET',
        redirect: 'error',
        headers: { Authorization: 'Bearer jwt-token', 'x-req-location': 'sg' },
      })
      return new Response(
        JSON.stringify({
          queryContext: { originalQuery: '人工智能最新进展' },
          webPages: {
            value: [
              {
                name: 'AI update',
                url: 'https://example.com/update',
                snippet: 'short',
                mainText: 'focused details',
                content: '# full article',
                datePublished: '2026-09-01',
                score: 0.9,
              },
            ],
          },
        }),
        { status: 200, headers: { 'content-type': 'application/json' } },
      )
    })

    const result = await wisUsageWebSearch('人工智能最新进展', 10, {
      fetchWithAuth,
      fetch: fetchImpl,
    })

    expect(fetchWithAuth).toHaveBeenCalledOnce()
    expect(result).toEqual({
      method: 'wisusage-xiaosu',
      results: [
        {
          title: 'AI update',
          url: 'https://example.com/update',
          snippet: 'focused details',
        },
      ],
    })
  })

  it('rejects invalid inputs, redirects, oversized bodies, and unsafe result URLs', async () => {
    const fetchWithAuth = async (request: (token: string) => Promise<Response>) => request('jwt')
    const response = (body: string, init?: ResponseInit) => async () =>
      new Response(body, {
        status: 200,
        headers: { 'content-type': 'application/json' },
        ...init,
      })

    await expect(
      wisUsageWebSearch('', 10, { fetchWithAuth, fetch: response('{}') }),
    ).rejects.toThrow('search_invalid_request')
    await expect(
      wisUsageWebSearch('query', 11, { fetchWithAuth, fetch: response('{}') }),
    ).rejects.toThrow('search_invalid_request')
    const redirected = new Response('{}', {
      headers: { 'content-type': 'application/json' },
    })
    Object.defineProperty(redirected, 'redirected', { value: true })
    await expect(
      wisUsageWebSearch('query', 10, {
        fetchWithAuth,
        fetch: async () => redirected,
      }),
    ).rejects.toThrow('search_upstream_error')
    await expect(
      wisUsageWebSearch('query', 10, {
        fetchWithAuth,
        fetch: response(
          JSON.stringify({
            webPages: { value: [{ name: 'x', url: 'http://127.0.0.1' }] },
          }),
        ),
      }),
    ).rejects.toThrow('search_invalid_response')
    await expect(
      wisUsageWebSearch('query', 10, {
        fetchWithAuth,
        fetch: response('x'.repeat(1_048_577)),
      }),
    ).rejects.toThrow('search_response_too_large')

    const cancelled = new AbortController()
    cancelled.abort()
    await expect(
      wisUsageWebSearch('query', 10, {
        fetchWithAuth,
        fetch: vi.fn(async (_input, init) => {
          if (init?.signal?.aborted) throw new DOMException('Aborted', 'AbortError')
          return response('{}')()
        }),
        signal: cancelled.signal,
      }),
    ).rejects.toThrow('search_cancelled')
  })
})

describe('imageSearch (Serper)', () => {
  it('parses images + filters copyright hosts', async () => {
    process.env.SERPER_API_KEY = 'test-key'
    mockFetch((url) => {
      expect(url).toBe('https://google.serper.dev/images')
      return {
        ok: true,
        json: {
          images: [
            {
              title: 'good',
              imageUrl: 'https://cdn.example.com/a.jpg',
              link: 'https://example.com',
              imageWidth: 800,
              imageHeight: 600,
            },
            {
              title: 'paid',
              imageUrl: 'https://gettyimages.com/x.jpg',
              link: 'https://gettyimages.com',
            },
          ],
        },
      }
    })
    const r = await imageSearch('cats', 8)
    expect(r.method).toBe('serper')
    expect(r.images).toHaveLength(1) // getty is filtered out
    expect(r.images[0]).toMatchObject({
      imageUrl: 'https://cdn.example.com/a.jpg',
      width: 800,
      height: 600,
    })
  })
})

describe('imageSearch (SerpApi)', () => {
  it('uses the SerpApi Google Images contract and parses image_results', async () => {
    process.env.SERPAPI_API_KEY = 'serpapi-test-key'
    mockFetch((rawUrl) => {
      const url = new URL(rawUrl)
      expect(url.origin + url.pathname).toBe('https://serpapi.com/search')
      expect(Object.fromEntries(url.searchParams)).toMatchObject({
        engine: 'google_images',
        q: 'flowers',
        hl: 'en',
        gl: 'us',
        api_key: 'serpapi-test-key',
      })
      return {
        ok: true,
        json: {
          images_results: [
            {
              title: 'Wildflowers',
              original: 'https://cdn.example.com/flowers.jpg',
              link: 'https://example.com/flowers',
              source: 'Example',
              original_width: 1200,
              original_height: 800,
            },
          ],
        },
      }
    })

    await expect(imageSearch('flowers', 3)).resolves.toEqual({
      method: 'serpapi',
      images: [
        {
          title: 'Wildflowers',
          imageUrl: 'https://cdn.example.com/flowers.jpg',
          sourceUrl: 'https://example.com/flowers',
          source: 'Example',
          width: 1200,
          height: 800,
        },
      ],
    })
  })
})
