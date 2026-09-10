import { describe, it, expect, vi, afterEach } from 'vitest'
import * as searchModule from '../src/index'
import {
  ImageSearchError,
  configureImageSearch,
  webSearch,
  imageSearch,
  wisUsageWebSearch,
} from '../src/index'

const realFetch = globalThis.fetch
afterEach(() => {
  globalThis.fetch = realFetch
  delete process.env.SERPER_API_KEY
  delete process.env.SERPAPI_API_KEY
  configureImageSearch({})
})

function mockFetch(
  handler: (
    url: string,
    init?: RequestInit,
  ) => {
    ok: boolean
    status?: number
    json?: any
    text?: string
  },
) {
  globalThis.fetch = vi.fn(async (url: any, init: any) => {
    const r = handler(String(url), init)
    return {
      ok: r.ok,
      status: r.status ?? (r.ok ? 200 : 500),
      headers: new Map(),
      json: async () => r.json,
      text: async () => r.text ?? '',
    } as any
  }) as any
}

describe('webSearch (Serper)', () => {
  it('exports search only and has no WisWork runtime surface', () => {
    expect(Object.keys(searchModule).sort()).toEqual([
      'ImageSearchError',
      'configureImageSearch',
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
  it.each(['serpapi', 'serper', 'duckduckgo'] as const)(
    'keeps safe %s candidates when individual results are malformed',
    async (provider) => {
      if (provider === 'serpapi') process.env.SERPAPI_API_KEY = 'test-key'
      if (provider === 'serper') process.env.SERPER_API_KEY = 'test-key'
      const entry = (id: string) => ({
        title: id,
        original: `https://cdn.example.com/${id}.jpg`,
        imageUrl: `https://cdn.example.com/${id}.jpg`,
        image: `https://cdn.example.com/${id}.jpg`,
        link: 'https://example.com/source',
        url: 'https://example.com/source',
      })
      const records = [
        entry('first'),
        null,
        { ...entry('unsafe'), link: 'http://example.com', url: 'http://example.com' },
        { ...entry('fractional'), width: 12.5, imageWidth: 12.5, original_width: 12.5 },
        { ...entry('oversized'), width: 100_001, imageWidth: 100_001, original_width: 100_001 },
        entry('last'),
      ]
      mockFetch((url) =>
        url.includes('duckduckgo.com/?')
          ? { ok: true, text: 'vqd="123-456"' }
          : { ok: true, json: { images_results: records, images: records, results: records } },
      )
      const result = await imageSearch('mixed candidates', 2)
      expect(result.method).toBe(provider)
      expect(result.images.map((image) => image.title)).toEqual(['first', 'last'])
    },
  )

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
              thumbnailUrl: 'https://cdn.example.com/a-thumbnail.jpg',
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
      fallbackImageUrl: 'https://cdn.example.com/a-thumbnail.jpg',
      width: 800,
      height: 600,
    })
  })
})

describe('imageSearch (SerpApi)', () => {
  it('uses a main-process key provider when the environment is unset', async () => {
    configureImageSearch({ serpApiKey: () => 'stored-test-key' })
    mockFetch((rawUrl) => {
      expect(new URL(rawUrl).searchParams.get('api_key')).toBe('stored-test-key')
      return { ok: true, json: { images_results: [] } }
    })

    await expect(imageSearch('flowers', 3)).resolves.toEqual({
      images: [],
      method: 'serpapi',
    })
  })

  it('surfaces secure key resolution failure as typed configuration failure', async () => {
    configureImageSearch({
      serpApiKey: () => {
        throw new Error('secure_storage_unavailable')
      },
    })

    const failure = await imageSearch('flowers', 3).catch((error: unknown) => error)
    expect(failure).toBeInstanceOf(ImageSearchError)
    expect(failure).toMatchObject({ code: 'config', provider: 'serpapi' })
  })

  it('keeps the environment key as the override', async () => {
    process.env.SERPAPI_API_KEY = 'environment-key'
    configureImageSearch({ serpApiKey: () => 'stored-key' })
    mockFetch((rawUrl) => {
      expect(new URL(rawUrl).searchParams.get('api_key')).toBe('environment-key')
      return { ok: true, json: { images_results: [] } }
    })

    await imageSearch('flowers', 3)
  })

  it.each([
    [401, 'auth'],
    [429, 'quota'],
  ] as const)('surfaces SerpApi HTTP %s as a typed %s failure', async (status, code) => {
    process.env.SERPAPI_API_KEY = 'serpapi-test-key'
    mockFetch(() => ({ ok: false, status }))

    const failure = await imageSearch('flowers', 3).catch((error: unknown) => error)
    expect(failure).toBeInstanceOf(ImageSearchError)
    expect(failure).toMatchObject({ code, provider: 'serpapi' })
  })

  it('can test the configured provider without accepting a fallback success', async () => {
    process.env.SERPAPI_API_KEY = 'invalid-key'
    let request = 0
    mockFetch(() => {
      request += 1
      return request === 1 ? { ok: false, status: 401 } : { ok: true, json: { images_results: [] } }
    })

    await expect(imageSearch('configuration test', 1, { fallback: false })).rejects.toMatchObject({
      code: 'auth',
      provider: 'serpapi',
    })
    expect(request).toBe(1)
  })

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
              thumbnail: 'https://cdn.example.com/flowers-thumbnail.jpg',
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
          fallbackImageUrl: 'https://cdn.example.com/flowers-thumbnail.jpg',
          sourceUrl: 'https://example.com/flowers',
          source: 'Example',
          width: 1200,
          height: 800,
        },
      ],
    })
  })

  it('accepts the alternate SerpApi image_results field', async () => {
    process.env.SERPAPI_API_KEY = 'serpapi-test-key'
    mockFetch(() => ({
      ok: true,
      json: {
        image_results: [
          {
            title: 'Team meeting',
            original: 'https://cdn.example.com/meeting.jpg',
            link: 'https://example.com/meeting',
            source: 'Example',
          },
        ],
      },
    }))

    await expect(imageSearch('team meeting office', 3)).resolves.toMatchObject({
      method: 'serpapi',
      images: [{ imageUrl: 'https://cdn.example.com/meeting.jpg' }],
    })
  })

  it('does not report an empty success when configured SerpApi and fallbacks fail', async () => {
    process.env.SERPAPI_API_KEY = 'serpapi-test-key'
    mockFetch(() => ({ ok: false }))

    await expect(imageSearch('team meeting office', 3)).rejects.toThrow(
      'image_search_upstream_error',
    )
  })

  it('preserves a legitimate empty SerpApi result', async () => {
    process.env.SERPAPI_API_KEY = 'serpapi-test-key'
    mockFetch(() => ({ ok: true, json: { images_results: [] } }))

    await expect(imageSearch('nothing here', 3)).resolves.toEqual({
      images: [],
      method: 'serpapi',
    })
  })

  it('rejects a configured provider failure when DuckDuckGo only confirms an empty fallback', async () => {
    process.env.SERPAPI_API_KEY = 'serpapi-test-key'
    let request = 0
    mockFetch(() => {
      request += 1
      if (request === 1) return { ok: false, status: 429 }
      if (request === 2) return { ok: true, text: 'vqd="123-456"' }
      return { ok: true, json: { results: [] } }
    })

    await expect(imageSearch('team meeting office', 3)).rejects.toMatchObject({
      code: 'quota',
      provider: 'serpapi',
    })
  })

  it.each([{}, { images_results: 'invalid' }, { images_results: [null] }])(
    'rejects malformed SerpApi response schema',
    async (json) => {
      process.env.SERPAPI_API_KEY = 'serpapi-test-key'
      mockFetch(() => ({ ok: true, json }))

      await expect(imageSearch('flowers', 3, { fallback: false })).rejects.toMatchObject({
        code: 'parse',
        provider: 'serpapi',
      })
    },
  )

  it.each([
    ['http://cdn.example.com/image.jpg', 'https://example.com/source'],
    ['https://localhost/image.jpg', 'https://example.com/source'],
    ['https://127.0.0.1/image.jpg', 'https://example.com/source'],
    ['https://10.0.0.4/image.jpg', 'https://example.com/source'],
    ['https://[::1]/image.jpg', 'https://example.com/source'],
    ['https://user:password@example.com/image.jpg', 'https://example.com/source'],
    ['https://cdn.example.com/image.jpg', 'https://192.168.1.2/source'],
  ])('does not return unsafe image/source URLs', async (original, link) => {
    process.env.SERPAPI_API_KEY = 'serpapi-test-key'
    mockFetch(() => ({
      ok: true,
      json: { images_results: [{ title: 'unsafe', original, link, source: 'Example' }] },
    }))

    await expect(imageSearch('flowers', 3, { fallback: false })).rejects.toMatchObject({
      code: 'parse',
      provider: 'serpapi',
    })
  })

  it('does not report an empty success when configured Serper and fallbacks fail', async () => {
    process.env.SERPER_API_KEY = 'serper-test-key'
    mockFetch(() => ({ ok: false }))

    await expect(imageSearch('team meeting office', 3)).rejects.toThrow(
      'image_search_upstream_error',
    )
  })

  it('rejects a failed DuckDuckGo image response instead of reporting zero results', async () => {
    process.env.SERPAPI_API_KEY = 'serpapi-test-key'
    let request = 0
    mockFetch(() => {
      request += 1
      if (request === 1) return { ok: false }
      if (request === 2) return { ok: true, text: 'vqd="123-456"' }
      return { ok: false, json: {} }
    })

    await expect(imageSearch('team meeting office', 3)).rejects.toThrow(
      'image_search_upstream_error',
    )
  })

  it('preserves DuckDuckGo timeout classification', async () => {
    mockFetch(() => {
      throw new DOMException('Aborted', 'AbortError')
    })

    await expect(imageSearch('team meeting office', 3)).rejects.toMatchObject({
      code: 'timeout',
      provider: 'duckduckgo',
    })
  })
})
