import { describe, expect, it, vi } from 'vitest'
import {
  collectBoundedImageBytes,
  createOfficeRemoteImageDownloader,
  createPinnedLookup,
  createOfficeLocalSearchProxy,
  createOfficeRetrievalProxy,
  officeRetrievalEndpointFromEnv,
  resolvePublicImageRedirect,
} from '../src/main/office-retrieval-proxy'

const IMAGE_FETCH_ENDPOINT = 'https://office.8-216-134-194.sslip.io/office-image-fetch'

const TEST_ENDPOINT = 'https://retrieval.test.invalid/v1/office/retrieval'
const TEST_SERVICES = {
  [TEST_ENDPOINT]: {
    contract: 'wiswork-office-retrieval-v1',
    ssrfProtection: 'dns-rebinding-and-redirect-hops-v1',
  },
} as const

describe('Office fixed retrieval proxy', () => {
  it('authenticates the fixed remote image request and returns raw image bytes', async () => {
    const fetch = vi.fn(
      async () =>
        new Response(new Uint8Array([0x89, 0x50, 0x4e, 0x47]), {
          headers: { 'content-type': 'image/png', 'content-length': '4' },
        }),
    )
    const download = createOfficeRemoteImageDownloader({
      fetch,
      fetchWithAuth: (request) => request('pc-token'),
    })

    await expect(download('https://images.example/cover.png')).resolves.toEqual({
      mime: 'image/png',
      bytes: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
    })
    expect(fetch).toHaveBeenCalledWith(
      IMAGE_FETCH_ENDPOINT,
      expect.objectContaining({
        method: 'POST',
        redirect: 'error',
        headers: expect.objectContaining({ authorization: 'Bearer pc-token' }),
        body: JSON.stringify({ url: 'https://images.example/cover.png' }),
      }),
    )
  })

  it.each([
    [new Response('no', { status: 503 }), 'image_fetch_unavailable'],
    [new Response('large', { status: 413 }), 'image_limit'],
    [new Response('webp', { status: 415 }), 'image_mime_unsupported'],
    [
      new Response('partial', { status: 206, headers: { 'content-type': 'image/jpeg' } }),
      'image_fetch_unavailable',
    ],
    [new Response('no', { headers: { 'content-type': 'text/plain' } }), 'image_mime_unsupported'],
    [
      new Response('x', {
        headers: { 'content-type': 'image/jpeg', 'content-length': String(10 * 1024 * 1024 + 1) },
      }),
      'image_limit',
    ],
    [
      new Response(new Uint8Array(10 * 1024 * 1024 + 1), {
        headers: { 'content-type': 'image/jpeg' },
      }),
      'image_limit',
    ],
  ])('bounds remote image responses %#', async (response, error) => {
    const download = createOfficeRemoteImageDownloader({
      fetch: vi.fn(async () => response),
      fetchWithAuth: (request) => request('token'),
    })
    await expect(download('https://images.example/cover.jpg')).rejects.toThrow(error)
  })

  it('cancels the remote response stream when its bytes exceed the limit', async () => {
    const cancel = vi.fn()
    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new Uint8Array(10 * 1024 * 1024))
        controller.enqueue(new Uint8Array([1]))
      },
      cancel,
    })
    const download = createOfficeRemoteImageDownloader({
      fetch: vi.fn(async () => new Response(body, { headers: { 'content-type': 'image/png' } })),
      fetchWithAuth: (request) => request('token'),
    })
    await expect(download('https://images.example/large.png')).rejects.toThrow('image_limit')
    expect(cancel).toHaveBeenCalledOnce()
  })

  it('distinguishes caller cancellation from remote unavailability', async () => {
    const download = createOfficeRemoteImageDownloader({
      fetch: vi.fn(
        async (_url, init) =>
          new Promise<Response>((_resolve, reject) =>
            init?.signal?.addEventListener('abort', () => reject(new Error('aborted')), {
              once: true,
            }),
          ),
      ),
      fetchWithAuth: (request) => request('token'),
    })
    const controller = new AbortController()
    const pending = download('https://images.example/cover.jpg', controller.signal)
    controller.abort()
    await expect(pending).rejects.toThrow('search_cancelled')
  })

  it('bounds authentication and token refresh within the same remote image deadline', async () => {
    vi.useFakeTimers()
    try {
      let rejectAuth: ((reason: Error) => void) | undefined
      const download = createOfficeRemoteImageDownloader({
        fetch: vi.fn(),
        fetchWithAuth: () =>
          new Promise<Response>((_resolve, reject) => {
            rejectAuth = reject
          }),
        timeoutMs: 25,
      })
      const pending = download('https://images.example/cover.jpg')
      const assertion = expect(pending).rejects.toThrow('image_fetch_unavailable')
      await vi.advanceTimersByTimeAsync(25)
      await assertion
      rejectAuth?.(new Error('late auth failure'))
      await Promise.resolve()
    } finally {
      vi.useRealTimers()
    }
  })

  it('ends while authentication is pending when the caller cancels', async () => {
    let rejectAuth: ((reason: Error) => void) | undefined
    const download = createOfficeRemoteImageDownloader({
      fetch: vi.fn(),
      fetchWithAuth: () =>
        new Promise<Response>((_resolve, reject) => {
          rejectAuth = reject
        }),
    })
    const controller = new AbortController()
    const pending = download('https://images.example/cover.jpg', controller.signal)
    const assertion = expect(pending).rejects.toThrow('search_cancelled')
    controller.abort()
    await assertion
    rejectAuth?.(new Error('late auth failure'))
    await Promise.resolve()
  })

  it('uses remote first and reserves strict local download for remote unavailability', async () => {
    const bytes = new Uint8Array([0xff, 0xd8])
    const remoteDownloadImage = vi.fn(async () => ({ mime: 'image/jpeg' as const, bytes }))
    const downloadImage = vi.fn(async () => ({ mime: 'image/jpeg' as const, bytes }))
    const proxy = createOfficeLocalSearchProxy({
      fetchWithAuth: vi.fn(),
      remoteDownloadImage,
      downloadImage,
      searchImages: async () => ({
        images: [
          {
            title: 'Cover',
            imageUrl: 'https://images.example/cover.jpg',
            sourceUrl: 'https://example.com',
            source: 'example',
          },
        ],
        method: 'test',
      }),
    })
    await expect(
      proxy('image-fetch.v1', { url: 'https://images.example/cover.jpg' }),
    ).rejects.toThrow('retrieval_invalid_request')
    expect(remoteDownloadImage).not.toHaveBeenCalled()
    await proxy('image-search.v1', { query: 'cover', max_results: 1 })
    await proxy('image-fetch.v1', { url: 'https://images.example/cover.jpg' })
    expect(remoteDownloadImage).toHaveBeenCalledOnce()
    expect(downloadImage).not.toHaveBeenCalled()
  })

  it.each(['image_limit', 'image_mime_unsupported', 'search_cancelled'])(
    'does not locally retry a remote semantic/cancellation failure: %s',
    async (code) => {
      const downloadImage = vi.fn()
      const proxy = createOfficeLocalSearchProxy({
        fetchWithAuth: vi.fn(),
        remoteDownloadImage: vi.fn(async () => {
          throw new Error(code)
        }),
        downloadImage,
        searchImages: async () => ({
          images: [
            {
              title: 'Cover',
              imageUrl: 'https://images.example/cover.jpg',
              sourceUrl: 'https://example.com',
              source: 'example',
            },
          ],
          method: 'test',
        }),
      })
      await proxy('image-search.v1', { query: 'cover', max_results: 1 })
      await expect(
        proxy('image-fetch.v1', { url: 'https://images.example/cover.jpg' }),
      ).rejects.toThrow(code)
      expect(downloadImage).not.toHaveBeenCalled()
    },
  )

  it.each([
    [413, 'image_limit'],
    [415, 'image_mime_unsupported'],
  ])('does not locally retry Relay semantic status %i', async (status, error) => {
    const downloadImage = vi.fn()
    const proxy = createOfficeLocalSearchProxy({
      fetchWithAuth: vi.fn(),
      remoteDownloadImage: createOfficeRemoteImageDownloader({
        fetch: vi.fn(async () => new Response('rejected', { status })),
        fetchWithAuth: (request) => request('token'),
      }),
      downloadImage,
      searchImages: async () => ({
        images: [
          {
            title: 'Cover',
            imageUrl: 'https://images.example/cover.jpg',
            sourceUrl: 'https://example.com',
            source: 'example',
          },
        ],
        method: 'test',
      }),
    })
    await proxy('image-search.v1', { query: 'cover', max_results: 1 })
    await expect(
      proxy('image-fetch.v1', { url: 'https://images.example/cover.jpg' }),
    ).rejects.toThrow(error)
    expect(downloadImage).not.toHaveBeenCalled()
  })

  it.each(['image_limit', 'image_mime_unsupported'])(
    'advances from a semantic original failure to its authorized fallback: %s',
    async (code) => {
      const remoteDownloadImage = vi.fn(async (url: string) => {
        if (url.endsWith('/original.webp')) throw new Error(code)
        return { mime: 'image/jpeg' as const, bytes: new Uint8Array([0xff, 0xd8]) }
      })
      const downloadImage = vi.fn()
      const proxy = createOfficeLocalSearchProxy({
        fetchWithAuth: vi.fn(),
        remoteDownloadImage,
        downloadImage,
        searchImages: async () => ({
          images: [
            {
              title: 'Cover',
              imageUrl: 'https://images.example/original.webp',
              fallbackImageUrl: 'https://images.example/thumbnail.jpg',
              sourceUrl: 'https://example.com',
              source: 'example',
            },
          ],
          method: 'test',
        }),
      })
      await proxy('image-search.v1', { query: 'cover', max_results: 1 })
      await expect(
        proxy('image-fetch.v1', { url: 'https://images.example/original.webp' }),
      ).resolves.toBeInstanceOf(Uint8Array)
      expect(remoteDownloadImage).toHaveBeenNthCalledWith(
        2,
        'https://images.example/thumbnail.jpg',
        expect.any(AbortSignal),
      )
      expect(downloadImage).not.toHaveBeenCalled()
    },
  )

  it('rescues the same candidate locally when the remote service is unavailable', async () => {
    const remoteDownloadImage = vi.fn(async () => {
      throw new Error('image_fetch_unavailable')
    })
    const downloadImage = vi.fn(async () => ({
      mime: 'image/png' as const,
      bytes: new Uint8Array([1]),
    }))
    const proxy = createOfficeLocalSearchProxy({
      fetchWithAuth: vi.fn(),
      remoteDownloadImage,
      downloadImage,
      searchImages: async () => ({
        images: [
          {
            title: 'Cover',
            imageUrl: 'https://images.example/original.png',
            fallbackImageUrl: 'https://images.example/fallback.png',
            sourceUrl: 'https://example.com',
            source: 'example',
          },
        ],
        method: 'test',
      }),
    })
    await proxy('image-search.v1', { query: 'cover', max_results: 1 })
    await proxy('image-fetch.v1', { url: 'https://images.example/original.png' })
    expect(remoteDownloadImage).toHaveBeenCalledWith(
      'https://images.example/original.png',
      expect.any(AbortSignal),
    )
    expect(downloadImage).toHaveBeenCalledWith(
      'https://images.example/original.png',
      expect.any(AbortSignal),
    )
  })

  it('shares one timeout across remote, local rescue, and same-candidate fallback', async () => {
    vi.useFakeTimers()
    try {
      const remoteDownloadImage = vi.fn(
        async (_url: string, signal?: AbortSignal) =>
          new Promise<never>((_resolve, reject) =>
            signal?.addEventListener('abort', () => reject(new Error('search_cancelled')), {
              once: true,
            }),
          ),
      )
      const downloadImage = vi.fn()
      const proxy = createOfficeLocalSearchProxy({
        fetchWithAuth: vi.fn(),
        remoteDownloadImage,
        downloadImage,
        imageTimeoutMs: 5,
        searchImages: async () => ({
          images: [
            {
              title: 'Cover',
              imageUrl: 'https://images.example/original.png',
              fallbackImageUrl: 'https://images.example/fallback.png',
              sourceUrl: 'https://example.com',
              source: 'example',
            },
          ],
          method: 'test',
        }),
      })
      await proxy('image-search.v1', { query: 'cover', max_results: 1 })
      const pending = proxy('image-fetch.v1', { url: 'https://images.example/original.png' })
      const rejected = expect(pending).rejects.toThrow('retrieval_upstream_error')
      await vi.advanceTimersByTimeAsync(5)
      await rejected
      expect(remoteDownloadImage).toHaveBeenCalledOnce()
      expect(downloadImage).not.toHaveBeenCalled()
    } finally {
      vi.useRealTimers()
    }
  })
  it('returns source dimensions and falls back to the same searched image rendition', async () => {
    const downloadImage = vi.fn(async (url: string) => {
      if (url.endsWith('/original.png')) throw new Error('retrieval_upstream_error')
      return { mime: 'image/jpeg' as const, bytes: new Uint8Array([0xff, 0xd8, 0xff, 0xd9]) }
    })
    const proxy = createOfficeLocalSearchProxy({
      fetchWithAuth: vi.fn(),
      downloadImage,
      searchImages: async () => ({
        images: [
          {
            title: 'Cover',
            imageUrl: 'https://images.example/original.png',
            fallbackImageUrl: 'https://images.example/thumbnail.jpg',
            sourceUrl: 'https://example.com/cover',
            source: 'example.com',
            width: 1600,
            height: 900,
          },
        ],
        method: 'serper',
      }),
    })

    const search = JSON.parse(
      new TextDecoder().decode(
        await proxy('image-search.v1', { query: 'private cover query', max_results: 1 }),
      ),
    )
    expect(search.images[0]).toMatchObject({ width: 1600, height: 900 })
    expect(JSON.stringify(search)).not.toContain('thumbnail.jpg')
    const fetched = JSON.parse(
      new TextDecoder().decode(
        await proxy('image-fetch.v1', { url: 'https://images.example/original.png' }),
      ),
    )
    expect(fetched.mime).toBe('image/jpeg')
    expect(downloadImage).toHaveBeenNthCalledWith(
      1,
      'https://images.example/original.png',
      undefined,
    )
    expect(downloadImage).toHaveBeenNthCalledWith(
      2,
      'https://images.example/thumbnail.jpg',
      undefined,
    )
  })

  it('refreshes expired fallback authority from the renewed matching result', async () => {
    vi.useFakeTimers()
    try {
      let fallbackImageUrl: string | undefined = 'https://images.example/old-thumbnail.jpg'
      const searchImages = vi.fn(async () => ({
        images: [
          {
            title: 'Cover',
            imageUrl: 'https://images.example/original.png',
            ...(fallbackImageUrl ? { fallbackImageUrl } : {}),
            sourceUrl: 'https://example.com/cover',
            source: 'example.com',
          },
        ],
        method: 'serper',
      }))
      const downloadImage = vi.fn(async (url: string) => {
        if (url.endsWith('/original.png')) throw new Error('retrieval_upstream_error')
        return { mime: 'image/jpeg' as const, bytes: new Uint8Array([0xff, 0xd8, 0xff, 0xd9]) }
      })
      const proxy = createOfficeLocalSearchProxy({
        fetchWithAuth: vi.fn(),
        downloadImage,
        searchImages,
      })
      await proxy('image-search.v1', { query: 'cover', max_results: 1 })
      fallbackImageUrl = 'https://images.example/new-thumbnail.jpg'
      await vi.advanceTimersByTimeAsync(15 * 60_000 + 1)
      await proxy('image-fetch.v1', { url: 'https://images.example/original.png' })
      expect(downloadImage).toHaveBeenLastCalledWith(
        'https://images.example/new-thumbnail.jpg',
        undefined,
      )
    } finally {
      vi.useRealTimers()
    }
  })

  it('omits unsafe layout dimensions from injected image search results', async () => {
    const proxy = createOfficeLocalSearchProxy({
      fetchWithAuth: vi.fn(),
      searchImages: async () => ({
        images: [
          {
            title: 'Cover',
            imageUrl: 'https://images.example/cover.png',
            sourceUrl: 'https://example.com/cover',
            source: 'example.com',
            width: 12.5,
            height: 100_001,
          },
        ],
        method: 'test',
      }),
    })
    const search = JSON.parse(
      new TextDecoder().decode(await proxy('image-search.v1', { query: 'cover', max_results: 1 })),
    )
    expect(search.images[0]).not.toHaveProperty('width')
    expect(search.images[0]).not.toHaveProperty('height')
  })

  it('clears image-search authority so a reconnected session must search again', async () => {
    const downloadImage = vi.fn(async () => ({
      mime: 'image/png' as const,
      bytes: new Uint8Array(4),
    }))
    const proxy = createOfficeLocalSearchProxy({
      fetchWithAuth: vi.fn(),
      downloadImage,
      searchImages: async () => ({
        images: [
          {
            title: 'Cover',
            imageUrl: 'https://images.example/cover.png',
            sourceUrl: 'https://example.com/cover',
            source: 'example.com',
          },
        ],
        method: 'serpapi',
      }),
    })
    await proxy('image-search.v1', { query: 'private cover query', max_results: 1 })
    proxy.clear?.()
    await expect(
      proxy('image-fetch.v1', { url: 'https://images.example/cover.png' }),
    ).rejects.toThrow('retrieval_invalid_request')
    expect(downloadImage).not.toHaveBeenCalled()
    await proxy('image-search.v1', { query: 'new cover query', max_results: 1 })
    await expect(
      proxy('image-fetch.v1', { url: 'https://images.example/cover.png' }),
    ).resolves.toBeInstanceOf(Uint8Array)
  })

  it.each(['initial search', 'expired revalidation'])(
    'does not restore cleared image authority after a late %s',
    async (phase) => {
      vi.useFakeTimers()
      try {
        const found = {
          images: [
            {
              title: 'Cover',
              imageUrl: 'https://images.example/cover.png',
              sourceUrl: 'https://example.com/cover',
              source: 'example.com',
            },
          ],
          method: 'serpapi',
        }
        let finish!: (value: typeof found) => void
        const searchImages = vi.fn(
          () =>
            new Promise<typeof found>((resolve) => {
              finish = resolve
            }),
        )
        const downloadImage = vi.fn(async () => ({
          mime: 'image/png' as const,
          bytes: new Uint8Array(4),
        }))
        const proxy = createOfficeLocalSearchProxy({
          fetchWithAuth: vi.fn(),
          searchImages,
          downloadImage,
        })
        if (phase === 'expired revalidation') {
          const first = proxy('image-search.v1', { query: 'private cover query', max_results: 1 })
          finish(found)
          await first
          vi.setSystemTime(Date.now() + 15 * 60_000 + 1)
        }
        const pending =
          phase === 'initial search'
            ? proxy('image-search.v1', { query: 'private cover query', max_results: 1 })
            : proxy('image-fetch.v1', { url: 'https://images.example/cover.png' })
        const settled = pending.then(
          () => 'success',
          (error: Error) => error.message,
        )
        proxy.clear?.()
        finish(found)
        expect(await settled).toBe('search_cancelled')
        await expect(
          proxy('image-fetch.v1', { url: 'https://images.example/cover.png' }),
        ).rejects.toThrow('retrieval_invalid_request')
        expect(downloadImage).not.toHaveBeenCalled()
      } finally {
        vi.useRealTimers()
      }
    },
  )

  it.each([true, false])(
    'revalidates expired search provenance before downloading (still listed: %s)',
    async (stillListed) => {
      const image = {
        title: 'Volcano',
        imageUrl: 'https://images.example/volcano.jpg',
        sourceUrl: 'https://example.com/',
        source: 'example',
      }
      const searchImages = vi.fn(async () => ({ images: [image], method: 'test' }))
      const downloadImage = vi.fn(async () => ({
        mime: 'image/jpeg' as const,
        bytes: new Uint8Array([0xff, 0xd8, 0xff, 0xd9]),
      }))
      const proxy = createOfficeLocalSearchProxy({
        fetchWithAuth: vi.fn(),
        searchImages,
        downloadImage,
      })
      vi.useFakeTimers()
      try {
        await proxy('image-search.v1', { query: 'volcano', max_results: 3 })
        if (!stillListed) searchImages.mockResolvedValue({ images: [], method: 'test' })
        await vi.advanceTimersByTimeAsync(15 * 60_000 + 1)
        if (stillListed) {
          await expect(proxy('image-fetch.v1', { url: image.imageUrl })).resolves.toBeInstanceOf(
            Uint8Array,
          )
          expect(downloadImage).toHaveBeenCalledTimes(1)
        } else {
          await expect(proxy('image-fetch.v1', { url: image.imageUrl })).rejects.toThrow(
            'retrieval_invalid_request',
          )
          expect(downloadImage).not.toHaveBeenCalled()
        }
        expect(searchImages).toHaveBeenCalledTimes(2)
        expect(searchImages).toHaveBeenLastCalledWith('volcano', 3)
      } finally {
        vi.useRealTimers()
      }
    },
  )

  it('accepts only bounded HTTPS redirects for searched images', () => {
    expect(resolvePublicImageRedirect('https://images.example/a', '/final.webp', 3)).toBe(
      'https://images.example/final.webp',
    )
    expect(() =>
      resolvePublicImageRedirect('https://images.example/a', 'http://images.example/final', 3),
    ).toThrow('retrieval_upstream_error')
    expect(() =>
      resolvePublicImageRedirect('https://images.example/a', 'https://127.0.0.1/final', 3),
    ).toThrow('retrieval_upstream_error')
    expect(() => resolvePublicImageRedirect('https://images.example/a', '/fourth-hop', 0)).toThrow(
      'retrieval_upstream_error',
    )
  })
  it('returns the pinned address array when Node requests lookup all mode', async () => {
    const lookup = createPinnedLookup({ address: '203.0.113.10', family: 4 })
    const result = await new Promise((resolve, reject) =>
      lookup('images.example', { all: true }, (error, addresses) =>
        error ? reject(error) : resolve(addresses),
      ),
    )
    expect(result).toEqual([{ address: '203.0.113.10', family: 4 }])
  })
  it('stops collecting a chunked image as soon as its byte budget is exceeded', async () => {
    async function* chunks() {
      yield new Uint8Array(3)
      yield new Uint8Array(3)
      throw new Error('must_not_read_past_limit')
    }
    await expect(collectBoundedImageBytes(chunks(), 5)).rejects.toThrow('image_limit')
  })
  it('is disabled without configuration and accepts only a compile-allowlisted exact endpoint', () => {
    expect(officeRetrievalEndpointFromEnv({}, TEST_SERVICES)).toBeNull()
    expect(() =>
      officeRetrievalEndpointFromEnv(
        { WISWORK_OFFICE_RETRIEVAL_URL: 'http://127.0.0.1' },
        TEST_SERVICES,
      ),
    ).toThrow('invalid_office_retrieval_url')
    expect(() =>
      officeRetrievalEndpointFromEnv(
        { WISWORK_OFFICE_RETRIEVAL_URL: 'https://attacker.invalid/retrieval' },
        TEST_SERVICES,
      ),
    ).toThrow('invalid_office_retrieval_url')
    expect(
      officeRetrievalEndpointFromEnv(
        { WISWORK_OFFICE_RETRIEVAL_URL: TEST_ENDPOINT },
        TEST_SERVICES,
      ),
    ).toBe(TEST_ENDPOINT)
  })

  it('uses the existing PC web and image search implementations', async () => {
    const webSearch = vi.fn(async () => ({
      results: [{ title: 'Web', url: 'https://example.com', snippet: 'Result' }],
      method: 'wisusage-xiaosu' as const,
    }))
    const searchImages = vi.fn(async () => ({
      images: [
        {
          title: 'Image',
          imageUrl: 'https://example.com/image.jpg',
          sourceUrl: 'https://example.com',
          source: 'example.com',
        },
      ],
      method: 'serpapi',
    }))
    const proxy = createOfficeLocalSearchProxy({
      fetchWithAuth: vi.fn(),
      webSearch,
      searchImages,
    })
    const web = await proxy('web-search.v1', { query: 'office', max_results: 3 })
    const images = await proxy('image-search.v1', { query: 'slides', max_results: 4 })
    expect(JSON.parse(new TextDecoder().decode(web))).toEqual({
      results: [{ title: 'Web', url: 'https://example.com', snippet: 'Result' }],
    })
    expect(JSON.parse(new TextDecoder().decode(images))).toEqual({
      images: [
        {
          title: 'Image',
          image_url: 'https://example.com/image.jpg',
          source_url: 'https://example.com/',
          source: 'example.com',
        },
      ],
    })
    expect(webSearch).toHaveBeenCalledWith(
      'office',
      3,
      expect.objectContaining({ fetchWithAuth: expect.any(Function) }),
    )
    expect(searchImages).toHaveBeenCalledWith('slides', 4)
  })

  it('downloads only a URL produced by image search and returns bounded image bytes', async () => {
    const bytes = new Uint8Array([0xff, 0xd8, 0xff, 0xd9])
    const downloadImage = vi.fn(async () => ({ mime: 'image/jpeg' as const, bytes }))
    const proxy = createOfficeLocalSearchProxy({
      fetchWithAuth: vi.fn(),
      downloadImage,
      searchImages: vi.fn(async () => ({
        images: [
          {
            title: 'LLM',
            imageUrl: 'https://images.example/llm.jpg',
            sourceUrl: 'https://example.com/llm',
            source: 'example.com',
          },
        ],
        method: 'serpapi',
      })),
    })
    await expect(
      proxy('image-fetch.v1', { url: 'https://images.example/llm.jpg' }),
    ).rejects.toThrow('retrieval_invalid_request')
    await proxy('image-search.v1', { query: 'llm', max_results: 1 })
    const result = await proxy('image-fetch.v1', { url: 'https://images.example/llm.jpg' })
    expect(JSON.parse(new TextDecoder().decode(result))).toEqual({
      mime: 'image/jpeg',
      data_base64: Buffer.from(bytes).toString('base64'),
    })
    expect(downloadImage).toHaveBeenCalledWith('https://images.example/llm.jpg', undefined)
  })

  it('normalizes downloaded search images before returning them to Office', async () => {
    const source = new Uint8Array(4 * 1024 * 1024)
    const normalized = new Uint8Array([0x89, 0x50, 0x4e, 0x47])
    const normalizeImage = vi.fn(async (_image: { mime: string; bytes: Uint8Array }) => ({
      mime: 'image/png' as const,
      bytes: normalized,
    }))
    const proxy = createOfficeLocalSearchProxy({
      fetchWithAuth: vi.fn(),
      downloadImage: vi.fn(async () => ({ mime: 'image/jpeg' as const, bytes: source })),
      normalizeImage,
      searchImages: vi.fn(async () => ({
        images: [
          {
            title: 'LLM',
            imageUrl: 'https://images.example/llm.jpg',
            sourceUrl: 'https://example.com/llm',
            source: 'example.com',
          },
        ],
        method: 'serpapi',
      })),
    })
    await proxy('image-search.v1', { query: 'llm', max_results: 1 })
    const result = await proxy('image-fetch.v1', { url: 'https://images.example/llm.jpg' })
    expect(normalizeImage).toHaveBeenCalledOnce()
    expect(normalizeImage.mock.calls[0]![0].mime).toBe('image/jpeg')
    expect(normalizeImage.mock.calls[0]![0].bytes).toBe(source)
    expect(JSON.parse(new TextDecoder().decode(result))).toEqual({
      mime: 'image/png',
      data_base64: Buffer.from(normalized).toString('base64'),
    })
  })

  it.each([
    { sourceBytes: 2 * 1024 * 1024 + 1, normalizedBytes: undefined, error: 'image_limit' },
    { sourceBytes: 10 * 1024 * 1024 + 1, normalizedBytes: 4, error: 'image_limit' },
    {
      sourceBytes: 4 * 1024 * 1024,
      normalizedBytes: 2 * 1024 * 1024 + 1,
      error: 'retrieval_upstream_error',
    },
  ])(
    'preserves source and transport boundaries: %o',
    async ({ sourceBytes, normalizedBytes, error }) => {
      const normalizeImage =
        normalizedBytes === undefined
          ? undefined
          : vi.fn(async () => ({
              mime: 'image/png' as const,
              bytes: new Uint8Array(normalizedBytes),
            }))
      const proxy = createOfficeLocalSearchProxy({
        fetchWithAuth: vi.fn(),
        normalizeImage,
        downloadImage: async () => ({ mime: 'image/jpeg', bytes: new Uint8Array(sourceBytes) }),
        searchImages: async () => ({
          images: [
            {
              title: 'Image',
              imageUrl: 'https://images.example/image.jpg',
              sourceUrl: 'https://example.com/',
              source: 'example',
            },
          ],
          method: 'test',
        }),
      })
      await proxy('image-search.v1', { query: 'image', max_results: 1 })
      await expect(
        proxy('image-fetch.v1', { url: 'https://images.example/image.jpg' }),
      ).rejects.toThrow(error)
      if (sourceBytes > 10 * 1024 * 1024) expect(normalizeImage).not.toHaveBeenCalled()
    },
  )

  it('rejects an image-search hostname that resolves to a private address', async () => {
    const proxy = createOfficeLocalSearchProxy({
      fetchWithAuth: vi.fn(),
      lookupAddresses: vi.fn(async () => [{ address: '127.0.0.1', family: 4 }]),
      searchImages: vi.fn(async () => ({
        images: [
          {
            title: 'Private',
            imageUrl: 'https://images.example/private.jpg',
            sourceUrl: 'https://example.com/private',
            source: 'example.com',
          },
        ],
        method: 'serpapi',
      })),
    })
    await proxy('image-search.v1', { query: 'private', max_results: 1 })
    await expect(
      proxy('image-fetch.v1', { url: 'https://images.example/private.jpg' }),
    ).rejects.toThrow('retrieval_upstream_error')
  })

  it('honors cancellation while public DNS resolution is pending', async () => {
    let resolveLookup!: (value: readonly { address: string; family: number }[]) => void
    const lookupAddresses = vi.fn(
      () =>
        new Promise<readonly { address: string; family: number }[]>((resolve) => {
          resolveLookup = resolve
        }),
    )
    const proxy = createOfficeLocalSearchProxy({
      fetchWithAuth: vi.fn(),
      lookupAddresses,
      searchImages: vi.fn(async () => ({
        images: [
          {
            title: 'Image',
            imageUrl: 'https://images.example/image.jpg',
            sourceUrl: 'https://example.com/image',
            source: 'example.com',
          },
        ],
        method: 'serpapi',
      })),
    })
    await proxy('image-search.v1', { query: 'image', max_results: 1 })
    const controller = new AbortController()
    const pending = proxy(
      'image-fetch.v1',
      { url: 'https://images.example/image.jpg' },
      controller.signal,
    )
    controller.abort()
    await expect(pending).rejects.toThrow('search_cancelled')
    resolveLookup([{ address: '93.184.216.34', family: 4 }])
  })

  it('times out a stalled image DNS lookup before opening HTTPS', async () => {
    const proxy = createOfficeLocalSearchProxy({
      fetchWithAuth: vi.fn(),
      imageTimeoutMs: 5,
      lookupAddresses: vi.fn(() => new Promise(() => undefined)),
      searchImages: vi.fn(async () => ({
        images: [
          {
            title: 'Stalled',
            imageUrl: 'https://images.example/stalled.jpg',
            sourceUrl: 'https://example.com/stalled',
            source: 'example.com',
          },
        ],
        method: 'serpapi',
      })),
    })
    await proxy('image-search.v1', { query: 'stalled', max_results: 1 })
    await expect(
      proxy('image-fetch.v1', { url: 'https://images.example/stalled.jpg' }),
    ).rejects.toThrow('retrieval_upstream_error')
  })

  it('sends an exact bounded request to the fixed service with PC auth and returns sanitized JSON', async () => {
    const fetch = vi.fn(async (_url: string, init?: RequestInit) => {
      expect(JSON.parse(String(init?.body))).toEqual({
        version: 1,
        operation: 'web-search',
        input: { query: 'office agents', max_results: 3 },
      })
      expect(init?.redirect).toBe('error')
      return new Response(
        JSON.stringify({
          results: [
            { title: 'Result', url: 'https://example.com/paper', snippet: 'Bounded snippet' },
          ],
        }),
        { headers: { 'content-type': 'application/json', 'content-length': '103' } },
      )
    })
    const proxy = createOfficeRetrievalProxy({
      endpoint: TEST_ENDPOINT,
      fetch,
      fetchWithAuth: (request) => request('pc-access-token'),
    })
    const bytes = await proxy('web-search.v1', { query: 'office agents', max_results: 3 })
    expect(JSON.parse(new TextDecoder().decode(bytes))).toEqual({
      results: [{ title: 'Result', url: 'https://example.com/paper', snippet: 'Bounded snippet' }],
    })
    expect(fetch).toHaveBeenCalledWith(
      TEST_ENDPOINT,
      expect.objectContaining({
        headers: expect.objectContaining({ authorization: 'Bearer pc-access-token' }),
      }),
    )
  })

  it.each([
    ['web-fetch.v1', { url: 'http://example.com' }],
    ['web-fetch.v1', { url: 'https://127.0.0.1/private' }],
    ['web-fetch.v1', { url: 'https://169.254.169.254/latest/meta-data' }],
    ['web-fetch.v1', { url: 'https://[::1]/private' }],
    ['web-fetch.v1', { url: 'https://2130706433/private' }],
    ['web-fetch.v1', { url: 'https://0x7f000001/private' }],
    ['web-fetch.v1', { url: 'https://0177.0.0.1/private' }],
    ['web-fetch.v1', { url: 'https://[::ffff:127.0.0.1]/private' }],
    ['web-fetch.v1', { url: 'https://[::7f00:1]/private' }],
    ['web-fetch.v1', { url: 'https://[fc00::1]/private' }],
    ['web-fetch.v1', { url: 'https://[fe80::1]/private' }],
    ['web-fetch.v1', { url: 'https://[fec0::1]/private' }],
    ['web-fetch.v1', { url: 'https://[2001::1]/private' }],
    ['web-fetch.v1', { url: 'https://[2001:20::1]/private' }],
    ['web-fetch.v1', { url: 'https://[2001:db8::1]/private' }],
    ['web-fetch.v1', { url: 'https://[2002:7f00:1::]/private' }],
    ['web-search.v1', { query: 'x', max_results: 21 }],
    ['image-search.v1', { query: 'x', max_results: 0 }],
  ])('rejects invalid or literal-private input for %s', async (capability, body) => {
    const proxy = createOfficeRetrievalProxy({
      endpoint: TEST_ENDPOINT,
      fetchWithAuth: vi.fn(),
    })
    await expect(proxy(capability, body)).rejects.toThrow('retrieval_invalid_request')
  })

  it.each([
    new Response('redirect', { status: 302, headers: { location: 'http://127.0.0.1' } }),
    new Response('<html>no</html>', { headers: { 'content-type': 'text/html' } }),
    new Response('x', {
      headers: { 'content-type': 'application/json', 'content-length': '999999' },
    }),
  ])('maps unsafe upstream responses to a stable error', async (response) => {
    const proxy = createOfficeRetrievalProxy({
      endpoint: TEST_ENDPOINT,
      fetch: vi.fn(async () => response),
      fetchWithAuth: (request) => request('token'),
    })
    await expect(proxy('web-search.v1', { query: 'safe', max_results: 3 })).rejects.toThrow(
      'retrieval_upstream_error',
    )
  })

  it('bounds concurrent upstream work', async () => {
    let release!: () => void
    const blocked = new Promise<void>((resolve) => {
      release = resolve
    })
    const proxy = createOfficeRetrievalProxy({
      endpoint: TEST_ENDPOINT,
      maxConcurrent: 1,
      fetch: vi.fn(async () => {
        await blocked
        return new Response('{"results":[]}', {
          headers: { 'content-type': 'application/json' },
        })
      }),
      fetchWithAuth: (request) => request('token'),
    })
    const first = proxy('web-search.v1', { query: 'first', max_results: 1 })
    await expect(proxy('web-search.v1', { query: 'second', max_results: 1 })).rejects.toThrow(
      'retrieval_busy',
    )
    release()
    await expect(first).resolves.toBeInstanceOf(Uint8Array)
  })
})
