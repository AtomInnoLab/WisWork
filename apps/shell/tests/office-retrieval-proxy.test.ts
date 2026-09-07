import { describe, expect, it, vi } from 'vitest'
import {
  collectBoundedImageBytes,
  createPinnedLookup,
  createOfficeLocalSearchProxy,
  createOfficeRetrievalProxy,
  officeRetrievalEndpointFromEnv,
} from '../src/main/office-retrieval-proxy'

const TEST_ENDPOINT = 'https://retrieval.test.invalid/v1/office/retrieval'
const TEST_SERVICES = {
  [TEST_ENDPOINT]: {
    contract: 'wiswork-office-retrieval-v1',
    ssrfProtection: 'dns-rebinding-and-redirect-hops-v1',
  },
} as const

describe('Office fixed retrieval proxy', () => {
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
    await expect(collectBoundedImageBytes(chunks(), 5)).rejects.toThrow('retrieval_upstream_error')
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
    await expect(pending).rejects.toThrow('retrieval_upstream_error')
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
