import { describe, expect, it, vi } from 'vitest'
import {
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
    ['web-fetch.v1', { url: 'https://[fc00::1]/private' }],
    ['web-fetch.v1', { url: 'https://[fe80::1]/private' }],
    ['web-fetch.v1', { url: 'https://[2001:db8::1]/private' }],
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
