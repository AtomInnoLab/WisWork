import { EventEmitter } from 'node:events'
import type { LookupAddress } from 'node:dns'
import type { IncomingMessage } from 'node:http'
import { request as httpsRequest } from 'node:https'
import type { LookupFunction } from 'node:net'
import { Readable } from 'node:stream'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import {
  createOfficeLocalSearchProxy,
  createPinnedLookup,
} from '../src/main/office-retrieval-proxy'

vi.mock('node:https', () => ({ request: vi.fn() }))

const IMAGE_URL = 'https://images.example/cover.png'
const PUBLIC_ADDRESSES = [
  { address: '2606:4700:4700::1111', family: 6 },
  { address: '1.1.1.1', family: 4 },
]
const IMAGE_BYTES = Buffer.from([0x89, 0x50, 0x4e, 0x47])

function retrieval(addresses = PUBLIC_ADDRESSES) {
  return createOfficeLocalSearchProxy({
    fetchWithAuth: vi.fn(),
    lookupAddresses: async () => addresses,
    searchImages: async () => ({
      images: [
        {
          title: 'Cover',
          imageUrl: IMAGE_URL,
          sourceUrl: 'https://example.com/cover',
          source: 'example.com',
        },
      ],
      method: 'test',
    }),
  })
}

function mockImageConnection(firstAddressUnavailable = false) {
  const attempts: string[] = []
  const offered: LookupAddress[] = []
  let connection!: { lookup: LookupFunction; autoSelectFamily?: boolean }
  vi.mocked(httpsRequest).mockImplementation(((
    url: URL,
    options: typeof connection,
    respond: (response: IncomingMessage) => void,
  ) => {
    connection = options
    const request = Object.assign(new EventEmitter(), {
      end: () =>
        queueMicrotask(() =>
          options.lookup(url.hostname, { all: true }, (error, addresses) => {
            if (error) return request.destroy(error)
            if (!Array.isArray(addresses))
              return request.destroy(new Error('expected all pinned addresses'))
            offered.push(...addresses)
            // Model Node's family selection using only the actual pinned lookup result.
            // No socket is opened, and the mock cannot invent the missing IPv4 candidate.
            const candidates =
              options.autoSelectFamily !== false ? addresses : addresses.slice(0, 1)
            for (const candidate of candidates) {
              attempts.push(candidate.address)
              if (firstAddressUnavailable && candidate.address === PUBLIC_ADDRESSES[0]!.address)
                continue
              const response = Object.assign(Readable.from([IMAGE_BYTES]), {
                statusCode: 200,
                headers: { 'content-type': 'image/png' },
              })
              response.once('close', () => request.emit('close'))
              respond(response as unknown as IncomingMessage)
              return
            }
            request.destroy(
              Object.assign(new Error('simulated unreachable address'), {
                code: 'ENETUNREACH',
              }),
            )
          }),
        ),
      destroy: (error: Error) => {
        request.emit('error', error)
        request.emit('close')
      },
    })
    return request
  }) as typeof httpsRequest)
  return { attempts, offered, connection: () => connection }
}

describe('Office image connection family fallback', () => {
  beforeEach(() => vi.resetAllMocks())

  it('snapshots the DNS answers before callers can change the pinned targets', async () => {
    expect(() => createPinnedLookup([])).toThrow('retrieval_upstream_error')
    const addresses = PUBLIC_ADDRESSES.map((entry) => ({ ...entry }))
    const lookup = createPinnedLookup(addresses)
    addresses[0]!.address = '127.0.0.1'
    addresses.push({ address: '127.0.0.2', family: 4 })
    const pinned = await new Promise((resolve, reject) =>
      lookup('images.example', { all: true }, (error, entries) =>
        error ? reject(error) : resolve(entries),
      ),
    )
    expect(pinned).toEqual(PUBLIC_ADDRESSES)
  })

  it('downloads through the second public family when the first address is unavailable', async () => {
    const network = mockImageConnection(true)
    const proxy = retrieval()
    await proxy('image-search.v1', { query: 'cover', max_results: 1 })
    const result = await proxy('image-fetch.v1', { url: IMAGE_URL })
    expect(JSON.parse(new TextDecoder().decode(result))).toEqual({
      mime: 'image/png',
      data_base64: IMAGE_BYTES.toString('base64'),
    })
    expect(network.attempts).toEqual(PUBLIC_ADDRESSES.map(({ address }) => address))
    expect(httpsRequest).toHaveBeenCalledOnce()
  })

  it('enables native family selection and retains every validated address in the pinned lookup', async () => {
    const network = mockImageConnection()
    const proxy = retrieval()
    await proxy('image-search.v1', { query: 'cover', max_results: 1 })
    await proxy('image-fetch.v1', { url: IMAGE_URL })
    expect(httpsRequest).toHaveBeenCalledWith(
      new URL(IMAGE_URL),
      expect.objectContaining({ agent: false }),
      expect.any(Function),
    )
    expect.soft(network.connection().autoSelectFamily).toBe(true)
    expect.soft(network.offered).toEqual(PUBLIC_ADDRESSES)
    const scalar = await new Promise((resolve, reject) =>
      network
        .connection()
        .lookup('images.example', { all: false }, (error, address, family) =>
          error ? reject(error) : resolve({ address, family }),
        ),
    )
    expect(scalar).toEqual(PUBLIC_ADDRESSES[0])
  })

  it.each([0, 1])('rejects all connections when DNS answer %i is private', async (privateIndex) => {
    mockImageConnection()
    const addresses = PUBLIC_ADDRESSES.map((entry, index) =>
      index === privateIndex ? { address: '127.0.0.1', family: 4 } : entry,
    )
    const proxy = retrieval(addresses)
    await proxy('image-search.v1', { query: 'cover', max_results: 1 })
    await expect(proxy('image-fetch.v1', { url: IMAGE_URL })).rejects.toThrow(
      'retrieval_upstream_error',
    )
    expect(httpsRequest).not.toHaveBeenCalled()
  })
})
