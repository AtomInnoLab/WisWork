import { afterEach, describe, expect, it, vi } from 'vitest'
import { deployedBuildId, resolveBuildVersion } from '../src/build-version.js'

vi.stubGlobal('location', { origin: 'https://office.example' })
afterEach(() => vi.useRealTimers())

describe('deployedBuildId', () => {
  it('reads a valid same-origin build without cache', async () => {
    const fetcher = vi.fn(
      async () =>
        new Response(
          JSON.stringify({
            buildId: 'release_123',
            presentationMinPcProtocol: 2,
            presentationMinRelayProtocol: 2,
          }),
        ),
    )
    expect(await deployedBuildId(fetcher as typeof fetch)).toBe('release_123')
    expect(fetcher).toHaveBeenCalledWith(new URL('https://office.example/version.json'), {
      cache: 'no-store',
      credentials: 'same-origin',
      signal: expect.any(AbortSignal),
    })
  })

  it.each([{}, { buildId: '../other' }, { buildId: 123 }, [], null])(
    'ignores invalid metadata: %j',
    async (payload) => {
      expect(
        await deployedBuildId(async () => new Response(JSON.stringify(payload))),
      ).toBeUndefined()
    },
  )

  it.each([
    { buildId: 'release_123' },
    { buildId: 'release_123', presentationMinPcProtocol: 1, presentationMinRelayProtocol: 2 },
    { buildId: 'release_123', presentationMinPcProtocol: 2, presentationMinRelayProtocol: 1 },
  ])('rejects missing or incompatible PPT protocol metadata: %j', async (payload) => {
    expect(await deployedBuildId(async () => new Response(JSON.stringify(payload)))).toBeUndefined()
  })

  it('allows older deployments without metadata', async () => {
    expect(await deployedBuildId(async () => new Response('', { status: 404 }))).toBeUndefined()
  })

  it('falls back when fetch never settles', async () => {
    vi.useFakeTimers()
    const result = deployedBuildId(() => new Promise(() => {}))
    await vi.advanceTimersByTimeAsync(5_000)
    await expect(result).resolves.toBeUndefined()
  })

  it('falls back when response JSON never settles', async () => {
    vi.useFakeTimers()
    const result = deployedBuildId(
      async () =>
        ({
          ok: true,
          json: () => new Promise(() => {}),
        }) as unknown as Response,
    )
    await vi.advanceTimersByTimeAsync(5_000)
    await expect(result).resolves.toBeUndefined()
  })
})

describe('resolveBuildVersion', () => {
  it('blocks production document tools when deployed metadata cannot be verified', () => {
    expect(resolveBuildVersion(undefined, 'release_123', true)).toEqual({
      status: 'unavailable',
    })
    expect(resolveBuildVersion('release_123', 'release_123', true)).toEqual({
      status: 'current',
    })
    expect(resolveBuildVersion('release_124', 'release_123', true)).toEqual({
      status: 'stale',
      buildId: 'release_124',
    })
  })

  it('treats sentinel-like build IDs as real deployed versions', () => {
    expect(resolveBuildVersion('checking', 'previous')).toEqual({
      status: 'stale',
      buildId: 'checking',
    })
    expect(resolveBuildVersion('current', 'previous')).toEqual({
      status: 'stale',
      buildId: 'current',
    })
    expect(resolveBuildVersion('checking', 'checking')).toEqual({ status: 'current' })
    expect(resolveBuildVersion(undefined, 'current')).toEqual({ status: 'current' })
  })
})
