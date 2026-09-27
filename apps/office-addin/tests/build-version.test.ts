import { afterEach, describe, expect, it, vi } from 'vitest'
import { deployedBuildId, resolveBuildVersion } from '../src/build-version.js'

vi.stubGlobal('location', { origin: 'https://office.example' })
afterEach(() => vi.useRealTimers())

describe('deployedBuildId', () => {
  it('reads a valid same-origin build without cache', async () => {
    const fetcher = vi.fn(async () => new Response(JSON.stringify({ buildId: 'release_123' })))
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
  it('treats sentinel-like build IDs as real deployed versions', () => {
    expect(resolveBuildVersion('checking', 'previous')).toEqual({
      status: 'stale', buildId: 'checking',
    })
    expect(resolveBuildVersion('current', 'previous')).toEqual({
      status: 'stale', buildId: 'current',
    })
    expect(resolveBuildVersion('checking', 'checking')).toEqual({ status: 'current' })
    expect(resolveBuildVersion(undefined, 'current')).toEqual({ status: 'current' })
  })
})
