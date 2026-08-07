import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createHttpBundleDownload } from '../src/download.js'

describe('createHttpBundleDownload', () => {
  const roots: string[] = []

  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  async function destination(initial = '') {
    const root = await mkdtemp(join(tmpdir(), 'latex-bundle-download-'))
    roots.push(root)
    const path = join(root, 'bundle.part')
    await writeFile(path, initial)
    return path
  }

  it('follows only approved redirects and resumes with a validated byte range', async () => {
    const path = await destination('abc')
    const fetchImplementation = vi
      .fn<typeof fetch>()
      .mockResolvedValueOnce(
        new Response(null, {
          status: 302,
          headers: { location: 'https://data1.fullyjustified.net/fixed.tar' },
        }),
      )
      .mockResolvedValueOnce(
        new Response('def', {
          status: 206,
          headers: { 'content-range': 'bytes 3-5/6' },
        }),
      )
    const onBytes = vi.fn()
    const download = createHttpBundleDownload(fetchImplementation)
    await download({
      url: 'https://relay.fullyjustified.net/default_bundle_v33.tar',
      destination: path,
      offset: 3,
      expectedBytes: 6,
      signal: new AbortController().signal,
      onBytes,
      allowedRedirectHosts: ['data1.fullyjustified.net'],
    })

    expect(await readFile(path, 'utf8')).toBe('abcdef')
    expect(fetchImplementation).toHaveBeenNthCalledWith(
      2,
      'https://data1.fullyjustified.net/fixed.tar',
      expect.objectContaining({
        headers: { Range: 'bytes=3-' },
        redirect: 'manual',
      }),
    )
    expect(onBytes).toHaveBeenLastCalledWith(3)
  })

  it('rejects an unapproved redirect without writing response bytes', async () => {
    const path = await destination()
    const download = createHttpBundleDownload(
      vi.fn<typeof fetch>().mockResolvedValue(
        new Response(null, {
          status: 302,
          headers: { location: 'https://evil.example/bundle.tar' },
        }),
      ),
    )
    await expect(
      download({
        url: 'https://relay.fullyjustified.net/default_bundle_v33.tar',
        destination: path,
        offset: 0,
        expectedBytes: 6,
        signal: new AbortController().signal,
        onBytes: vi.fn(),
        allowedRedirectHosts: ['data1.fullyjustified.net'],
      }),
    ).rejects.toThrow(/redirect host/i)
    expect(await readFile(path, 'utf8')).toBe('')
  })

  it('safely restarts from zero when a server ignores the Range request', async () => {
    const path = await destination('stale')
    const download = createHttpBundleDownload(
      vi.fn<typeof fetch>().mockResolvedValue(new Response('fresh', { status: 200 })),
    )
    await download({
      url: 'https://data1.fullyjustified.net/fixed.tar',
      destination: path,
      offset: 5,
      expectedBytes: 5,
      signal: new AbortController().signal,
      onBytes: vi.fn(),
      allowedRedirectHosts: ['data1.fullyjustified.net'],
    })
    expect(await readFile(path, 'utf8')).toBe('fresh')
  })

  it('aborts before writing beyond the pinned size', async () => {
    const path = await destination()
    const download = createHttpBundleDownload(
      vi.fn<typeof fetch>().mockResolvedValue(new Response('too-large', { status: 200 })),
    )
    await expect(
      download({
        url: 'https://data1.fullyjustified.net/fixed.tar',
        destination: path,
        offset: 0,
        expectedBytes: 3,
        signal: new AbortController().signal,
        onBytes: vi.fn(),
        allowedRedirectHosts: ['data1.fullyjustified.net'],
      }),
    ).rejects.toThrow(/expected size/i)
    expect((await readFile(path)).byteLength).toBeLessThanOrEqual(3)
  })
})
