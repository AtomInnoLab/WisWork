import { createWriteStream } from 'node:fs'
import { Transform } from 'node:stream'
import { pipeline } from 'node:stream/promises'
import type { BundleDownload } from './bundle-installer.js'

const INITIAL_BUNDLE_HOSTS = new Set(['relay.fullyjustified.net', 'data1.fullyjustified.net'])

async function fetchPinned(
  url: string,
  init: RequestInit,
  allowedRedirectHosts: ReadonlySet<string>,
  fetchImplementation: typeof fetch,
  redirects = 0,
): Promise<Response> {
  if (redirects > 5) throw new Error('Bundle download exceeded redirect limit')
  const parsed = new URL(url)
  if (parsed.protocol !== 'https:') throw new Error('Bundle download requires HTTPS')
  if (redirects === 0 && !INITIAL_BUNDLE_HOSTS.has(parsed.hostname)) {
    throw new Error('Bundle download host is not approved')
  }
  const response = await fetchImplementation(parsed.href, { ...init, redirect: 'manual' })
  if (response.status >= 300 && response.status < 400) {
    const location = response.headers.get('location')
    if (!location) throw new Error('Bundle redirect is missing Location')
    const next = new URL(location, parsed)
    if (next.protocol !== 'https:' || !allowedRedirectHosts.has(next.hostname)) {
      throw new Error('Bundle redirect host is not approved')
    }
    return fetchPinned(next.href, init, allowedRedirectHosts, fetchImplementation, redirects + 1)
  }
  if (!response.ok || !response.body) {
    throw new Error(`Bundle download failed with HTTP ${response.status}`)
  }
  return response
}

export function createHttpBundleDownload(
  fetchImplementation: typeof fetch = globalThis.fetch,
): BundleDownload {
  return async ({
    url,
    destination,
    offset,
    expectedBytes,
    signal,
    onBytes,
    allowedRedirectHosts,
  }) => {
    const headers = offset > 0 ? { Range: `bytes=${offset}-` } : undefined
    const response = await fetchPinned(
      url,
      { headers, signal },
      new Set(allowedRedirectHosts),
      fetchImplementation,
    )
    let append = offset > 0
    if (append && response.status === 206) {
      const contentRange = response.headers.get('content-range')
      if (!contentRange?.startsWith(`bytes ${offset}-`)) {
        throw new Error('Bundle range response does not match requested offset')
      }
    } else if (append && response.status === 200) {
      append = false
    } else if (response.status !== 200 && response.status !== 206) {
      throw new Error(`Bundle download returned unexpected HTTP ${response.status}`)
    }

    let receivedBytes = 0
    const baseBytes = append ? offset : 0
    const counter = new Transform({
      transform(chunk, _encoding, callback) {
        receivedBytes += Buffer.byteLength(chunk)
        if (baseBytes + receivedBytes > expectedBytes) {
          callback(new Error('Bundle download exceeded expected size'))
          return
        }
        onBytes(receivedBytes)
        callback(null, chunk)
      },
    })
    await pipeline(
      response.body as never,
      counter,
      createWriteStream(destination, { flags: append ? 'a' : 'w' }),
      { signal },
    )
  }
}
