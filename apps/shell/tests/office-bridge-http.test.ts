import { EventEmitter } from 'node:events'
import { request as httpRequest } from 'node:http'
import { describe, expect, it, vi } from 'vitest'

import { createOfficeBridge } from '@wiswork/office-bridge'
import {
  startOfficeBridgeHttpServer,
  writeOfficeBridgeResponse,
} from '../src/main/office-bridge-http'

describe('Office bridge HTTP adapter', () => {
  it('rejects denied and oversized raw bodies before invoking the bridge', async () => {
    const bridge = createOfficeBridge({
      allowedOrigin: 'https://office.example.test',
      proxy: vi.fn(),
    })
    const handle = vi.spyOn(bridge, 'handle')
    const server = await startOfficeBridgeHttpServer({
      bridge,
      host: '127.0.0.1',
      port: 0,
      allowedOrigin: 'https://office.example.test',
      maxBodyBytes: 32,
    })
    const raw = (origin: string, body: string) =>
      new Promise<number>((resolve, reject) => {
        const req = httpRequest(
          {
            host: '127.0.0.1',
            port: server.port,
            path: '/v1/office/pairings',
            method: 'POST',
            headers: { origin, 'content-type': 'application/json' },
          },
          (response) => {
            response.resume()
            response.on('end', () => resolve(response.statusCode ?? 0))
          },
        )
        req.on('error', reject)
        req.end(body)
      })
    await expect(raw('https://evil.example', 'x'.repeat(1000))).resolves.toBe(403)
    await expect(raw('https://office.example.test', 'x'.repeat(1000))).resolves.toBe(413)
    expect(handle).not.toHaveBeenCalled()
    await server.stop()
  })

  it('times out a slow incomplete inbound body before bridge allocation', async () => {
    const bridge = createOfficeBridge({
      allowedOrigin: 'https://office.example.test',
      proxy: vi.fn(),
    })
    const handle = vi.spyOn(bridge, 'handle')
    const server = await startOfficeBridgeHttpServer({
      bridge,
      host: '127.0.0.1',
      port: 0,
      allowedOrigin: 'https://office.example.test',
      inboundTimeoutMs: 20,
    })
    const status = await new Promise<number>((resolve, reject) => {
      const req = httpRequest(
        {
          host: '127.0.0.1',
          port: server.port,
          path: '/v1/office/pairings',
          method: 'POST',
          headers: { origin: 'https://office.example.test', 'content-type': 'application/json' },
        },
        (response) => {
          response.resume()
          response.on('end', () => resolve(response.statusCode ?? 0))
        },
      )
      req.on('error', reject)
      req.write('{')
    })
    expect(status).toBe(408)
    expect(handle).not.toHaveBeenCalled()
    await server.stop()
  })

  it('cancels a streaming response immediately when a backpressured client closes', async () => {
    let cancelled = false
    const body = new ReadableStream<Uint8Array>({
      pull(controller) {
        controller.enqueue(new Uint8Array([1]))
      },
      cancel() {
        cancelled = true
      },
    })
    class FakeResponse extends EventEmitter {
      statusCode = 0
      headersSent = false
      destroyed = false
      setHeader() {}
      write() {
        queueMicrotask(() => {
          this.destroyed = true
          this.emit('close')
        })
        return false
      }
      end() {}
      destroy() {
        this.destroyed = true
      }
    }
    const outgoing = new FakeResponse()
    const abort = vi.fn()
    await writeOfficeBridgeResponse(new Response(body), outgoing, abort)
    expect(abort).toHaveBeenCalledOnce()
    expect(cancelled).toBe(true)
  })

  it('aborts bridge response streaming when the HTTP client disconnects', async () => {
    let observeAbort!: () => void
    const aborted = new Promise<void>((resolve) => (observeAbort = resolve))
    const bridge = {
      handle: vi.fn(async (request: Request) => {
        request.signal.addEventListener('abort', observeAbort, { once: true })
        return new Response(
          new ReadableStream<Uint8Array>({
            pull(controller) {
              controller.enqueue(new Uint8Array(64 * 1024))
            },
          }),
        )
      }),
      listPending: () => [],
      approve: () => false,
      reject: () => false,
      revokeAll() {},
      setSessionAvailable() {},
      shutdown() {},
    }
    const server = await startOfficeBridgeHttpServer({
      bridge,
      host: '127.0.0.1',
      port: 0,
      allowedOrigin: 'https://office.example.test',
    })
    await new Promise<void>((resolve, reject) => {
      const req = httpRequest(
        {
          host: '127.0.0.1',
          port: server.port,
          path: '/v1/office/messages',
          method: 'POST',
          headers: { origin: 'https://office.example.test', 'content-type': 'application/json' },
        },
        (response) => {
          response.once('data', () => {
            response.destroy()
            resolve()
          })
        },
      )
      req.on('error', reject)
      req.end()
    })
    await aborted
    await server.stop()
  })

  it('binds only numeric loopback and forwards Fetch requests', async () => {
    const bridge = createOfficeBridge({
      allowedOrigin: 'https://office.example.test',
      proxy: vi.fn(),
    })
    const onPending = vi.fn()
    await expect(
      startOfficeBridgeHttpServer({
        bridge,
        host: '0.0.0.0',
        port: 43127,
        allowedOrigin: 'https://office.example.test',
      }),
    ).rejects.toThrow('loopback_host_required')

    const server = await startOfficeBridgeHttpServer({
      bridge,
      host: '127.0.0.1',
      port: 0,
      allowedOrigin: 'https://office.example.test',
      onPending,
    })
    const response = await new Promise<{ status: number; body: string }>((resolve, reject) => {
      const req = httpRequest(
        {
          host: '127.0.0.1',
          port: server.port,
          path: '/v1/office/pairings',
          method: 'POST',
          headers: {
            origin: 'https://office.example.test',
            'content-type': 'application/json',
          },
        },
        (res) => {
          const chunks: Buffer[] = []
          res.on('data', (chunk: Buffer) => chunks.push(chunk))
          res.on('end', () =>
            resolve({ status: res.statusCode ?? 0, body: Buffer.concat(chunks).toString() }),
          )
        },
      )
      req.on('error', reject)
      req.end(JSON.stringify({ host_label: 'Word' }))
    })
    expect(response.status).toBe(202)
    expect(JSON.parse(response.body)).toMatchObject({ expires_in: 120 })
    expect(onPending).toHaveBeenCalledWith(
      expect.objectContaining({ hostLabel: 'Word', origin: 'https://office.example.test' }),
    )
    await server.stop()
  })
})
