import { EventEmitter } from 'node:events'
import { request as httpRequest } from 'node:http'
import { describe, expect, it, vi } from 'vitest'

import { createOfficeBridge } from '@wiswork/office-bridge'
import {
  startOfficeBridgeHttpServer,
  writeOfficeBridgeResponse,
} from '../src/main/office-bridge-http'

describe('Office bridge HTTP adapter', () => {
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
      shutdown() {},
    }
    const server = await startOfficeBridgeHttpServer({ bridge, host: '127.0.0.1', port: 0 })
    await new Promise<void>((resolve, reject) => {
      const req = httpRequest(
        { host: '127.0.0.1', port: server.port, path: '/stream' },
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
      startOfficeBridgeHttpServer({ bridge, host: '0.0.0.0', port: 43127 }),
    ).rejects.toThrow('loopback_host_required')

    const server = await startOfficeBridgeHttpServer({
      bridge,
      host: '127.0.0.1',
      port: 0,
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
