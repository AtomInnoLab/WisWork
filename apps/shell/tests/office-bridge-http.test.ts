import { request as httpRequest } from 'node:http'
import { describe, expect, it, vi } from 'vitest'

import { createOfficeBridge } from '@wiswork/office-bridge'
import { startOfficeBridgeHttpServer } from '../src/main/office-bridge-http'

describe('Office bridge HTTP adapter', () => {
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
