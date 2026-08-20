import { createServer, type IncomingMessage, type ServerResponse } from 'node:http'

import { assertLoopbackHost, type OfficeBridge, type PendingPairing } from '@wiswork/office-bridge'

export interface OfficeBridgeHttpServer {
  port: number
  stop(): Promise<void>
}

function requestBody(request: IncomingMessage): ReadableStream<Uint8Array> | undefined {
  if (request.method === 'GET' || request.method === 'HEAD' || request.method === 'OPTIONS') return
  return new ReadableStream({
    start(controller) {
      request.on('data', (chunk: Buffer) => controller.enqueue(new Uint8Array(chunk)))
      request.on('end', () => controller.close())
      request.on('error', (error) => controller.error(error))
    },
    cancel() {
      request.destroy()
    },
  })
}

async function writeResponse(response: Response, target: ServerResponse): Promise<void> {
  target.statusCode = response.status
  response.headers.forEach((value, name) => target.setHeader(name, value))
  if (!response.body) return void target.end()
  const reader = response.body.getReader()
  try {
    while (true) {
      const result = await reader.read()
      if (result.done) break
      if (!target.write(result.value))
        await new Promise<void>((resolve) => target.once('drain', resolve))
    }
    target.end()
  } catch {
    target.destroy()
  }
}

export async function startOfficeBridgeHttpServer(options: {
  bridge: OfficeBridge
  host: string
  port: number
  onPending?: (pairing: PendingPairing) => void
}): Promise<OfficeBridgeHttpServer> {
  const host = assertLoopbackHost(options.host)
  if (!Number.isSafeInteger(options.port) || options.port < 0 || options.port > 65_535)
    throw new Error('invalid_office_bridge_port')
  const announced = new Set<string>()
  const server = createServer(async (incoming, outgoing) => {
    const controller = new AbortController()
    incoming.on('aborted', () => controller.abort())
    try {
      const headers = new Headers()
      for (const [name, value] of Object.entries(incoming.headers)) {
        if (Array.isArray(value)) value.forEach((item) => headers.append(name, item))
        else if (value !== undefined) headers.set(name, value)
      }
      const body = requestBody(incoming)
      const request = new Request(`http://${host}:${options.port}${incoming.url ?? '/'}`, {
        method: incoming.method,
        headers,
        body,
        signal: controller.signal,
        ...(body ? { duplex: 'half' } : {}),
      } as RequestInit & { duplex?: 'half' })
      const response = await options.bridge.handle(request)
      for (const pairing of options.bridge.listPending()) {
        if (!announced.has(pairing.pairingId)) {
          announced.add(pairing.pairingId)
          options.onPending?.(pairing)
        }
      }
      await writeResponse(response, outgoing)
    } catch {
      if (!outgoing.headersSent) {
        outgoing.statusCode = 500
        outgoing.setHeader('content-type', 'application/json; charset=utf-8')
        outgoing.end('{"error":"bridge_failure"}')
      } else outgoing.destroy()
    }
  })
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject)
    server.listen(options.port, host, () => {
      server.off('error', reject)
      resolve()
    })
  })
  const address = server.address()
  if (!address || typeof address === 'string') throw new Error('bridge_start_failed')
  return {
    port: address.port,
    stop: () => new Promise<void>((resolve) => server.close(() => resolve())),
  }
}
