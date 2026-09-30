import { createServer } from 'node:http'
import { mkdtemp, readdir, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { afterEach, expect, it, vi } from 'vitest'
import WebSocket, { WebSocketServer } from 'ws'
import pngjs from 'pngjs'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { createPresentationService } from '../src/main/presentation-service.ts'
import { createOfficeRelayClient } from '../src/main/office-relay-client.ts'
import { inspectPcBusiness } from '../../../tools/ppt-agent-pc-business-smoke.mjs'

const cleanup = []
afterEach(async () => {
  for (const dispose of cleanup.splice(0).reverse()) await dispose()
})

it('runs release smoke against actual PC presentation and attachment services', async () => {
  const userDataPath = await mkdtemp(join(tmpdir(), 'ppt-pc-business-smoke-'))
  cleanup.push(() => rm(userDataPath, { recursive: true, force: true }))
  // Electron's native decoder is unavailable to Vitest; keep the real PC storage and
  // upload protocol while decoding this generated PNG fixture with pngjs.
  const presentation = createPresentationService({
    userDataPath,
    compile: compilePresentationDeck,
    normalizeImage: async (raw) => {
      const png = pngjs.PNG.sync.read(Buffer.from(raw))
      return { bytes: raw, width: png.width, height: png.height }
    },
  })
  const signal = new AbortController().signal
  const documentId = 'document-1'
  const projectId = 'project-1'
  const deck = {
    version: 1,
    id: projectId,
    title: 'Release smoke project',
    style: { fontFace: 'Arial', background: 'FFFFFF', textColor: '111111', accentColor: '3366FF' },
    assets: [],
    claims: [],
    slides: [
      {
        id: 'slide-1',
        title: 'One',
        elements: [{ id: 'title', kind: 'text', text: 'Release smoke', x: 1, y: 1, w: 8, h: 1 }],
      },
    ],
  }
  const compiled = JSON.parse(
    Buffer.from(
      await presentation({ operation: 'compile', documentId, requestId: 'run-1', deck }, signal),
    ).toString('utf8'),
  )
  expect(compiled.status).toBe('compiled')
  const server = createServer()
  const websocket = new WebSocketServer({ server })
  cleanup.push(async () => {
    for (const socket of websocket.clients) socket.terminate()
    await new Promise((resolve) => websocket.close(resolve))
    await new Promise((resolve) => server.close(resolve))
  })
  let officeSocket
  let pcSocket
  let client
  websocket.on('connection', (socket, request) => {
    if (request.headers.origin) officeSocket = socket
    else pcSocket = socket
    socket.on('message', async (bytes) => {
      const frame = JSON.parse(bytes.toString())
      if (frame.type === 'office.create') {
        expect(request.headers.origin).toBe('https://office.8-216-134-194.sslip.io')
        socket.send(
          JSON.stringify({
            version: 2,
            type: 'office.created',
            pairing_id: 'pairing_12345678',
            verification_code: '123456',
            expires_in: 120,
          }),
        )
        return
      }
      if (frame.type === 'pc.negotiate') {
        expect(frame.capabilities).toContain('presentation-assets.v1')
        socket.send(
          JSON.stringify({
            version: 2,
            type: 'pc.negotiated',
            pairing_version: 2,
            capabilities: ['presentation.v1', 'presentation-assets.v1'],
          }),
        )
        return
      }
      if (frame.type === 'pc.claim') {
        socket.send(
          JSON.stringify({
            version: 2,
            type: 'pc.claimed',
            pairing_id: 'pairing_12345678',
            host: 'PowerPoint',
            origin: 'https://office.8-216-134-194.sslip.io',
            verification_code: '123456',
            expires_in: 120,
            capabilities: ['presentation.v1', 'presentation-assets.v1'],
          }),
        )
        return
      }
      if (frame.type === 'pc.approve') {
        expect(frame.pairing_id).toBe('pairing_12345678')
        socket.send(
          JSON.stringify({
            version: 2,
            type: 'pc.approved',
            session_id: 'session_12345678',
            capability: 'pc_capability_12345678',
            capabilities: ['presentation.v1', 'presentation-assets.v1'],
            expires_in: 1800,
          }),
        )
        await vi.waitFor(() => expect(client.status()).toBe('paired'))
        officeSocket.send(
          JSON.stringify({
            version: 2,
            type: 'office.approved',
            session_id: 'session_12345678',
            capability: 'office_capability_12345678',
            capabilities: ['presentation.v1', 'presentation-assets.v1'],
            expires_in: 1800,
          }),
        )
        return
      }
      if (frame.type === 'office.request') {
        pcSocket.send(
          JSON.stringify({
            version: 2,
            type: 'relay.request',
            session_id: frame.session_id,
            request_id: frame.request_id,
            capability_name: frame.capability_name,
            body: frame.body,
          }),
        )
        return
      }
      if (['pc.start', 'pc.chunk', 'pc.done'].includes(frame.type)) {
        const { capability: _credential, ...forwarded } = frame
        officeSocket.send(
          JSON.stringify({ ...forwarded, type: frame.type.replace('pc.', 'relay.') }),
        )
      }
    })
  })
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve))
  const endpoint = `ws://127.0.0.1:${server.address().port}/office-relay`
  client = createOfficeRelayClient({
    endpoint,
    connect: (_url, token) =>
      new WebSocket(endpoint, { headers: { Authorization: `Bearer ${token}` } }),
    getValidAccountStatus: async () => ({ loggedIn: true }),
    getAccessToken: async () => 'test-token',
    proxy: async () => ({ status: 200, body: new Uint8Array() }),
    presentationProxy: (body, requestSignal) => presentation(body, requestSignal),
    onPending: (pending) => {
      void client.approve(pending.pairingId)
    },
  })
  cleanup.push(async () => client.revoke())
  let claim
  const result = await inspectPcBusiness(
    `http://127.0.0.1:${server.address().port}`,
    documentId,
    projectId,
    {
      onCode: (code) => {
        claim = client.claim(code)
      },
      timeoutMs: 1000,
      uploadFixtures: true,
    },
  )
  await claim
  expect(result).toEqual({
    projectId,
    status: 'compiled',
    slideCount: 1,
    attachmentPageCount: 0,
    textChecked: true,
    imageChecked: true,
    uploadChecked: true,
  })
  const documentDir = join(
    userDataPath,
    'presentation-attachments',
    createHash('sha256').update(documentId).digest('hex'),
  )
  expect(await readdir(documentDir)).toEqual([])
})
