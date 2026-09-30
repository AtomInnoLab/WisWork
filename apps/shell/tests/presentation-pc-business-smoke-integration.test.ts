import { createServer } from 'node:http'
import { mkdtemp, readdir, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { afterEach, expect, it } from 'vitest'
import { WebSocketServer } from 'ws'
import pngjs from 'pngjs'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { createPresentationService } from '../src/main/presentation-service.ts'
import { createPresentationAttachmentService } from '../src/main/presentation-attachments.ts'
import { inspectPcBusiness } from '../../../tools/ppt-agent-pc-business-smoke.mjs'

const cleanup = []
afterEach(async () => {
  await Promise.all(cleanup.splice(0).map((dispose) => dispose()))
})

it('runs release smoke against actual PC presentation and attachment services', async () => {
  const userDataPath = await mkdtemp(join(tmpdir(), 'ppt-pc-business-smoke-'))
  cleanup.push(() => rm(userDataPath, { recursive: true, force: true }))
  const presentation = createPresentationService({ userDataPath, compile: compilePresentationDeck })
  // Electron's native decoder is unavailable to Vitest; keep the real PC storage and
  // upload protocol while decoding this generated PNG fixture with pngjs.
  const attachments = createPresentationAttachmentService({
    userDataPath,
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
  websocket.on('connection', (socket) => {
    socket.on('message', async (bytes) => {
      const frame = JSON.parse(bytes.toString())
      if (frame.type === 'office.create') {
        socket.send(
          JSON.stringify({
            version: 2,
            type: 'office.created',
            pairing_id: 'pair',
            verification_code: '123456',
            expires_in: 120,
          }),
        )
        socket.send(
          JSON.stringify({
            version: 2,
            type: 'office.approved',
            session_id: 'session',
            capability: 'office-capability',
            capabilities: ['presentation.v1', 'presentation-assets.v1'],
            expires_in: 1800,
          }),
        )
        return
      }
      if (frame.type !== 'office.request') return
      let value
      try {
        value =
          frame.capability_name === 'presentation.v1'
            ? JSON.parse(Buffer.from(await presentation(frame.body, signal)).toString('utf8'))
            : await attachments(frame.body, signal)
      } catch (error) {
        value = { error: error.message }
      }
      const common = { version: 2, session_id: frame.session_id, request_id: frame.request_id }
      socket.send(
        JSON.stringify({
          ...common,
          type: 'relay.start',
          status: 200,
          content_type: 'application/json',
        }),
      )
      socket.send(
        JSON.stringify({
          ...common,
          type: 'relay.chunk',
          sequence: 0,
          data: Buffer.from(JSON.stringify(value)).toString('base64'),
        }),
      )
      socket.send(JSON.stringify({ ...common, type: 'relay.done' }))
    })
  })
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve))
  const result = await inspectPcBusiness(
    `http://127.0.0.1:${server.address().port}`,
    documentId,
    projectId,
    { onCode: () => {}, timeoutMs: 1000, uploadFixtures: true },
  )
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
