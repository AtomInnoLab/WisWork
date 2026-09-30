import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { createServer } from 'node:http'
import { test } from 'node:test'
import { WebSocketServer } from 'ws'
import { checkPcStatusResponse, inspectPcBusiness } from './ppt-agent-pc-business-smoke.mjs'

async function fakeRelay(t, response, mutate = (frame) => frame) {
  const server = createServer()
  const sockets = new WebSocketServer({ server })
  t.after(async () => {
    for (const client of sockets.clients) client.terminate()
    await new Promise((resolve) => sockets.close(resolve))
    await new Promise((resolve) => server.close(resolve))
  })
  sockets.on('connection', (socket) => {
    socket.on('message', (data) => {
      const frame = JSON.parse(data.toString())
      if (frame.type === 'office.create') {
        assert.deepEqual(frame.capabilities, ['presentation.v1', 'presentation-assets.v1'])
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
      } else if (frame.type === 'office.request') {
        assert.equal(frame.body.documentId, 'document-1')
        if (frame.body.operation === 'status') assert.equal(frame.body.projectId, 'project-1')
        const value =
          frame.body.operation === 'status'
            ? response
            : frame.body.operation === 'attachment_list_assets'
              ? { attachments: [] }
              : frame.body.operation === 'attachment_metadata'
                ? response.attachment_metadata?.[frame.body.attachmentId]
                : response[frame.body.operation]
        assert.notEqual(value, undefined)
        const common = { version: 2, session_id: frame.session_id, request_id: frame.request_id }
        const frames = [
          { ...common, type: 'relay.start', status: 200, content_type: 'application/json' },
          {
            ...common,
            type: 'relay.chunk',
            sequence: 0,
            data: Buffer.from(JSON.stringify(value)).toString('base64'),
          },
          { ...common, type: 'relay.done' },
        ]
        for (const reply of frames) socket.send(JSON.stringify(mutate(reply)))
      }
    })
  })
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve))
  return `http://127.0.0.1:${server.address().port}`
}

test('real PC smoke pairs and verifies a buffered, chunked status response', async (t) => {
  const origin = await fakeRelay(t, {
    projectId: 'project-1',
    status: 'compiled',
    slideCount: 1,
    slides: [{ id: 'slide-1', title: 'Title' }],
    history: [],
  })
  const codes = []
  assert.deepEqual(
    await inspectPcBusiness(origin, 'document-1', 'project-1', {
      onCode: (code) => codes.push(code),
      timeoutMs: 1000,
    }),
    {
      projectId: 'project-1',
      status: 'compiled',
      slideCount: 1,
      attachmentPageCount: 0,
      textChecked: false,
      imageChecked: false,
    },
  )
  assert.deepEqual(codes, ['123456'])
})

test('real PC smoke rejects an unrelated project response', async (t) => {
  const origin = await fakeRelay(t, {
    projectId: 'other',
    status: 'compiled',
    slideCount: 0,
    slides: [],
    history: [],
  })
  await assert.rejects(
    inspectPcBusiness(origin, 'document-1', 'project-1', {
      onCode: () => {},
      timeoutMs: 1000,
    }),
    /PC presentation status response invalid/,
  )
})

test('real PC smoke rejects a response with mismatched session identity', async (t) => {
  const origin = await fakeRelay(
    t,
    {
      projectId: 'project-1',
      status: 'compiled',
      slideCount: 0,
      slides: [],
      history: [],
    },
    (frame) => (frame.type === 'relay.chunk' ? { ...frame, session_id: 'other' } : frame),
  )
  await assert.rejects(
    inspectPcBusiness(origin, 'document-1', 'project-1', {
      onCode: () => {},
      timeoutMs: 1000,
    }),
    /PC response identity mismatch/,
  )
})

test('PC status validator rejects shape errors', () => {
  assert.throws(() =>
    checkPcStatusResponse(
      { projectId: 'project-1', status: 'compiled', slideCount: 2, slides: [], history: [] },
      'project-1',
    ),
  )
})

test('real PC smoke reads text and validates normalized image bytes', async (t) => {
  const textId = 'a'.repeat(64)
  const imageId = 'b'.repeat(64)
  const image = Buffer.from('normalized-png-fixture')
  const origin = await fakeRelay(t, {
    projectId: 'project-1',
    status: 'compiled',
    slideCount: 0,
    slides: [],
    history: [],
    attachment_metadata: {
      [textId]: { attachmentId: textId, status: 'ready', kind: 'text', totalChars: 5 },
      [imageId]: {
        attachmentId: imageId,
        status: 'ready',
        kind: 'image',
        width: 1,
        height: 1,
        assetSha256: createHash('sha256').update(image).digest('hex'),
      },
    },
    attachment_read: {
      attachmentId: textId,
      offset: 0,
      sourceUri: `attachment:${textId}`,
      totalChars: 5,
      text: 'hello',
    },
    attachment_asset: {
      id: imageId,
      mime: 'image/png',
      width: 1,
      height: 1,
      base64: image.toString('base64'),
    },
  })
  assert.deepEqual(
    await inspectPcBusiness(origin, 'document-1', 'project-1', {
      onCode: () => {},
      timeoutMs: 1000,
      textAttachmentId: textId,
      imageAttachmentId: imageId,
    }),
    {
      projectId: 'project-1',
      status: 'compiled',
      slideCount: 0,
      attachmentPageCount: 0,
      textChecked: true,
      imageChecked: true,
    },
  )
})

test('real PC smoke rejects an image whose bytes differ from PC metadata', async (t) => {
  const imageId = 'b'.repeat(64)
  const origin = await fakeRelay(t, {
    projectId: 'project-1',
    status: 'compiled',
    slideCount: 0,
    slides: [],
    history: [],
    attachment_metadata: {
      [imageId]: {
        attachmentId: imageId,
        status: 'ready',
        kind: 'image',
        width: 1,
        height: 1,
        assetSha256: '0'.repeat(64),
      },
    },
    attachment_asset: {
      id: imageId,
      mime: 'image/png',
      width: 1,
      height: 1,
      base64: Buffer.from('different-image').toString('base64'),
    },
  })
  await assert.rejects(
    inspectPcBusiness(origin, 'document-1', 'project-1', {
      onCode: () => {},
      timeoutMs: 1000,
      imageAttachmentId: imageId,
    }),
    /image attachment digest mismatch/,
  )
})
