import assert from 'node:assert/strict'
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
        assert.deepEqual(frame.capabilities, ['presentation.v1'])
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
            capabilities: ['presentation.v1'],
            expires_in: 1800,
          }),
        )
      } else if (frame.type === 'office.request') {
        assert.equal(frame.body.operation, 'status')
        assert.equal(frame.body.documentId, 'document-1')
        assert.equal(frame.body.projectId, 'project-1')
        const common = { version: 2, session_id: frame.session_id, request_id: frame.request_id }
        const frames = [
          { ...common, type: 'relay.start', status: 200, content_type: 'application/json' },
          {
            ...common,
            type: 'relay.chunk',
            sequence: 0,
            data: Buffer.from(JSON.stringify(response)).toString('base64'),
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
    { projectId: 'project-1', status: 'compiled', slideCount: 1 },
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
