import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { createServer } from 'node:http'
import { test } from 'node:test'
import { PDFDocument } from 'pdf-lib'
import JSZip from 'jszip'
import { WebSocketServer } from 'ws'
import {
  checkCompiledDelivery,
  checkPcStatusResponse,
  inspectPcBusiness,
} from './ppt-agent-pc-business-smoke.mjs'

test('compiled delivery check binds all PPTX text pages to the rendered PDF count', async () => {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', '<p:sld><a:t>Smoke title</a:t></p:sld>')
  zip.file('ppt/slides/slide2.xml', '<p:sld><a:t>Second page</a:t></p:sld>')
  const pdf = await PDFDocument.create()
  pdf.addPage()
  pdf.addPage()
  const compiled = {
    projectId: 'project-1',
    requestId: 'run-1',
    status: 'compiled',
    report: { slideCount: 2 },
    pptxBase64: (await zip.generateAsync({ type: 'nodebuffer' })).toString('base64'),
  }
  const exported = {
    projectId: 'project-1',
    requestId: 'run-1',
    status: 'exported',
    source: 'compiled',
    slideCount: 2,
    pdfBase64: Buffer.from(await pdf.save()).toString('base64'),
  }
  const expected = ['Smoke title', 'Second page']
  const checked = await checkCompiledDelivery(compiled, exported, 'project-1', expected)
  assert.match(checked.pptxSha256, /^[a-f0-9]{64}$/)
  assert.ok(checked.pdfBytes > 100)
  await assert.rejects(
    checkCompiledDelivery(compiled, { ...exported, requestId: 'other' }, 'project-1', expected),
    /response invalid/,
  )
  await assert.rejects(
    checkCompiledDelivery(compiled, exported, 'project-1', ['Smoke title', 'Different title']),
    /content invalid/,
  )
})

async function fakeRelay(t, response, mutate = (frame) => frame, handle) {
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
          handle?.(frame.body) ??
          (frame.body.operation === 'status'
            ? response
            : frame.body.operation === 'attachment_list_assets'
              ? { attachments: [] }
              : frame.body.operation === 'attachment_metadata'
                ? response.attachment_metadata?.[frame.body.attachmentId]
                : response[frame.body.operation])
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
      uploadChecked: false,
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
      uploadChecked: false,
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

test('real PC smoke uploads, reads, and deletes unique text and PNG fixtures', async (t) => {
  const fixtures = new Map()
  const deleted = []
  const response = {
    projectId: 'project-1',
    status: 'compiled',
    slideCount: 0,
    slides: [],
    history: [],
  }
  const origin = await fakeRelay(
    t,
    response,
    (frame) => frame,
    (body) => {
      const id = body.attachmentId
      if (body.operation === 'status') return response
      if (body.operation === 'attachment_list_assets') return { attachments: [] }
      if (body.operation === 'attachment_metadata')
        return fixtures.get(id)?.metadata ?? { error: 'not_found' }
      if (body.operation === 'attachment_begin') {
        assert.equal(fixtures.has(id), false)
        assert.equal(body.sha256, id)
        fixtures.set(id, {
          name: body.name,
          bytes: Buffer.alloc(0),
          metadata: {
            attachmentId: id,
            sha256: id,
            status: 'uploading',
            receivedBytes: 0,
          },
        })
        return fixtures.get(id).metadata
      }
      if (body.operation === 'attachment_chunk') {
        const fixture = fixtures.get(id)
        assert.equal(body.offset, 0)
        fixture.bytes = Buffer.from(body.base64, 'base64')
        return { attachmentId: id, status: 'uploading', receivedBytes: fixture.bytes.length }
      }
      if (body.operation === 'attachment_finish') {
        const fixture = fixtures.get(id)
        fixture.metadata = {
          attachmentId: id,
          sha256: id,
          status: 'ready',
          kind: fixture.name.endsWith('.txt') ? 'text' : 'image',
          ...(fixture.name.endsWith('.txt')
            ? { totalChars: fixture.bytes.toString('utf8').length }
            : {
                width: 1,
                height: 1,
                assetSha256: createHash('sha256').update(fixture.bytes).digest('hex'),
              }),
        }
        return fixture.metadata
      }
      if (body.operation === 'attachment_read') {
        const text = fixtures.get(id).bytes.toString('utf8')
        return {
          attachmentId: id,
          offset: 0,
          sourceUri: `attachment:${id}`,
          totalChars: text.length,
          text,
        }
      }
      if (body.operation === 'attachment_asset')
        return {
          id,
          mime: 'image/png',
          width: 1,
          height: 1,
          base64: fixtures.get(id).bytes.toString('base64'),
        }
      if (body.operation === 'attachment_delete') {
        assert.equal(fixtures.delete(id), true)
        deleted.push(id)
        return { attachmentId: id, deleted: true }
      }
    },
  )
  assert.deepEqual(
    await inspectPcBusiness(origin, 'document-1', 'project-1', {
      onCode: () => {},
      timeoutMs: 1000,
      uploadFixtures: true,
    }),
    {
      projectId: 'project-1',
      status: 'compiled',
      slideCount: 0,
      attachmentPageCount: 0,
      textChecked: true,
      imageChecked: true,
      uploadChecked: true,
    },
  )
  assert.equal(fixtures.size, 0)
  assert.equal(deleted.length, 2)
})

test('real PC smoke cleans an attachment after a failed upload chunk', async (t) => {
  const response = {
    projectId: 'project-1',
    status: 'compiled',
    slideCount: 0,
    slides: [],
    history: [],
  }
  let created
  let deleted
  const origin = await fakeRelay(
    t,
    response,
    (frame) => frame,
    (body) => {
      if (body.operation === 'status') return response
      if (body.operation === 'attachment_list_assets') return { attachments: [] }
      if (body.operation === 'attachment_metadata') return { error: 'not_found' }
      if (body.operation === 'attachment_begin') {
        created = body.attachmentId
        return { attachmentId: created, status: 'uploading', receivedBytes: 0 }
      }
      if (body.operation === 'attachment_chunk') return { error: 'disk_full' }
      if (body.operation === 'attachment_delete') {
        deleted = body.attachmentId
        return { attachmentId: deleted, deleted: true }
      }
    },
  )
  await assert.rejects(
    inspectPcBusiness(origin, 'document-1', 'project-1', {
      onCode: () => {},
      timeoutMs: 1000,
      uploadFixtures: true,
    }),
    /upload chunk failed/,
  )
  assert.equal(deleted, created)
})
