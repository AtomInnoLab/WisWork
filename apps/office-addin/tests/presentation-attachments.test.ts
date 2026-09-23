import { createHash } from 'node:crypto'
import { expect, it, vi } from 'vitest'
import { createPresentationAttachmentSkill } from '../src/skills/powerpoint/presentation-attachments.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'
function setup() {
  const bytes = new Uint8Array(300000)
  const attachmentId = createHash('sha256').update(bytes).digest('hex')
  let receivedBytes = 0
  let status = 'uploading'
  const request = vi.fn(async (body: any, _signal?: AbortSignal) => {
    if (body.operation === 'attachment_chunk')
      receivedBytes = body.offset + atob(body.base64).length
    if (body.operation === 'attachment_finish') status = 'ready'
    const metadata = {
      attachmentId,
      name: 'notes.txt',
      sizeBytes: bytes.length,
      sha256: attachmentId,
      receivedBytes,
      status,
      ...(status === 'ready' ? { kind: 'text', totalChars: 5 } : {}),
    }
    return new Response(
      JSON.stringify(
        body.operation === 'attachment_list'
          ? { attachments: [metadata] }
          : body.operation === 'attachment_read'
            ? {
                attachmentId,
                name: 'notes.txt',
                offset: 0,
                totalChars: 5,
                text: 'hello',
                sourceUri: `attachment:${attachmentId}`,
              }
            : metadata,
      ),
    )
  })
  const documentId = vi.fn(async () => 'doc1')
  const vfs = new InMemoryVfs()
  const skill = createPresentationAttachmentSkill({
    available: () => true,
    request,
    documentId,
    vfs,
  })
  return { bytes, attachmentId, request, documentId, vfs, skill }
}
it('uploads chunks and reads durable sources after reconnect', async () => {
  const f = setup()
  await f.skill.upload('notes.txt', Promise.resolve(f.bytes.buffer))
  expect(
    f.request.mock.calls
      .filter(([b]) => b.operation === 'attachment_chunk')
      .map(([b]) => atob(b.base64).length),
  ).toEqual([131072, 131072, 37856])
  expect(f.vfs.list('/home/user')).toContain('/home/user/notes.txt')
  f.skill.clear()
  expect(
    JSON.parse(
      (
        await f.skill.executeTool({
          id: 'read',
          name: 'read_presentation_attachment',
          input: { attachment_id: f.attachmentId },
        })
      ).output,
    ),
  ).toMatchObject({ text: 'hello', sourceUri: `attachment:${f.attachmentId}` })
})
it('resumes after a lost chunk acknowledgement', async () => {
  const f = setup()
  const original = f.request.getMockImplementation()!
  let lost = false
  f.request.mockImplementation(async (body, signal) => {
    const response = await original(body, signal)
    if (body.operation === 'attachment_chunk' && !lost) {
      lost = true
      throw new Error('lost')
    }
    return response
  })
  await expect(f.skill.upload('notes.txt', Promise.resolve(f.bytes.buffer))).rejects.toThrow()
  await f.skill.upload('notes.txt', Promise.resolve(f.bytes.buffer))
  expect(
    f.request.mock.calls.filter(([b]) => b.operation === 'attachment_chunk').map(([b]) => b.offset),
  ).toEqual([0, 131072, 262144])
})
it('aborts and prevents late publication after clear', async () => {
  const f = setup()
  const original = f.request.getMockImplementation()!
  f.request.mockImplementation(async (body, signal) => {
    f.skill.clear()
    expect(signal?.aborted).toBe(true)
    return original(body, signal)
  })
  await expect(f.skill.upload('notes.txt', Promise.resolve(f.bytes.buffer))).rejects.toThrow(
    'upload_cancelled',
  )
  expect(f.vfs.list('/home/user')).toEqual([])
})
it('rejects changed documents and mismatched responses', async () => {
  const f = setup()
  f.documentId.mockResolvedValueOnce('doc1').mockResolvedValue('doc2')
  await expect(f.skill.upload('notes.txt', Promise.resolve(f.bytes.buffer))).rejects.toThrow(
    'presentation_document_changed',
  )
  expect(f.request).not.toHaveBeenCalled()
  const g = setup()
  g.request.mockResolvedValue(new Response('{}'))
  await expect(g.skill.upload('notes.txt', Promise.resolve(g.bytes.buffer))).rejects.toThrow(
    'presentation_response_invalid',
  )
  expect(g.vfs.list('/home/user')).toEqual([])
})
it('retries extraction after a complete upload previously failed parsing', async () => {
  const f = setup()
  const original = f.request.getMockImplementation()!
  await f.skill.upload('notes.txt', Promise.resolve(f.bytes.buffer))
  f.request.mockImplementation(async (body, signal) => {
    const response = await original(body, signal)
    if (body.operation === 'attachment_begin') {
      return new Response(
        JSON.stringify({ ...(await response.json()), status: 'failed', error: 'parse_failed' }),
      )
    }
    return response
  })
  f.request.mockClear()
  await f.skill.upload('notes.txt', Promise.resolve(f.bytes.buffer))
  expect(f.request.mock.calls.map(([b]) => b.operation)).toEqual([
    'attachment_begin',
    'attachment_finish',
  ])
})
it('rejects foreign text, duplicate inventory and oversized read inputs', async () => {
  const f = setup()
  await f.skill.upload('notes.txt', Promise.resolve(f.bytes.buffer))
  const original = f.request.getMockImplementation()!
  f.request.mockImplementation(async (body, signal) => {
    const data = await (await original(body, signal)).json()
    return new Response(
      JSON.stringify(
        body.operation === 'attachment_read'
          ? { ...data, sourceUri: 'attachment:wrong' }
          : { attachments: [...data.attachments, ...data.attachments] },
      ),
    )
  })
  expect(
    await f.skill.executeTool({
      id: 'read',
      name: 'read_presentation_attachment',
      input: { attachment_id: f.attachmentId },
    }),
  ).toMatchObject({ isError: true, output: 'presentation_response_invalid' })
  expect(
    await f.skill.executeTool({ id: 'list', name: 'list_presentation_attachments', input: {} }),
  ).toMatchObject({ isError: true, output: 'presentation_response_invalid' })
  expect(
    await f.skill.executeTool({
      id: 'read',
      name: 'read_presentation_attachment',
      input: { attachment_id: f.attachmentId, max_chars: 24001 },
    }),
  ).toMatchObject({ isError: true, output: 'invalid_tool_input' })
})
it('does not publish read text after abort or a document switch during the request', async () => {
  const f = setup(),
    original = f.request.getMockImplementation()!
  f.request.mockImplementation(async (body, signal) => {
    f.documentId.mockResolvedValue('doc2')
    return original(body, signal)
  })
  expect(
    await f.skill.executeTool({
      id: 'read',
      name: 'read_presentation_attachment',
      input: { attachment_id: f.attachmentId },
    }),
  ).toMatchObject({ isError: true, output: 'presentation_document_changed' })
  const g = setup(),
    controller = new AbortController()
  controller.abort()
  expect(
    await g.skill.executeTool(
      { id: 'list', name: 'list_presentation_attachments', input: {} },
      controller.signal,
    ),
  ).toMatchObject({ isError: true, output: 'upload_cancelled' })
  expect(g.request).not.toHaveBeenCalled()
})
it('accepts a durable original above the session file limit without a local copy', async () => {
  const bytes = new Uint8Array(21 * 1024 * 1024),
    attachmentId = createHash('sha256').update(bytes).digest('hex')
  const vfs = new InMemoryVfs()
  const request = vi.fn(
    async () =>
      new Response(
        JSON.stringify({
          attachmentId,
          name: 'large.pdf',
          sha256: attachmentId,
          sizeBytes: bytes.length,
          receivedBytes: bytes.length,
          status: 'ready',
          kind: 'text',
          totalChars: 7,
        }),
      ),
  )
  const skill = createPresentationAttachmentSkill({
    available: () => true,
    documentId: async () => 'doc',
    vfs,
    request,
  })
  await skill.upload('large.pdf', Promise.resolve(bytes.buffer))
  expect(request).toHaveBeenCalledOnce()
  expect(vfs.list('/home/user')).toEqual([])
})
