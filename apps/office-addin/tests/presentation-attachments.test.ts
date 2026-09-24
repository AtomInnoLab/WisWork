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
    if (body.operation === 'attachment_delete')
      return new Response(JSON.stringify({ attachmentId: body.attachmentId, deleted: true }))
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
        body.operation === 'attachment_list' || body.operation === 'attachment_list_assets'
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
it('lists document-scoped PC copies and deletes only a validated selected ID', async () => {
  const f = setup()
  expect(await f.skill.list()).toMatchObject([{ attachmentId: f.attachmentId }])
  await expect(f.skill.remove('bad')).rejects.toThrow('invalid_tool_input')
  await f.skill.remove(f.attachmentId)
  expect(f.request).toHaveBeenCalledWith(
    expect.objectContaining({
      operation: 'attachment_delete',
      documentId: 'doc1',
      attachmentId: f.attachmentId,
    }),
    expect.any(AbortSignal),
  )
  expect(f.skill.tools.map((tool) => tool.name)).not.toContain('delete_presentation_attachment')
})
it('pages more than 32 durable images for the UI and Agent', async () => {
  const f = setup()
  const ids = Array.from({ length: 33 }, (_, i) => i.toString(16).padStart(64, '0'))
  f.request.mockImplementation(async (body) => {
    const start = body.after ? ids.indexOf(body.after) + 1 : 0
    const page = ids.slice(start, start + 32)
    return new Response(
      JSON.stringify({
        attachments: page.map((attachmentId) => ({
          attachmentId,
          sha256: attachmentId,
          name: `${attachmentId}.png`,
          sizeBytes: 1,
          receivedBytes: 0,
          status: 'uploading',
        })),
        ...(start + page.length < ids.length ? { nextAfter: page.at(-1) } : {}),
      }),
    )
  })
  const skill = createPresentationAttachmentSkill({
    available: () => true,
    imagesAvailable: () => true,
    request: f.request,
    documentId: f.documentId,
    vfs: f.vfs,
  })
  expect((await skill.list()).map((item) => item.attachmentId)).toEqual(ids)
  const first = JSON.parse(
    (await skill.executeTool({ id: 'list', name: 'list_presentation_attachments', input: {} }))
      .output,
  )
  expect(first.attachments).toHaveLength(32)
  expect(first.nextAfter).toBe(ids[31])
  const second = JSON.parse(
    (
      await skill.executeTool({
        id: 'next',
        name: 'list_presentation_attachments',
        input: { after: first.nextAfter },
      })
    ).output,
  )
  expect(second.attachments.map((item: { attachmentId: string }) => item.attachmentId)).toEqual([
    ids[32],
  ])
})
it('imports a URL image through the PC asset endpoint without exposing an Agent fetch tool', async () => {
  const f = setup()
  f.request.mockResolvedValue(
    new Response(
      JSON.stringify({
        attachmentId: f.attachmentId,
        sha256: f.attachmentId,
        name: 'remote.png',
        sizeBytes: 1,
        receivedBytes: 1,
        status: 'ready',
        kind: 'image',
        mime: 'image/png',
        width: 1,
        height: 1,
        assetSha256: f.attachmentId,
        source: 'https://example.com/image.png',
      }),
    ),
  )
  const skill = createPresentationAttachmentSkill({
    available: () => true,
    imagesAvailable: () => true,
    remoteImagesAvailable: () => true,
    request: f.request,
    documentId: f.documentId,
    vfs: f.vfs,
  })
  expect(await skill.importUrl('https://example.com/image.png')).toMatchObject({
    kind: 'image',
    source: 'https://example.com/image.png',
  })
  expect(f.request).toHaveBeenCalledWith(
    expect.objectContaining({ operation: 'attachment_import_url', documentId: 'doc1' }),
    expect.any(AbortSignal),
  )
  expect(skill.tools.map((tool) => tool.name)).not.toContain('import_presentation_image_url')
})
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
function imageFixture(overrides: Record<string, unknown> = {}) {
  const bytes = new Uint8Array([1, 2, 3]),
    attachmentId = createHash('sha256').update(bytes).digest('hex')
  const value = {
    attachmentId,
    name: 'photo.jpg',
    sha256: attachmentId,
    sizeBytes: 3,
    receivedBytes: 3,
    status: 'ready',
    kind: 'image',
    mime: 'image/png',
    width: 800,
    height: 600,
    assetSha256: 'b'.repeat(64),
    ...overrides,
  }
  const imagesAvailable = vi.fn(() => true),
    documentId = vi.fn(async () => 'doc')
  const request = vi.fn(
    async (body: unknown) =>
      new Response(
        JSON.stringify(
          (body as { operation: string }).operation === 'attachment_list_assets'
            ? { attachments: [value] }
            : value,
        ),
      ),
  )
  const vfs = new InMemoryVfs(),
    skill = createPresentationAttachmentSkill({
      available: () => true,
      imagesAvailable,
      request,
      documentId,
      vfs,
    })
  return { bytes, value, imagesAvailable, request, vfs, skill, documentId }
}
it('uploads image originals and lists validated compact asset metadata on the negotiated channel', async () => {
  const f = imageFixture()
  await f.skill.upload('photo.jpg', Promise.resolve(f.bytes.buffer))
  const result = await f.skill.executeTool({
    id: 'list',
    name: 'list_presentation_attachments',
    input: {},
  })
  expect(JSON.parse(result.output)).toEqual({ attachments: [f.value] })
  expect(result.output).not.toContain('base64')
  expect(f.request).toHaveBeenLastCalledWith(
    { operation: 'attachment_list_assets', documentId: 'doc' },
    expect.any(AbortSignal),
  )
  expect(f.skill.systemPrompt).toContain('attachmentId: listed_attachmentId')
})
it.each([
  { width: 0 },
  { height: 16385 },
  { width: 9000, height: 1 },
  { width: 5000, height: 4000 },
  { width: 10000, height: 10000 },
  { assetSha256: 'wrong' },
  { mime: 'image/jpeg' },
  { totalChars: 5 },
  { kind: 'text' },
  { source: 'file:///secret' },
])('rejects forged image metadata %j', async (overrides) => {
  const f = imageFixture(overrides)
  expect(
    await f.skill.executeTool({ id: 'list', name: 'list_presentation_attachments', input: {} }),
  ).toMatchObject({ isError: true, output: 'presentation_response_invalid' })
  await expect(f.skill.upload('photo.jpg', Promise.resolve(f.bytes.buffer))).rejects.toThrow(
    'presentation_response_invalid',
  )
  expect(f.vfs.list('/home/user')).toEqual([])
})
it('does not publish image upload results after image capability loss', async () => {
  const f = imageFixture(),
    original = f.request.getMockImplementation()!
  f.request.mockImplementation(async (body) => {
    const response = await original(body)
    f.imagesAvailable.mockReturnValue(false)
    return response
  })
  await expect(f.skill.upload('photo.jpg', Promise.resolve(f.bytes.buffer))).rejects.toThrow(
    'presentation_assets_unavailable',
  )
  expect(f.vfs.list('/home/user')).toEqual([])
})
it('rejects oversized images and unnegotiated image uploads without dispatch', async () => {
  const f = imageFixture()
  await expect(
    f.skill.upload('large.png', Promise.resolve(new ArrayBuffer(10 * 1024 * 1024 + 1))),
  ).rejects.toThrow('presentation_image_too_large')
  f.imagesAvailable.mockReturnValue(false)
  await expect(f.skill.upload('photo.jpg', Promise.resolve(f.bytes.buffer))).rejects.toThrow(
    'presentation_assets_unavailable',
  )
  expect(f.request).not.toHaveBeenCalled()
  expect(f.skill.systemPrompt).not.toContain('attachmentId: listed_attachmentId')
})
