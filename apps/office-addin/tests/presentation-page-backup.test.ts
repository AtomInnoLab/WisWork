import { createHash } from 'node:crypto'
import { expect, it, vi } from 'vitest'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'
import { presentationArtifactContent } from '../src/skills/powerpoint/presentation-page-delivery.js'
import { createPresentationPageBackupSkill } from '../src/skills/powerpoint/presentation-page-backup.js'
function setup() {
  const bytes = Buffer.alloc(140000, 12)
  const sha256 = createHash('sha256').update(bytes).digest('hex')
  const artifact = {
    documentId: 'doc',
    projectId: 'project',
    requestId: 'parent',
    pptxBase64: '',
    slideCount: 1,
    planRevision: 1,
    pages: [{ id: 'page', title: 'Page', sourceSlideId: '256#' }],
    pagePptxBase64: ['UEsDBAAAAAA='],
  }
  const receipt = {
    state: 'complete' as const,
    documentId: 'doc',
    slideIds: ['host'],
    checkpoint: {
      version: 2 as const,
      artifactDigest: createHash('sha256')
        .update(presentationArtifactContent(artifact))
        .digest('hex'),
      sourceSlideIds: ['256#'],
      pageIds: ['page'],
      baselineSlideIds: ['original'],
      completed: [{ sourceSlideId: '256#', slideId: 'host' }],
    },
  }
  const status = {
    projectId: 'project',
    requestId: 'child',
    planRevision: 1,
    status: 'compiled',
    compiledCount: 1,
    total: 1,
    pages: [{ id: 'page', title: 'Page', state: 'compiled', attempt: 1 }],
    revision: { parentRequestId: 'parent', pageId: 'page', parentInputDigest: 'a'.repeat(64) },
  }
  const metadata = {
    backupId: 'backup',
    projectId: 'project',
    documentId: 'doc',
    requestId: 'child',
    parentRequestId: 'parent',
    pageId: 'page',
    hostSlideId: 'host',
    slideIds: ['original', 'host'],
    sha256,
    sizeBytes: bytes.length,
    parentInputDigest: 'a'.repeat(64),
    inputDigest: 'b'.repeat(64),
    status: 'uploading',
    receivedBytes: 0,
  }
  const request = vi.fn(async (body: unknown) => {
    const b = body as Record<string, unknown>
    if (b.operation === 'production_status') return Response.json(status)
    if (b.operation === 'page_backup_chunk')
      metadata.receivedBytes = Number(b.offset) + Buffer.from(String(b.base64), 'base64').length
    if (b.operation === 'page_backup_finish') metadata.status = 'ready'
    if (b.operation === 'page_backup_read')
      return Response.json({
        backupId: 'backup',
        offset: b.offset,
        sizeBytes: bytes.length,
        sha256,
        base64: bytes
          .subarray(Number(b.offset), Number(b.offset) + Number(b.length))
          .toString('base64'),
      })
    return Response.json(metadata)
  })
  const adapter = {
    exportPresentationPagePackage: vi.fn(async () => ({
      slideId: 'host',
      slideIds: ['original', 'host'],
      base64: bytes.toString('base64'),
    })),
  }
  const options = {
    available: () => true,
    request,
    documentId: vi.fn(async () => 'doc'),
    vfs: new InMemoryVfs(),
    artifact: () => artifact,
    readReceipt: () => receipt,
    adapter,
  }
  const skill = createPresentationPageBackupSkill(options)
  const save = {
    id: 'save',
    name: 'save_presentation_page_backup',
    input: { project_id: 'project', request_id: 'child', page_id: 'page', backup_id: 'backup' },
  }
  const read = {
    id: 'read',
    name: 'read_presentation_page_backup',
    input: { project_id: 'project', backup_id: 'backup' },
  }
  return {
    bytes,
    sha256,
    artifact,
    receipt,
    status,
    metadata,
    request,
    adapter,
    options,
    skill,
    save,
    read,
  }
}
it('exports the stable imported parent page, uploads bounded chunks, and downloads exact bytes without model base64', async () => {
  const f = setup()
  const result = await f.skill.executeTool(f.save)
  expect(result.isError).not.toBe(true)
  expect(JSON.parse(result.output)).toMatchObject({
    status: 'ready',
    historical: true,
    backupId: 'backup',
  })
  expect(f.adapter.exportPresentationPagePackage).toHaveBeenCalledWith('host', undefined)
  expect(
    f.request.mock.calls.filter(
      ([b]) => (b as { operation: string }).operation === 'page_backup_chunk',
    ),
  ).toHaveLength(2)
  const downloaded = await f.skill.executeTool(f.read)
  expect(downloaded.isError).not.toBe(true)
  const value = JSON.parse(downloaded.output)
  expect(value.base64).toBeUndefined()
  expect(f.options.vfs.readBytes(value.path)).toEqual(new Uint8Array(f.bytes))
})
it('resumes an interrupted upload at the persisted offset and never exports a different page', async () => {
  const f = setup()
  f.metadata.receivedBytes = 131072
  expect((await f.skill.executeTool(f.save)).isError).not.toBe(true)
  const chunks = f.request.mock.calls.filter(
    ([b]) => (b as { operation: string }).operation === 'page_backup_chunk',
  )
  expect(chunks).toHaveLength(1)
  expect(chunks[0][0]).toMatchObject({ offset: 131072 })
})
it('does not repeat chunks when the immutable snapshot is already ready', async () => {
  const f = setup()
  f.metadata.status = 'ready'
  f.metadata.receivedBytes = f.bytes.length
  expect((await f.skill.executeTool(f.save)).isError).not.toBe(true)
  expect(
    f.request.mock.calls.some(
      ([b]) => (b as { operation: string }).operation === 'page_backup_chunk',
    ),
  ).toBe(false)
})
it.each([
  'wrong-parent',
  'uncompiled',
  'wrong-target',
  'wrong-digest',
  'partial-import',
  'wrong-document',
])('blocks invalid binding before export: %s', async (reason) => {
  const f = setup()
  if (reason === 'wrong-parent') f.status.revision.parentRequestId = 'other'
  if (reason === 'uncompiled') {
    f.status.status = 'pending'
    f.status.compiledCount = 0
    f.status.pages[0].state = 'pending'
    f.status.pages[0].attempt = 0
  }
  if (reason === 'wrong-target') f.status.revision.pageId = 'other'
  if (reason === 'wrong-digest') f.receipt.checkpoint.artifactDigest = 'c'.repeat(64)
  if (reason === 'partial-import')
    Object.assign(f.receipt, { state: 'pending', slideIds: undefined })
  if (reason === 'wrong-document') f.artifact.documentId = 'other'
  expect((await f.skill.executeTool(f.save)).isError).toBe(true)
  expect(f.adapter.exportPresentationPagePackage).not.toHaveBeenCalled()
})
it.each(['export', 'begin', 'chunk', 'download'])(
  'stops cleared sessions at asynchronous boundaries: %s',
  async (boundary) => {
    const f = setup()
    if (boundary === 'export')
      f.adapter.exportPresentationPagePackage.mockImplementationOnce(async () => {
        f.skill.clear()
        return { slideId: 'host', slideIds: ['host'], base64: f.bytes.toString('base64') }
      })
    else {
      const original = f.request.getMockImplementation()!
      f.request.mockImplementation(async (body) => {
        const result = await original(body)
        const op = (body as { operation: string }).operation
        if (
          op ===
          (
            {
              begin: 'page_backup_begin',
              chunk: 'page_backup_chunk',
              download: 'page_backup_read',
            } as Record<string, string>
          )[boundary]
        )
          f.skill.clear()
        return result
      })
    }
    if (boundary === 'download') {
      f.metadata.status = 'ready'
      f.metadata.receivedBytes = f.bytes.length
    }
    expect(await f.skill.executeTool(boundary === 'download' ? f.read : f.save)).toMatchObject({
      output: 'cancelled',
      isError: true,
    })
    expect(f.options.vfs.list('/home/user')).toEqual([])
    expect(
      f.request.mock.calls.some(
        ([b]) => (b as { operation: string }).operation === 'page_backup_finish',
      ),
    ).toBe(false)
  },
)
it.each(['document', 'artifact', 'receipt'])(
  'stops changed context after exporting: %s',
  async (change) => {
    const f = setup()
    const original = f.adapter.exportPresentationPagePackage.getMockImplementation()!
    f.adapter.exportPresentationPagePackage.mockImplementationOnce(async () => {
      const result = await original()
      if (change === 'document') f.options.documentId.mockResolvedValue('changed')
      if (change === 'artifact') f.artifact.pages[0].title = 'changed'
      if (change === 'receipt') f.receipt.slideIds[0] = 'changed'
      return result
    })
    expect((await f.skill.executeTool(f.save)).isError).toBe(true)
    expect(
      f.request.mock.calls.some(
        ([b]) => (b as { operation: string }).operation === 'page_backup_begin',
      ),
    ).toBe(false)
  },
)
it.each(['id', 'oversize', 'order', 'extra'])('rejects malformed exports: %s', async (kind) => {
  const f = setup()
  const result = { slideId: 'host', slideIds: ['host'], base64: f.bytes.toString('base64') }
  if (kind === 'id') result.slideId = 'wrong'
  if (kind === 'oversize') result.base64 = Buffer.alloc(8 * 1024 * 1024 + 1).toString('base64')
  if (kind === 'order') result.slideIds = ['host', 'host']
  if (kind === 'extra') Object.assign(result, { unexpected: true })
  f.adapter.exportPresentationPagePackage.mockResolvedValue(result)
  expect((await f.skill.executeTool(f.save)).isError).toBe(true)
  expect(
    f.request.mock.calls.some(
      ([b]) => (b as { operation: string }).operation === 'page_backup_begin',
    ),
  ).toBe(false)
})
it.each(['identity', 'digest', 'size', 'progress', 'extra'])(
  'rejects malformed or mismatched upload metadata: %s',
  async (kind) => {
    const f = setup()
    if (kind === 'identity') f.metadata.requestId = 'foreign'
    if (kind === 'digest') f.metadata.sha256 = 'c'.repeat(64)
    if (kind === 'size') f.metadata.sizeBytes++
    if (kind === 'progress') f.metadata.receivedBytes = -1
    if (kind === 'extra') Object.assign(f.metadata, { base64: 'secret' })
    expect((await f.skill.executeTool(f.save)).isError).toBe(true)
    expect(
      f.request.mock.calls.some(
        ([b]) => (b as { operation: string }).operation === 'page_backup_chunk',
      ),
    ).toBe(false)
  },
)
it.each(['corrupt', 'identity', 'offset', 'size', 'extra'])(
  'never publishes incorrect downloads: %s',
  async (kind) => {
    const f = setup()
    f.metadata.status = 'ready'
    f.metadata.receivedBytes = f.bytes.length
    const original = f.request.getMockImplementation()!
    f.request.mockImplementation(async (body) => {
      const response = await original(body)
      if ((body as { operation: string }).operation !== 'page_backup_read') return response
      const result = await response.json()
      if (kind === 'corrupt')
        result.base64 = Buffer.alloc(Buffer.from(result.base64, 'base64').length, 13).toString(
          'base64',
        )
      if (kind === 'identity') result.backupId = 'other'
      if (kind === 'offset') result.offset++
      if (kind === 'size') result.sizeBytes++
      if (kind === 'extra') result.unexpected = true
      return Response.json(result)
    })
    expect((await f.skill.executeTool(f.read)).isError).toBe(true)
    expect(f.options.vfs.list('/home/user')).toEqual([])
  },
)
it('downloads historical backups without an active artifact and refuses to overwrite a session file', async () => {
  const f = setup()
  f.metadata.status = 'ready'
  f.metadata.receivedBytes = f.bytes.length
  f.options.artifact = () => {
    throw new Error('must not read active artifact')
  }
  const first = await f.skill.executeTool(f.read)
  expect(first.isError).not.toBe(true)
  const path = JSON.parse(first.output).path
  f.options.vfs.writeFile(path, 'user changes')
  expect(await f.skill.executeTool(f.read)).toMatchObject({
    isError: true,
    output: 'presentation_backup_file_exists',
  })
  expect(f.options.vfs.readText(path)).toBe('user changes')
})
it('cancels aborted calls and rejects invalid tool arguments before PC or Office calls', async () => {
  const f = setup()
  const controller = new AbortController()
  controller.abort()
  expect(await f.skill.executeTool(f.save, controller.signal)).toMatchObject({
    isError: true,
    output: 'cancelled',
  })
  expect(
    await f.skill.executeTool({ ...f.save, input: { ...f.save.input, extra: true } }),
  ).toMatchObject({ isError: true, output: 'invalid_tool_input' })
  expect(f.request).not.toHaveBeenCalled()
  expect(f.adapter.exportPresentationPagePackage).not.toHaveBeenCalled()
})
it.each([
  [{ error: 'unsupported_operation' }, 'presentation_upgrade_required'],
  [
    { error: 'invalid_request', message: 'Unsupported operation: page_backup_begin' },
    'presentation_upgrade_required',
  ],
  [{ error: 'invalid_request' }, 'presentation_invalid_request'],
] as const)(
  'only labels explicit unsupported-operation errors as an upgrade: %j',
  async (error, expected) => {
    const f = setup()
    f.request.mockResolvedValue(Response.json(error, { status: 400 }))
    expect(await f.skill.executeTool(f.read)).toMatchObject({ isError: true, output: expected })
  },
)
it.each(['inputError', 'truncated'])('rejects incomplete parser inputs: %s', async (field) => {
  const f = setup()
  expect(
    await f.skill.executeTool({
      ...f.save,
      [field]: field === 'truncated' ? true : 'invalid json',
    }),
  ).toMatchObject({ isError: true, output: 'invalid_tool_input' })
  expect(f.request).not.toHaveBeenCalled()
})
it('supports a 128-character backup ID without exceeding VFS filename limits', async () => {
  const f = setup()
  const backupId = 'b'.repeat(128)
  f.metadata.backupId = backupId
  f.save.input.backup_id = backupId
  expect((await f.skill.executeTool(f.save)).isError).not.toBe(true)
  f.read.input.backup_id = backupId
  const original = f.request.getMockImplementation()!
  f.request.mockImplementation(async (body) => {
    const response = await original(body)
    const value = await response.json()
    if ('backupId' in value) value.backupId = backupId
    return Response.json(value)
  })
  const result = await f.skill.executeTool(f.read)
  expect(result.isError).not.toBe(true)
  expect(f.options.vfs.readBytes(JSON.parse(result.output).path)).toEqual(new Uint8Array(f.bytes))
})
