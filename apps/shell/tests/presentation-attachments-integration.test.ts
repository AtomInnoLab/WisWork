import { createHash } from 'node:crypto'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import {
  buildDocxFixture,
  buildPdfFixture,
} from '../../../packages/file-parse/tests/helpers/fixtures'
import { createPresentationAttachmentSkill } from '../../office-addin/src/skills/powerpoint/presentation-attachments'
import { InMemoryVfs } from '../../office-addin/src/skills/shared/vfs'
import { createPresentationService } from '../src/main/presentation-service'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'

const roots: string[] = []
afterEach(() => {
  for (const path of roots.splice(0)) rmSync(path, { recursive: true, force: true })
})
const signal = () => new AbortController().signal
const decode = (bytes: Uint8Array) => JSON.parse(Buffer.from(bytes).toString('utf8'))

describe('presentation attachment service integration', () => {
  it('reports a saved HTML paragraph mismatch and accepts the corrected locator', async () => {
    const userDataPath = mkdtempSync(join(tmpdir(), 'ppt-html-locator-'))
    roots.push(userDataPath)
    const service = createPresentationService({ userDataPath })
    const plan = benchmarkPlan()
    const documentId = 'source-document'
    const raw = Buffer.from('<p>Earlier finding</p><p>Target finding</p>')
    const attachmentId = createHash('sha256').update(raw).digest('hex')
    plan.sources[0]!.uri = `attachment:${attachmentId}`
    plan.sources[0]!.excerpt = 'Target finding'
    plan.sources[0]!.locator = '第 1 段'
    const send = async (body: Record<string, unknown>) =>
      decode(
        await service(
          {
            documentId,
            ...(!String(body.operation).startsWith('attachment_')
              ? { projectId: plan.projectId }
              : {}),
            ...body,
          },
          signal(),
        ),
      )
    await send({ operation: 'save_plan', expectedRevision: 0, plan })
    await send({
      operation: 'attachment_begin',
      attachmentId,
      sha256: attachmentId,
      name: 'study.html',
      sizeBytes: raw.length,
    })
    await send({
      operation: 'attachment_chunk',
      attachmentId,
      offset: 0,
      base64: raw.toString('base64'),
    })
    await send({ operation: 'attachment_finish', attachmentId })
    expect((await send({ operation: 'audit_sources' })).sources[0]).toMatchObject({
      status: 'found',
      locator: '第 2 段',
    })
    expect((await send({ operation: 'status' })).sourcePreparation[0].status).toBe(
      'locator_mismatch',
    )
    plan.sources[0]!.locator = '第 2 段'
    await send({ operation: 'save_plan', expectedRevision: 1, plan })
    expect((await send({ operation: 'status' })).sourcePreparation[0].status).toBe(
      'excerpt_matched',
    )
  })
  it('shows a PDF locator mismatch in project readiness before compilation', async () => {
    const userDataPath = mkdtempSync(join(tmpdir(), 'ppt-pdf-locator-'))
    roots.push(userDataPath)
    const service = createPresentationService({ userDataPath })
    const plan = benchmarkPlan()
    const documentId = 'source-document'
    const raw = Buffer.from(buildPdfFixture(['Earlier finding', 'Target finding']))
    const attachmentId = createHash('sha256').update(raw).digest('hex')
    plan.sources[0]!.uri = `attachment:${attachmentId}`
    plan.sources[0]!.excerpt = 'Target finding'
    plan.sources[0]!.locator = '第 1 页'
    const send = async (body: Record<string, unknown>) =>
      decode(
        await service(
          {
            documentId,
            ...(!String(body.operation).startsWith('attachment_')
              ? { projectId: plan.projectId }
              : {}),
            ...body,
          },
          signal(),
        ),
      )
    await send({ operation: 'save_plan', expectedRevision: 0, plan })
    await send({
      operation: 'attachment_begin',
      attachmentId,
      sha256: attachmentId,
      name: 'study.pdf',
      sizeBytes: raw.length,
    })
    await send({
      operation: 'attachment_chunk',
      attachmentId,
      offset: 0,
      base64: raw.toString('base64'),
    })
    await send({ operation: 'attachment_finish', attachmentId })
    expect((await send({ operation: 'audit_sources' })).sources[0]).toMatchObject({
      status: 'found',
      locator: '第 2 页',
    })
    expect((await send({ operation: 'status' })).sourcePreparation[0].status).toBe(
      'locator_mismatch',
    )
    plan.sources[0]!.locator = '第 2 页'
    await send({ operation: 'save_plan', expectedRevision: 1, plan })
    expect((await send({ operation: 'status' })).sourcePreparation[0].status).toBe(
      'excerpt_matched',
    )
  })
  it('rejects a fetched webpage snapshot attributed to a different plan URL', async () => {
    const userDataPath = mkdtempSync(join(tmpdir(), 'ppt-url-source-'))
    roots.push(userDataPath)
    const service = createPresentationService({
      userDataPath,
      fetchPage: async () =>
        new Response('<html><body><p>示例数据仅用于测试</p></body></html>', {
          headers: { 'content-type': 'text/html' },
        }),
    })
    const plan = benchmarkPlan()
    const documentId = 'source-document'
    const send = async (body: Record<string, unknown>) =>
      decode(
        await service(
          {
            documentId,
            ...(!String(body.operation).startsWith('attachment_')
              ? { projectId: plan.projectId }
              : {}),
            ...body,
          },
          signal(),
        ),
      )
    const correctUrl = 'https://8.8.8.8/research?version=1'
    const snapshot = await send({ operation: 'attachment_import_webpage', url: correctUrl })
    plan.sources[0]!.snapshotAttachmentId = snapshot.attachmentId
    plan.sources[0]!.locator = '第 1 段'
    plan.sources[0]!.uri = 'https://8.8.8.8/research?version=2'
    await send({ operation: 'save_plan', expectedRevision: 0, plan })
    expect((await send({ operation: 'audit_sources' })).sources[0].status).toBe('source_mismatch')
    expect((await send({ operation: 'status' })).sourcePreparation[0].status).toBe(
      'source_mismatch',
    )
    plan.sources[0]!.uri = correctUrl
    await send({ operation: 'save_plan', expectedRevision: 1, plan })
    expect((await send({ operation: 'audit_sources' })).sources[0].status).toBe('found')
    expect((await send({ operation: 'status' })).sourcePreparation[0].status).toBe(
      'excerpt_matched',
    )
  })
  it('rebuilds plan source readiness from document-bound attachment metadata', async () => {
    const userDataPath = mkdtempSync(join(tmpdir(), 'ppt-source-readiness-'))
    roots.push(userDataPath)
    let service = createPresentationService({ userDataPath })
    const documentId = 'source-document'
    const plan = benchmarkPlan()
    const bytes = Buffer.from(buildPdfFixture('Traceable research source'))
    const attachmentId = createHash('sha256').update(bytes).digest('hex')
    plan.sources[0]!.uri = `attachment:${attachmentId}`
    plan.sources[0]!.excerpt = 'Traceable research source'
    const send = async (body: Record<string, unknown>) =>
      decode(
        await service(
          {
            documentId,
            ...(!String(body.operation).startsWith('attachment_')
              ? { projectId: plan.projectId }
              : {}),
            ...body,
          },
          signal(),
        ),
      )
    await send({ operation: 'save_plan', expectedRevision: 0, plan })
    expect((await send({ operation: 'audit_sources' })).sources[0].status).toBe('missing')
    expect((await send({ operation: 'status' })).sourcePreparation).toEqual([
      { sourceId: plan.sources[0]!.id, attachmentId, status: 'missing' },
    ])
    await send({
      operation: 'attachment_begin',
      attachmentId,
      name: 'research.pdf',
      sizeBytes: bytes.length,
      sha256: attachmentId,
    })
    expect((await send({ operation: 'status' })).sourcePreparation[0].status).toBe('uploading')
    expect((await send({ operation: 'audit_sources' })).sources[0].status).toBe('not_ready')
    await send({
      operation: 'attachment_chunk',
      attachmentId,
      offset: 0,
      base64: bytes.toString('base64'),
    })
    await send({ operation: 'attachment_finish', attachmentId })
    service = createPresentationService({ userDataPath })
    expect((await send({ operation: 'status' })).sourcePreparation[0].status).toBe(
      'excerpt_matched',
    )
    expect(await send({ operation: 'audit_sources' })).toMatchObject({
      projectId: plan.projectId,
      planRevision: 1,
      sources: [
        { sourceId: plan.sources[0]!.id, attachmentId, status: 'found', locator: '第 1 页' },
      ],
      checks: {
        support: 'not_verified',
        sourceAuthority: 'not_verified',
        timeliness: 'not_verified',
      },
    })
    const revised = structuredClone(plan)
    revised.sources[0]!.uri = 'https://example.com/research'
    revised.sources[0]!.snapshotAttachmentId = attachmentId
    await send({ operation: 'save_plan', expectedRevision: 1, plan: revised })
    expect((await send({ operation: 'status' })).sourcePreparation[0].status).toBe(
      'excerpt_matched',
    )
    expect((await send({ operation: 'audit_sources' })).sources[0]).toMatchObject({
      sourceId: revised.sources[0]!.id,
      attachmentId,
      status: 'found',
    })
    revised.sources[0]!.excerpt = 'An excerpt absent from the file'
    await send({ operation: 'save_plan', expectedRevision: 2, plan: revised })
    expect((await send({ operation: 'audit_sources' })).sources[0].status).toBe('not_found')
    expect((await send({ operation: 'status' })).sourcePreparation[0].status).toBe(
      'excerpt_mismatch',
    )
    revised.sources[0]!.excerpt = ''
    await send({ operation: 'save_plan', expectedRevision: 3, plan: revised })
    expect((await send({ operation: 'audit_sources' })).sources[0].status).toBe('empty_excerpt')
    expect((await send({ operation: 'status' })).sourcePreparation[0].status).toBe(
      'excerpt_missing',
    )
    const docHash = createHash('sha256').update(documentId).digest('hex')
    writeFileSync(
      join(userDataPath, 'presentation-attachments', docHash, attachmentId, 'text.txt'),
      'tampered research source',
    )
    expect(await send({ operation: 'audit_sources' })).toEqual({ error: 'invalid_state' })
    expect(
      await send({
        operation: 'attachment_metadata',
        attachmentId,
        documentId: 'another-document',
      }),
    ).toEqual({ error: 'not_found' })
    await send({ operation: 'attachment_delete', attachmentId })
    expect((await send({ operation: 'status' })).sourcePreparation[0].status).toBe('missing')
    expect((await send({ operation: 'audit_sources' })).sources[0].status).toBe('missing')
  })
  it('uploads a real PDF, parses it and reads it after PC recreation without a presentation plan', async () => {
    const userDataPath = mkdtempSync(join(tmpdir(), 'ppt-attachment-integration-'))
    roots.push(userDataPath)
    let service = createPresentationService({ userDataPath })
    const bytes = Buffer.from(buildPdfFixture('Evidence from the uploaded PDF'))
    const attachmentId = createHash('sha256').update(bytes).digest('hex')
    const documentId = 'office-document-A'
    const send = async (body: Record<string, unknown>) =>
      decode(await service({ documentId, ...body }, signal()))
    expect(
      await send({
        operation: 'attachment_begin',
        attachmentId,
        name: 'evidence.pdf',
        sizeBytes: bytes.length,
        sha256: attachmentId,
      }),
    ).toMatchObject({ attachmentId, status: 'uploading', receivedBytes: 0 })
    const chunk = {
      operation: 'attachment_chunk',
      attachmentId,
      offset: 0,
      base64: bytes.toString('base64'),
    }
    expect(await send(chunk)).toMatchObject({ receivedBytes: bytes.length })
    service = createPresentationService({ userDataPath })
    expect(await send(chunk)).toMatchObject({ receivedBytes: bytes.length })
    expect(await send({ operation: 'attachment_finish', attachmentId })).toMatchObject({
      status: 'ready',
      kind: 'text',
    })
    service = createPresentationService({ userDataPath })
    const result = await send({
      operation: 'attachment_read',
      attachmentId,
      offset: 0,
      maxChars: 24000,
    })
    expect(result).toMatchObject({
      attachmentId,
      sourceUri: `attachment:${attachmentId}`,
      offset: 0,
    })
    expect(result.text).toContain('Evidence from the uploaded PDF')
    expect((await send({ operation: 'attachment_list' })).attachments).toHaveLength(1)
    expect(
      await send({
        operation: 'attachment_read',
        documentId: 'office-document-B',
        attachmentId,
        offset: 0,
        maxChars: 24000,
      }),
    ).toEqual({ error: 'not_found' })
    expect(
      await send({
        operation: 'attachment_read',
        attachmentId,
        offset: 0,
        maxChars: 24000,
        path: '/etc/passwd',
      }),
    ).toEqual({ error: 'invalid_request' })
  })
})

// Exercise the browser protocol client against the real PC service and parser.
it('recovers a lost acknowledgement then exposes real DOCX text to the Agent after restart', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'ppt-attachment-client-'))
  roots.push(userDataPath)
  let service = createPresentationService({ userDataPath })
  let loseAck = true
  const options = {
    available: () => true,
    documentId: async () => 'document-client',
    request: async (body: unknown, abort?: AbortSignal) => {
      const result = await service(body, abort ?? signal())
      if ((body as { operation: string }).operation === 'attachment_chunk' && loseAck) {
        loseAck = false
        throw new Error('connection_lost')
      }
      return new Response(Buffer.from(result))
    },
    vfs: new InMemoryVfs(),
  }
  let client = createPresentationAttachmentSkill(options)
  const bytes = new Uint8Array(await buildDocxFixture())
  await expect(client.upload('research.docx', Promise.resolve(bytes.buffer))).rejects.toThrow(
    'connection_lost',
  )
  client.clear()
  service = createPresentationService({ userDataPath })
  client = createPresentationAttachmentSkill({ ...options, vfs: new InMemoryVfs() })
  await client.upload('research.docx', Promise.resolve(bytes.buffer))
  client.clear()
  service = createPresentationService({ userDataPath })
  client = createPresentationAttachmentSkill({ ...options, vfs: new InMemoryVfs() })
  const listed = await client.executeTool({
    id: 'list',
    name: 'list_presentation_attachments',
    input: {},
  })
  expect(listed.isError).not.toBe(true)
  const [attachment] = JSON.parse(listed.output).attachments
  expect(attachment).toMatchObject({ name: 'research.docx', status: 'ready' })
  const read = await client.executeTool({
    id: 'read',
    name: 'read_presentation_attachment',
    input: { attachment_id: attachment.attachmentId },
  })
  expect(read.isError).not.toBe(true)
  expect(JSON.parse(read.output).text).toContain('hello docx')
  expect(JSON.parse(read.output).sourceUri).toBe(`attachment:${attachment.attachmentId}`)
})

it('retains the unreadable page of the real P0-11 NACA scan through PC restart and Agent listing', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'ppt-p0-11-scan-'))
  roots.push(userDataPath)
  const bytes = readFileSync(
    new URL(
      '../../../docs/product/ppt-benchmark-materials/PPT-P0-11/naca-rm-l50b01-1950-real-scan.pdf',
      import.meta.url,
    ),
  )
  const attachmentId = createHash('sha256').update(bytes).digest('hex')
  let service = createPresentationService({ userDataPath })
  const options = {
    available: () => true,
    documentId: async () => 'p0-11-scan-document',
    request: async (body: unknown, abort?: AbortSignal) =>
      new Response(Buffer.from(await service(body, abort ?? signal()))),
    vfs: new InMemoryVfs(),
  }
  let client = createPresentationAttachmentSkill(options)
  await client.upload(
    'naca-rm-l50b01-1950-real-scan.pdf',
    Promise.resolve(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)),
  )
  client.clear()
  service = createPresentationService({ userDataPath })
  client = createPresentationAttachmentSkill({ ...options, vfs: new InMemoryVfs() })
  const listed = await client.executeTool({
    id: 'list',
    name: 'list_presentation_attachments',
    input: {},
  })
  expect(listed.isError, listed.output).not.toBe(true)
  expect(JSON.parse(listed.output).attachments).toEqual([
    expect.objectContaining({
      attachmentId,
      status: 'ready',
      sectionCount: 30,
      pagesWithoutExtractedText: [2],
      pagesWithSparseExtractedText: expect.arrayContaining([24, 30]),
      // Page 26 has hundreds of OCR characters despite appearing blank; keep its OCR warning.
      pagesWithFullPageImage: Array.from({ length: 30 }, (_, index) => index + 1),
      pagesWithInvisibleTextLayer: Array.from({ length: 30 }, (_, index) => index + 1).filter(
        (page) => page !== 2,
      ),
    }),
  ])
  const read = await client.executeTool({
    id: 'read',
    name: 'read_presentation_attachment',
    input: { attachment_id: attachmentId, max_chars: 4000 },
  })
  expect(read.isError, read.output).not.toBe(true)
  const page = JSON.parse(read.output)
  expect(page.pageSpans).toEqual(
    expect.arrayContaining([
      expect.objectContaining({ locator: '第 2 页', start: expect.any(Number), imageBacked: true }),
      expect.objectContaining({
        locator: '第 3 页',
        start: expect.any(Number),
        imageBacked: true,
        invisibleTextLayer: true,
      }),
    ]),
  )
  const blank = page.pageSpans.find((span: { locator: string }) => span.locator === '第 2 页')
  expect(blank.start).toBe(blank.end)
  expect(page.text).toContain('NATIONAL ADVISORY COMMITTEE FOR AERONAUTICS')
}, 60_000)

it('resumes the frozen P0-16 real PDF after a lost first chunk acknowledgement and PC restart', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'ppt-p0-16-upload-'))
  roots.push(userDataPath)
  const bytes = readFileSync(
    new URL(
      '../../../docs/product/ppt-benchmark-materials/PPT-P0-16/deardorff-2020-article.pdf',
      import.meta.url,
    ),
  )
  const expectedId = createHash('sha256').update(bytes).digest('hex')
  let service = createPresentationService({ userDataPath })
  let loseAck = true
  const offsets: number[] = []
  const options = {
    available: () => true,
    documentId: async () => 'p0-16-document',
    request: async (body: unknown, abort?: AbortSignal) => {
      const input = body as { operation: string; offset?: number }
      const result = await service(body, abort ?? signal())
      if (input.operation === 'attachment_chunk') {
        offsets.push(input.offset!)
        if (loseAck) {
          loseAck = false
          throw new Error('connection_lost')
        }
      }
      return new Response(Buffer.from(result))
    },
    vfs: new InMemoryVfs(),
  }
  let client = createPresentationAttachmentSkill(options)
  await expect(
    client.upload(
      'deardorff-2020-article.pdf',
      Promise.resolve(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)),
    ),
  ).rejects.toThrow('connection_lost')
  client.clear()
  service = createPresentationService({ userDataPath })
  client = createPresentationAttachmentSkill({ ...options, vfs: new InMemoryVfs() })
  await client.upload(
    'deardorff-2020-article.pdf',
    Promise.resolve(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)),
  )
  expect(offsets).toEqual([0, 131072, 262144])
  client.clear()
  service = createPresentationService({ userDataPath })
  client = createPresentationAttachmentSkill({ ...options, vfs: new InMemoryVfs() })
  const listed = await client.executeTool({
    id: 'list',
    name: 'list_presentation_attachments',
    input: {},
  })
  expect(listed.isError, listed.output).not.toBe(true)
  expect(JSON.parse(listed.output).attachments).toEqual([
    expect.objectContaining({ attachmentId: expectedId, sizeBytes: bytes.length, status: 'ready' }),
  ])
  const read = await client.executeTool({
    id: 'read',
    name: 'read_presentation_attachment',
    input: { attachment_id: expectedId },
  })
  expect(read.isError, read.output).not.toBe(true)
  expect(JSON.parse(read.output)).toMatchObject({
    attachmentId: expectedId,
    sourceUri: `attachment:${expectedId}`,
  })
  expect(JSON.parse(read.output).text.length).toBeGreaterThan(100)
}, 60_000)

it('streams a 50 MiB PDF through the client beyond the session VFS limit', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'ppt-attachment-large-'))
  roots.push(userDataPath)
  const service = createPresentationService({ userDataPath })
  const bytes = new Uint8Array(50 * 1024 * 1024).fill(32)
  bytes.set(buildPdfFixture('Large PDF evidence'))
  let chunks = 0
  const vfs = new InMemoryVfs()
  const client = createPresentationAttachmentSkill({
    available: () => true,
    documentId: async () => 'document-large',
    vfs,
    request: async (body, abort) => {
      if ((body as { operation: string }).operation === 'attachment_chunk') chunks++
      return new Response(Buffer.from(await service(body, abort ?? signal())))
    },
  })
  await client.upload('large.pdf', Promise.resolve(bytes.buffer))
  expect(chunks).toBe(400)
  expect(vfs.list('/home/user')).not.toContain('/home/user/large.pdf')
  const listed = await client.executeTool({
    id: 'list',
    name: 'list_presentation_attachments',
    input: {},
  })
  expect(JSON.parse(listed.output).attachments[0]).toMatchObject({
    sizeBytes: bytes.length,
    status: 'ready',
  })
}, 60_000)

const realPdfPath = process.env.WISWORK_P0_10_PDF
if (realPdfPath) {
  it('uploads and reads a real near-limit PDF beyond the first million characters', async () => {
    const bytes = readFileSync(realPdfPath)
    expect(bytes.length).toBeGreaterThan(45 * 1024 * 1024)
    expect(bytes.length).toBeLessThanOrEqual(50 * 1024 * 1024)
    const attachmentId = createHash('sha256').update(bytes).digest('hex')
    expect(attachmentId).toBe('cadeefed4b0f0627384b6b7f3730afc729570270b8794b71759f9dc6511a36b2')
    const userDataPath = mkdtempSync(join(tmpdir(), 'ppt-p0-10-real-'))
    roots.push(userDataPath)
    const service = createPresentationService({ userDataPath })
    const client = createPresentationAttachmentSkill({
      available: () => true,
      documentId: async () => 'document-p0-10',
      vfs: new InMemoryVfs(),
      request: async (body, abort) =>
        new Response(Buffer.from(await service(body, abort ?? signal()))),
    })
    await client.upload(
      'SROCC_FullReport_FINAL.pdf',
      Promise.resolve(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)),
    )
    const listed = await client.executeTool({
      id: 'list',
      name: 'list_presentation_attachments',
      input: {},
    })
    expect(listed.isError, listed.output).not.toBe(true)
    expect(JSON.parse(listed.output).attachments[0]).toMatchObject({
      attachmentId,
      sizeBytes: bytes.length,
      status: 'ready',
      totalChars: expect.any(Number),
    })
    expect(JSON.parse(listed.output).attachments[0].totalChars).toBeGreaterThan(4_000_000)
    const read = await client.executeTool({
      id: 'read',
      name: 'read_presentation_attachment',
      input: { attachment_id: attachmentId, offset: 4_000_000, max_chars: 2000 },
    })
    expect(read.isError, read.output).not.toBe(true)
    expect(JSON.parse(read.output)).toMatchObject({
      offset: 4_000_000,
      totalChars: expect.any(Number),
    })
    expect(JSON.parse(read.output).text.length).toBe(2000)
    const window = JSON.parse(read.output) as {
      text: string
      pageSpans: { locator: string; start: number; end: number }[]
    }
    const excerpt = window.text.slice(100, 180).trim()
    expect(excerpt.length).toBeGreaterThan(20)
    const plan = benchmarkPlan()
    plan.sources[0]!.uri = `attachment:${attachmentId}`
    plan.sources[0]!.excerpt = excerpt
    const send = async (body: Record<string, unknown>) =>
      decode(
        await service(
          { documentId: 'document-p0-10', projectId: plan.projectId, ...body },
          signal(),
        ),
      )
    await send({ operation: 'save_plan', expectedRevision: 0, plan })
    const audit = await send({ operation: 'audit_sources' })
    expect(audit.sources[0]).toMatchObject({
      sourceId: plan.sources[0]!.id,
      attachmentId,
      status: 'found',
      offset: expect.any(Number),
      locator: expect.stringMatching(/^第 \d+ 页$/),
    })
    expect(audit.sources[0].offset).toBeGreaterThan(1_000_000)
    expect(window.pageSpans.some((span) => span.locator === audit.sources[0].locator)).toBe(true)
  }, 180_000)
}
