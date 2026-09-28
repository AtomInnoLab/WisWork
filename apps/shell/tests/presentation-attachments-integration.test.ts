import { createHash } from 'node:crypto'
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
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
    expect((await send({ operation: 'status' })).sourcePreparation[0].status).toBe('excerpt_matched')
    expect(await send({ operation: 'audit_sources' })).toMatchObject({
      projectId: plan.projectId,
      planRevision: 1,
      sources: [{ sourceId: plan.sources[0]!.id, attachmentId, status: 'found' }],
      checks: {
        support: 'not_verified',
        sourceAuthority: 'not_verified',
        timeliness: 'not_verified',
      },
    })
    const revised = structuredClone(plan)
    revised.sources[0]!.excerpt = 'An excerpt absent from the file'
    await send({ operation: 'save_plan', expectedRevision: 1, plan: revised })
    expect((await send({ operation: 'audit_sources' })).sources[0].status).toBe('not_found')
    expect((await send({ operation: 'status' })).sourcePreparation[0].status).toBe('excerpt_mismatch')
    revised.sources[0]!.excerpt = ''
    await send({ operation: 'save_plan', expectedRevision: 2, plan: revised })
    expect((await send({ operation: 'audit_sources' })).sources[0].status).toBe('empty_excerpt')
    expect((await send({ operation: 'status' })).sourcePreparation[0].status).toBe('excerpt_missing')
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
