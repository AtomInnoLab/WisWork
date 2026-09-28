import { afterEach, describe, expect, it } from 'vitest'
import { mkdtemp, rm, writeFile, symlink } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import JSZip from 'jszip'
import { buildPdfFixture } from '../../../packages/file-parse/tests/helpers/fixtures'
import { createPresentationAttachmentService } from '../src/main/presentation-attachments'
const dirs: string[] = []
afterEach(async () => {
  await Promise.all(dirs.splice(0).map((p) => rm(p, { recursive: true, force: true })))
})
const hash = (v: Uint8Array | string) => createHash('sha256').update(v).digest('hex')
async function setup(parse?: Parameters<typeof createPresentationAttachmentService>[0]['parse']) {
  const userDataPath = await mkdtemp(join(tmpdir(), 'ppt-attachments-'))
  dirs.push(userDataPath)
  const service = createPresentationAttachmentService({ userDataPath, ...(parse ? { parse } : {}) })
  const call = (body: Record<string, unknown>, signal = new AbortController().signal) =>
    service({ documentId: 'doc-1', ...body }, signal)
  return { call, userDataPath }
}
async function upload(
  call: Awaited<ReturnType<typeof setup>>['call'],
  bytes: Buffer,
  name = 'source.txt',
) {
  const attachmentId = hash(bytes)
  await call({
    operation: 'attachment_begin',
    attachmentId,
    sha256: attachmentId,
    name,
    sizeBytes: bytes.length,
  })
  await call({
    operation: 'attachment_chunk',
    attachmentId,
    offset: 0,
    base64: bytes.toString('base64'),
  })
  return attachmentId
}
describe('durable presentation attachments', () => {
  it('returns the PDF page locator for a matched excerpt and verifies the page index', async () => {
    const { call, userDataPath } = await setup()
    const id = await upload(
      call,
      Buffer.from(buildPdfFixture(['First page evidence', 'Second page finding'])),
      'study.pdf',
    )
    expect(await call({ operation: 'attachment_finish', attachmentId: id })).toMatchObject({
      status: 'ready',
      sectionCount: 2,
    })
    expect(
      await call({
        operation: 'attachment_match_excerpt',
        attachmentId: id,
        excerpt: 'Second page finding',
      }),
    ).toMatchObject({ status: 'found', locator: '第 2 页' })
    expect(
      await call({
        operation: 'attachment_match_excerpt',
        attachmentId: id,
        excerpt: 'evidence\n\nSecond',
      }),
    ).toMatchObject({ status: 'not_found' })
    expect(
      await call({ operation: 'attachment_read', attachmentId: id, offset: 0, maxChars: 100 }),
    ).toMatchObject({
      pageSpans: [
        { locator: '第 1 页', start: 0, end: 19 },
        { locator: '第 2 页', start: 21, end: 40 },
      ],
    })
    const sectionPath = join(
      userDataPath,
      'presentation-attachments',
      hash('doc-1'),
      id,
      'sections.json',
    )
    await writeFile(sectionPath, '[]')
    await expect(
      call({
        operation: 'attachment_match_excerpt',
        attachmentId: id,
        excerpt: 'First page evidence',
      }),
    ).rejects.toThrow('invalid_state')
  })
  it('prefers the cited page when the same PDF excerpt appears more than once', async () => {
    const { call } = await setup()
    const id = await upload(
      call,
      Buffer.from(buildPdfFixture(['Repeated evidence', 'Repeated evidence'])),
      'repeat.pdf',
    )
    await call({ operation: 'attachment_finish', attachmentId: id })
    expect(
      await call({
        operation: 'attachment_match_excerpt',
        attachmentId: id,
        excerpt: 'Repeated evidence',
        locator: '第 2 页',
      }),
    ).toMatchObject({ status: 'found', locator: '第 2 页', offset: 19 })
    expect(
      await call({
        operation: 'attachment_match_excerpt',
        attachmentId: id,
        excerpt: 'Repeated evidence',
        locator: '第 3 页',
      }),
    ).toMatchObject({ status: 'found', locator: '第 1 页', offset: 0 })
  })
  it('fetches an HTML URL into a document-scoped, deduplicated source snapshot', async () => {
    const userDataPath = await mkdtemp(join(tmpdir(), 'ppt-webpage-'))
    dirs.push(userDataPath)
    let requests = 0
    const fetchPage = async () => {
      requests++
      return new Response(
        '<html><body><h1>Study</h1><p>Result &amp; method</p><script>ignore()</script></body></html>',
        {
          headers: { 'content-type': 'text/html; charset=utf-8' },
        },
      )
    }
    const service = createPresentationAttachmentService({ userDataPath, fetchPage })
    const call = (body: Record<string, unknown>) =>
      service({ documentId: 'doc-1', ...body }, new AbortController().signal)
    const url = 'https://8.8.8.8/study?session=private#section'
    const first = (await call({ operation: 'attachment_import_webpage', url })) as {
      attachmentId: string
      source: string
      kind: string
    }
    expect(first).toMatchObject({
      source: 'https://8.8.8.8/study',
      sourceUrlHash: hash(url),
      kind: 'text',
    })
    expect((first as { retrievedAt?: number }).retrievedAt).toBeGreaterThan(0)
    expect(requests).toBe(1)
    expect(await call({ operation: 'attachment_import_webpage', url })).toMatchObject({
      attachmentId: first.attachmentId,
    })
    expect(requests).toBe(1)
    expect(
      await call({
        operation: 'attachment_read',
        attachmentId: first.attachmentId,
        offset: 0,
        maxChars: 100,
      }),
    ).toMatchObject({
      text: 'Study\nResult & method',
      sourceUri: `attachment:${first.attachmentId}`,
    })
    await expect(
      service(
        {
          documentId: 'doc-2',
          operation: 'attachment_read',
          attachmentId: first.attachmentId,
          offset: 0,
          maxChars: 100,
        },
        new AbortController().signal,
      ),
    ).rejects.toThrow('not_found')
    await expect(
      call({ operation: 'attachment_import_webpage', url: 'https://8.8.8.8/other' }),
    ).rejects.toThrow('remote_webpage_source_conflict')
  })

  it('rejects unsafe or non-HTML webpage sources and oversized responses', async () => {
    const userDataPath = await mkdtemp(join(tmpdir(), 'ppt-webpage-'))
    dirs.push(userDataPath)
    let response = new Response('plain', { headers: { 'content-type': 'text/plain' } })
    const service = createPresentationAttachmentService({
      userDataPath,
      fetchPage: async () => response,
    })
    const call = (url: string) =>
      service(
        { documentId: 'doc-1', operation: 'attachment_import_webpage', url },
        new AbortController().signal,
      )
    await expect(call('http://127.0.0.1/private')).rejects.toThrow('remote_webpage_unavailable')
    await expect(call('https://8.8.8.8/page')).rejects.toThrow('remote_webpage_unavailable')
    response = new Response('html', {
      headers: { 'content-type': 'text/html', 'content-length': String(6 * 1024 * 1024) },
    })
    await expect(call('https://8.8.8.8/page')).rejects.toThrow('quota_exceeded')
  })
  it('uses the HTTP charset for a fetched non-UTF-8 page', async () => {
    const userDataPath = await mkdtemp(join(tmpdir(), 'ppt-webpage-charset-'))
    dirs.push(userDataPath)
    const raw = Buffer.concat([
      Buffer.from('<html><body><p>'),
      Buffer.from('d6d0cec4', 'hex'),
      Buffer.from('</p></body></html>'),
    ])
    const service = createPresentationAttachmentService({
      userDataPath,
      fetchPage: async () =>
        new Response(raw, { headers: { 'content-type': 'text/html; charset=gbk' } }),
    })
    const call = (body: Record<string, unknown>) =>
      service({ documentId: 'doc-1', ...body }, new AbortController().signal)
    const imported = (await call({
      operation: 'attachment_import_webpage',
      url: 'https://8.8.8.8/gbk',
    })) as { attachmentId: string }
    expect(
      await call({
        operation: 'attachment_read',
        attachmentId: imported.attachmentId,
        offset: 0,
        maxChars: 20,
      }),
    ).toMatchObject({ text: '中文' })
  })
  it('resumes after a lost acknowledgement and restart, and reads bounded durable text', async () => {
    const { call, userDataPath } = await setup()
    const bytes = Buffer.from('资料 evidence')
    const id = await upload(call, bytes)
    expect(
      await call({
        operation: 'attachment_chunk',
        attachmentId: id,
        offset: 0,
        base64: bytes.toString('base64'),
      }),
    ).toMatchObject({ receivedBytes: bytes.length })
    const service = createPresentationAttachmentService({ userDataPath })
    expect(
      await service(
        { operation: 'attachment_finish', documentId: 'doc-1', attachmentId: id },
        new AbortController().signal,
      ),
    ).toMatchObject({ status: 'ready', totalChars: 11 })
    expect(
      await call({ operation: 'attachment_read', attachmentId: id, offset: 0, maxChars: 2 }),
    ).toMatchObject({ text: '资料', sourceUri: `attachment:${id}` })
    await expect(
      call({
        operation: 'attachment_read',
        documentId: 'doc-2',
        attachmentId: id,
        offset: 0,
        maxChars: 2,
      }),
    ).rejects.toThrow('not_found')
  })
  it('extracts an actual DOCX via the shared parser', async () => {
    const { call } = await setup()
    const zip = new JSZip()
    zip.file(
      '[Content_Types].xml',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    )
    zip.file(
      'word/document.xml',
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Evidence DOCX</w:t></w:r></w:p></w:body></w:document>',
    )
    const id = await upload(
      call,
      await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
      '资料.docx',
    )
    expect(await call({ operation: 'attachment_finish', attachmentId: id })).toMatchObject({
      status: 'ready',
    })
    expect(
      await call({ operation: 'attachment_read', attachmentId: id, offset: 0, maxChars: 100 }),
    ).toMatchObject({ text: 'Evidence DOCX' })
  })
  it('stores a webpage snapshot as readable text while retaining the original HTML', async () => {
    const { call } = await setup()
    const html = Buffer.from(
      '<html><body><h1>Public study</h1><p>Sample &amp; methods</p><script>private()</script></body></html>',
    )
    const id = await upload(call, html, 'study.html')
    expect(await call({ operation: 'attachment_finish', attachmentId: id })).toMatchObject({
      status: 'ready',
      kind: 'text',
      sha256: id,
    })
    expect(
      await call({ operation: 'attachment_read', attachmentId: id, offset: 0, maxChars: 100 }),
    ).toMatchObject({ text: 'Public study\nSample & methods', sourceUri: `attachment:${id}` })
  })
  it('rejects conflicting chunks, wrong digest, unknown keys, unsafe names and unsupported files', async () => {
    const { call } = await setup()
    const id = await upload(call, Buffer.from('abc'))
    await expect(
      call({
        operation: 'attachment_chunk',
        attachmentId: id,
        offset: 0,
        base64: Buffer.from('bad').toString('base64'),
      }),
    ).rejects.toThrow('attachment_conflict')
    await expect(call({ operation: 'attachment_list', path: '/etc/passwd' })).rejects.toThrow(
      'invalid_request',
    )
    for (const name of ['../a.txt', 'a.exe'])
      await expect(
        call({ operation: 'attachment_begin', attachmentId: id, sha256: id, name, sizeBytes: 3 }),
      ).rejects.toThrow()
    const wrong = 'a'.repeat(64)
    await call({
      operation: 'attachment_begin',
      attachmentId: wrong,
      sha256: wrong,
      name: 'wrong.txt',
      sizeBytes: 3,
    })
    await call({ operation: 'attachment_chunk', attachmentId: wrong, offset: 0, base64: 'YWJj' })
    await expect(call({ operation: 'attachment_finish', attachmentId: wrong })).rejects.toThrow(
      'digest_mismatch',
    )
  })
  it('retries failed parsing and never stores raw parser errors', async () => {
    let fails = true
    const { call, userDataPath } = await setup(async () =>
      fails
        ? { ok: false, kind: 'text', error: '/private/token secret' }
        : { ok: true, kind: 'text', text: 'recovered' },
    )
    const id = await upload(call, Buffer.from('abc'))
    expect(await call({ operation: 'attachment_finish', attachmentId: id })).toMatchObject({
      status: 'failed',
      error: 'parse_failed',
    })
    fails = false
    expect(await call({ operation: 'attachment_finish', attachmentId: id })).toMatchObject({
      status: 'ready',
    })
    const dir = join(userDataPath, 'presentation-attachments', hash('doc-1'), id)
    const cache = join(dir, 'text.txt')
    await writeFile(cache, 'tampered')
    await expect(
      call({ operation: 'attachment_read', attachmentId: id, offset: 0, maxChars: 10 }),
    ).rejects.toThrow('invalid_state')
  })
  it('rejects cancellation, quota exhaustion and symlinked storage', async () => {
    const { call, userDataPath } = await setup()
    const ac = new AbortController()
    ac.abort()
    await expect(call({ operation: 'attachment_list' }, ac.signal)).rejects.toThrow('aborted')
    for (let i = 0; i < 10; i++) {
      const id = hash(String(i))
      await call({
        operation: 'attachment_begin',
        attachmentId: id,
        sha256: id,
        name: `${i}.txt`,
        sizeBytes: 10 * 1024 * 1024,
      })
    }
    const id = hash('overflow')
    await expect(
      call({
        operation: 'attachment_begin',
        attachmentId: id,
        sha256: id,
        name: 'overflow.txt',
        sizeBytes: 1,
      }),
    ).rejects.toThrow('quota_exceeded')
    const root = join(userDataPath, 'presentation-attachments')
    await rm(root, { recursive: true })
    await symlink(tmpdir(), root)
    await expect(call({ operation: 'attachment_list' })).rejects.toThrow('invalid_state')
  })
  it('deletes only the named document attachment and releases declared capacity', async () => {
    const { call } = await setup()
    for (let i = 0; i < 10; i++) {
      const id = hash(String(i))
      await call({
        operation: 'attachment_begin',
        attachmentId: id,
        sha256: id,
        name: `${i}.txt`,
        sizeBytes: 10 * 1024 * 1024,
      })
    }
    const removed = hash('0')
    await expect(
      call({ operation: 'attachment_delete', documentId: 'other', attachmentId: removed }),
    ).rejects.toThrow('not_found')
    expect(await call({ operation: 'attachment_delete', attachmentId: removed })).toEqual({
      attachmentId: removed,
      deleted: true,
    })
    await expect(call({ operation: 'attachment_delete', attachmentId: removed })).rejects.toThrow(
      'not_found',
    )
    const replacement = hash('replacement')
    expect(
      await call({
        operation: 'attachment_begin',
        attachmentId: replacement,
        sha256: replacement,
        name: 'replacement.txt',
        sizeBytes: 10 * 1024 * 1024,
      }),
    ).toMatchObject({ attachmentId: replacement })
  })
  it('recovers a partially persisted chunk and serializes separate service instances', async () => {
    const { call, userDataPath } = await setup()
    const data = Buffer.from('abcdef')
    const id = hash(data)
    const begin = {
      operation: 'attachment_begin',
      attachmentId: id,
      sha256: id,
      name: 'a.txt',
      sizeBytes: data.length,
    }
    const second = createPresentationAttachmentService({ userDataPath })
    await Promise.all([
      call(begin),
      second({ ...begin, documentId: 'doc-1' }, new AbortController().signal),
    ])
    const raw = join(userDataPath, 'presentation-attachments', hash('doc-1'), id, 'raw.txt')
    await writeFile(raw, 'ab')
    expect(
      await call({
        operation: 'attachment_chunk',
        attachmentId: id,
        offset: 0,
        base64: data.toString('base64'),
      }),
    ).toMatchObject({ receivedBytes: 6 })
    expect(await call({ operation: 'attachment_finish', attachmentId: id })).toMatchObject({
      status: 'ready',
    })
  })
  it('rejects actual oversized DOCX inflation before invoking the parser', async () => {
    let calls = 0
    const { call } = await setup(async () => {
      calls++
      return { ok: true, kind: 'text', text: 'unsafe' }
    })
    const zip = new JSZip()
    zip.file('word/document.xml', 'a'.repeat(11 * 1024 * 1024))
    const data = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
    const id = await upload(call, data, 'oversized.docx')
    expect(await call({ operation: 'attachment_finish', attachmentId: id })).toMatchObject({
      status: 'failed',
      error: 'parse_failed',
    })
    expect(calls).toBe(0)
  })
  it('does not publish parsed text if cancelled while parsing and permits retry', async () => {
    const controller = new AbortController()
    let first = true
    const { call } = await setup(async () => {
      if (first) {
        first = false
        controller.abort()
      }
      return { ok: true, kind: 'text', text: 'valid' }
    })
    const id = await upload(call, Buffer.from('abc'))
    await expect(
      call({ operation: 'attachment_finish', attachmentId: id }, controller.signal),
    ).rejects.toThrow('aborted')
    expect(await call({ operation: 'attachment_finish', attachmentId: id })).toMatchObject({
      status: 'ready',
    })
  })
  it('rejects too many empty files and files above the 50MiB ceiling', async () => {
    const { call } = await setup()
    for (let i = 0; i < 32; i++) {
      const id = hash(String(i))
      await call({
        operation: 'attachment_begin',
        attachmentId: id,
        sha256: id,
        name: `${i}.txt`,
        sizeBytes: 0,
      })
    }
    const id = hash('extra')
    await expect(
      call({
        operation: 'attachment_begin',
        attachmentId: id,
        sha256: id,
        name: 'extra.txt',
        sizeBytes: 0,
      }),
    ).rejects.toThrow('quota_exceeded')
    await expect(
      call({
        operation: 'attachment_begin',
        attachmentId: id,
        sha256: id,
        name: 'large.pdf',
        sizeBytes: 50 * 1024 * 1024 + 1,
      }),
    ).rejects.toThrow('invalid_request')
  })
  it('reserves minimum storage for empty image uploads', async () => {
    const { call } = await setup()
    for (let i = 0; i < 2; i++) {
      const id = hash(`large-${i}`)
      await call({
        operation: 'attachment_begin',
        attachmentId: id,
        sha256: id,
        name: `large-${i}.txt`,
        sizeBytes: 50 * 1024 * 1024 - 1,
      })
    }
    const id = hash('empty-image')
    await expect(
      call({
        operation: 'attachment_begin',
        attachmentId: id,
        sha256: id,
        name: 'empty.png',
        sizeBytes: 0,
      }),
    ).rejects.toThrow('quota_exceeded')
  })
})
