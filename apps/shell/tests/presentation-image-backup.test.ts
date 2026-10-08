import { afterEach, expect, it } from 'vitest'
import { mkdtemp, rm, writeFile, symlink } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { createPresentationAttachmentService } from '../src/main/presentation-attachments'
import { createPresentationService } from '../src/main/presentation-service'
const png = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII=',
  'base64',
)
const jpeg = Buffer.from([255, 216, 255, 192, 0, 11, 8, 0, 1, 0, 1, 1, 1, 17, 0, 255, 217])
const hash = (v: string | Uint8Array) => createHash('sha256').update(v).digest('hex')
const dirs: string[] = []
afterEach(async () => {
  await Promise.all(dirs.splice(0).map((p) => rm(p, { recursive: true, force: true })))
})
async function setup() {
  const userDataPath = await mkdtemp(join(tmpdir(), 'image-original-'))
  dirs.push(userDataPath)
  const make = () =>
    createPresentationAttachmentService({
      userDataPath,
      normalizeImage: async () => ({ bytes: png, width: 1, height: 1 }),
    })
  const service = make(),
    attachmentId = hash(jpeg)
  const call = (body: Record<string, unknown>, signal = new AbortController().signal) =>
    service({ documentId: 'doc', ...body }, signal)
  await call({
    operation: 'attachment_begin',
    attachmentId,
    sha256: attachmentId,
    name: 'source.jpeg',
    sizeBytes: jpeg.length,
  })
  await call({
    operation: 'attachment_chunk',
    attachmentId,
    offset: 0,
    base64: jpeg.toString('base64'),
  })
  await call({ operation: 'attachment_finish', attachmentId })
  const body = {
    operation: 'attachment_original',
    documentId: 'doc',
    attachmentId,
    offset: 0,
    length: 128 * 1024,
  }
  return {
    call,
    make,
    userDataPath,
    body,
    path: join(userDataPath, 'presentation-attachments', hash('doc'), attachmentId, 'raw.jpeg'),
  }
}
it('reads original JPEG bytes after restart even though normalized cache is PNG', async () => {
  const f = await setup()
  expect(await f.make()(f.body, new AbortController().signal)).toEqual({
    attachmentId: hash(jpeg),
    offset: 0,
    sizeBytes: jpeg.length,
    sha256: hash(jpeg),
    mime: 'image/jpeg',
    base64: jpeg.toString('base64'),
  })
  expect(await f.call({ ...f.body, offset: 2, length: 3 })).toMatchObject({
    offset: 2,
    base64: jpeg.subarray(2, 5).toString('base64'),
  })
})
it('rejects cross-document, invalid bounds, cancellation and damaged or symlinked originals', async () => {
  const f = await setup()
  await expect(f.call({ ...f.body, documentId: 'other' })).rejects.toThrow('not_found')
  for (const patch of [
    { offset: -1 },
    { offset: jpeg.length + 1 },
    { length: 0 },
    { length: 128 * 1024 + 1 },
  ])
    await expect(f.call({ ...f.body, ...patch })).rejects.toThrow('invalid_request')
  const c = new AbortController()
  c.abort()
  await expect(f.call(f.body, c.signal)).rejects.toThrow('aborted')
  await writeFile(f.path, Buffer.alloc(jpeg.length))
  await expect(f.call(f.body)).rejects.toThrow('digest_mismatch')
  await rm(f.path)
  await symlink('/etc/passwd', f.path)
  await expect(f.call(f.body)).rejects.toThrow('invalid_state')
})

import { createPresentationImageBackup } from '../../office-addin/src/skills/powerpoint/presentation-image-backup'
it('client reuses differently named source and verifies original bytes through service restart', async () => {
  const f = await setup()
  const client = () =>
    createPresentationImageBackup({
      available: () => true,
      request: async (body, signal) =>
        new Response(
          (await createPresentationService({
            userDataPath: f.userDataPath,
            normalizeImage: async () => ({ bytes: png, width: 1, height: 1 }),
          })(body, signal ?? new AbortController().signal)) as Uint8Array<ArrayBuffer>,
        ),
    })
  const metadata = await client().save('doc', jpeg.toString('base64'))
  expect(metadata).toEqual({ attachmentId: hash(jpeg), sizeBytes: jpeg.length, mime: 'image/jpeg' })
  expect(await client().load('doc', metadata)).toBe(jpeg.toString('base64'))
  const pngMetadata = await client().save('doc', png.toString('base64'))
  expect(await client().load('doc', pngMetadata)).toBe(png.toString('base64'))
  await expect(client().load('other', metadata)).rejects.toThrow()
  await writeFile(f.path, Buffer.alloc(jpeg.length))
  await expect(client().load('doc', metadata)).rejects.toThrow()
})
it('does not expose incomplete uploads or non-image attachments as original images', async () => {
  const f = await setup()
  const id = hash(png)
  await f.call({
    operation: 'attachment_begin',
    attachmentId: id,
    sha256: id,
    name: 'pending.png',
    sizeBytes: png.length,
  })
  await expect(f.call({ ...f.body, attachmentId: id })).rejects.toThrow('invalid_state')
  const text = Buffer.from('reference'),
    textId = hash(text)
  await f.call({
    operation: 'attachment_begin',
    attachmentId: textId,
    sha256: textId,
    name: 'source.txt',
    sizeBytes: text.length,
  })
  await f.call({
    operation: 'attachment_chunk',
    attachmentId: textId,
    offset: 0,
    base64: text.toString('base64'),
  })
  await f.call({ operation: 'attachment_finish', attachmentId: textId })
  await expect(f.call({ ...f.body, attachmentId: textId })).rejects.toThrow('invalid_state')
})
it('backup save works after the source attachment count reaches its limit', async () => {
  const f = await setup()
  for (let i = 0; i < 32; i++) {
    const id = hash(String(i))
    await f.call({
      operation: 'attachment_begin',
      attachmentId: id,
      sha256: id,
      name: `source-${i}.txt`,
      sizeBytes: 1,
    })
  }
  const client = createPresentationImageBackup({
    available: () => true,
    request: async (body, signal) =>
      new Response(JSON.stringify(await f.call(body, signal ?? new AbortController().signal))),
  })
  await expect(client.save('doc', png.toString('base64'))).resolves.toMatchObject({
    attachmentId: hash(png),
  })
  const result = (await f.call({ operation: 'attachment_list_assets' })) as {
    attachments: unknown[]
  }
  expect(result.attachments).toHaveLength(32)
  const page = (await f.call({
    operation: 'attachment_list_assets',
    after: (result as { nextAfter: string }).nextAfter,
  })) as { attachments: { attachmentId: string }[] }
  expect(
    [...(result.attachments as { attachmentId: string }[]), ...page.attachments].map(
      (item) => item.attachmentId,
    ),
  ).toContain(hash(png))
})
