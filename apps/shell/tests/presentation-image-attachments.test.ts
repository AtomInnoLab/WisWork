import { afterEach, describe, expect, it, vi } from 'vitest'
import { mkdtemp, rm, writeFile, readFile, symlink, mkdir, utimes, access } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { createPresentationAttachmentService } from '../src/main/presentation-attachments'
const fixture = (name: string) => readFileSync(join(__dirname, 'fixtures/presentation-image', name))
const png = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII=',
  'base64',
)
const hash = (v: string | Uint8Array) => createHash('sha256').update(v).digest('hex')
const dirs: string[] = []
afterEach(async () => {
  await Promise.all(dirs.splice(0).map((path) => rm(path, { recursive: true, force: true })))
})
async function setup(
  normalizeImage: NonNullable<
    Parameters<typeof createPresentationAttachmentService>[0]['normalizeImage']
  > = async () => ({ bytes: png, width: 1, height: 1 }),
) {
  const userDataPath = await mkdtemp(join(tmpdir(), 'ppt-image-'))
  dirs.push(userDataPath)
  const service = createPresentationAttachmentService({ userDataPath, normalizeImage })
  const call = (body: Record<string, unknown>, signal = new AbortController().signal) =>
    service({ documentId: 'doc', ...body }, signal)
  return { call, userDataPath }
}
async function upload(
  call: Awaited<ReturnType<typeof setup>>['call'],
  bytes = png,
  name = 'photo.png',
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
describe('durable presentation image assets', () => {
  it('normalizes static GIF/WebP uploads into durable PNG assets', async () => {
    const { call } = await setup(async () => ({
      bytes: fixture('control.png'),
      width: 2,
      height: 3,
    }))
    for (const name of ['static.gif', 'static.webp']) {
      const id = await upload(call, fixture(name), name)
      await expect(
        call({ operation: 'attachment_finish', attachmentId: id }),
      ).resolves.toMatchObject({ status: 'ready' })
      await expect(
        call({ operation: 'attachment_asset', attachmentId: id }),
      ).resolves.toMatchObject({ mime: 'image/png', width: 2, height: 3 })
    }
  })
  it('records animation as a distinct unsupported input for uploads and remote images', async () => {
    const { call } = await setup()
    for (const name of ['animated.gif', 'animated.webp']) {
      const id = await upload(call, fixture(name), name)
      await expect(
        call({ operation: 'attachment_finish', attachmentId: id }),
      ).resolves.toMatchObject({
        status: 'failed',
        error: 'animated_image_unsupported',
      })
      const listed = (await call({ operation: 'attachment_list_assets' })) as {
        attachments: { attachmentId: string; status: string; error?: string }[]
      }
      expect(listed.attachments.find((item) => item.attachmentId === id)).toMatchObject({
        status: 'failed',
        error: 'animated_image_unsupported',
      })
    }
    const userDataPath = await mkdtemp(join(tmpdir(), 'ppt-animated-url-'))
    dirs.push(userDataPath)
    const service = createPresentationAttachmentService({
      userDataPath,
      fetchImage: async () =>
        new Response(fixture('animated.webp'), { headers: { 'content-type': 'image/webp' } }),
    })
    await expect(
      service(
        {
          documentId: 'doc',
          operation: 'attachment_import_url',
          url: 'https://93.184.216.34/animated.webp',
        },
        new AbortController().signal,
      ),
    ).rejects.toThrow('animated_image_unsupported')
  })
  it('converts an explicitly selected animated upload into a durable first-frame asset', async () => {
    const userDataPath = await mkdtemp(join(tmpdir(), 'ppt-first-frame-'))
    dirs.push(userDataPath)
    const firstFrame = vi.fn(async () => ({ bytes: fixture('control.png'), width: 2, height: 3 }))
    const service = createPresentationAttachmentService({
      userDataPath,
      normalizeFirstFrame: firstFrame,
    })
    const call = (body: Record<string, unknown>, documentId = 'doc') =>
      service({ documentId, ...body }, new AbortController().signal)
    const raw = fixture('animated.gif')
    const id = await upload(call, raw, 'animated.gif')
    expect(await call({ operation: 'attachment_finish', attachmentId: id })).toMatchObject({
      status: 'failed',
      error: 'animated_image_unsupported',
    })
    await expect(
      call({ operation: 'attachment_extract_first_frame', attachmentId: id }, 'other'),
    ).rejects.toThrow('not_found')
    const converted = await call({ operation: 'attachment_extract_first_frame', attachmentId: id })
    expect(converted).toMatchObject({
      attachmentId: id,
      sha256: id,
      status: 'ready',
      kind: 'image',
      animationHandling: 'first_frame',
      assetSha256: hash(fixture('control.png')),
    })
    expect(firstFrame).toHaveBeenCalledTimes(1)
    expect(await call({ operation: 'attachment_extract_first_frame', attachmentId: id })).toEqual(
      converted,
    )
    expect(firstFrame).toHaveBeenCalledTimes(1)
    expect(await call({ operation: 'attachment_asset', attachmentId: id })).toMatchObject({
      mime: 'image/png',
      width: 2,
      height: 3,
    })
    await expect(
      call({ operation: 'attachment_original', attachmentId: id, offset: 0, length: 1 }),
    ).rejects.toThrow('invalid_state')
    const docDir = join(userDataPath, 'presentation-attachments', hash('doc'), id)
    expect(await readFile(join(docDir, 'raw.gif'))).toEqual(raw)
    const reopened = createPresentationAttachmentService({ userDataPath })
    expect(
      await reopened(
        { documentId: 'doc', operation: 'attachment_metadata', attachmentId: id },
        new AbortController().signal,
      ),
    ).toEqual(converted)
  })
  it('keeps the uploaded animation retryable when first-frame decoding fails', async () => {
    const userDataPath = await mkdtemp(join(tmpdir(), 'ppt-first-frame-fail-'))
    dirs.push(userDataPath)
    const service = createPresentationAttachmentService({
      userDataPath,
      normalizeFirstFrame: async () => {
        throw new Error('parse_failed')
      },
    })
    const call = (body: Record<string, unknown>) =>
      service({ documentId: 'doc', ...body }, new AbortController().signal)
    const raw = fixture('animated.webp')
    const id = await upload(call, raw, 'animated.webp')
    await call({ operation: 'attachment_finish', attachmentId: id })
    await expect(
      call({ operation: 'attachment_extract_first_frame', attachmentId: id }),
    ).rejects.toThrow('parse_failed')
    expect(await call({ operation: 'attachment_metadata', attachmentId: id })).toMatchObject({
      status: 'failed',
      error: 'animated_image_unsupported',
    })
    expect(
      await readFile(join(userDataPath, 'presentation-attachments', hash('doc'), id, 'raw.webp')),
    ).toEqual(raw)
  })
  it('imports a static WebP URL and serves the converted asset after a restart', async () => {
    const userDataPath = await mkdtemp(join(tmpdir(), 'ppt-webp-url-'))
    dirs.push(userDataPath)
    const raw = fixture('static.webp')
    const service = createPresentationAttachmentService({
      userDataPath,
      fetchImage: async () => new Response(raw, { headers: { 'content-type': 'image/webp' } }),
      normalizeImage: async () => ({ bytes: fixture('control.png'), width: 2, height: 3 }),
    })
    const call = (body: Record<string, unknown>) =>
      service({ documentId: 'doc', ...body }, new AbortController().signal)
    const imported = (await call({
      operation: 'attachment_import_url',
      url: 'https://93.184.216.34/figure.webp',
    })) as { attachmentId: string }
    expect(imported).toMatchObject({
      status: 'ready',
      name: `remote-${imported.attachmentId}.webp`,
    })
    expect(
      await call({ operation: 'attachment_asset', attachmentId: imported.attachmentId }),
    ).toMatchObject({ mime: 'image/png', width: 2, height: 3 })
    const reopened = createPresentationAttachmentService({ userDataPath })
    expect(
      await reopened(
        { documentId: 'doc', operation: 'attachment_asset', attachmentId: imported.attachmentId },
        new AbortController().signal,
      ),
    ).toMatchObject({ mime: 'image/png', width: 2, height: 3 })
  })
  it('removes only old service staging directories on startup', async () => {
    const userDataPath = await mkdtemp(join(tmpdir(), 'ppt-staging-cleanup-'))
    dirs.push(userDataPath)
    const root = join(userDataPath, 'presentation-attachments')
    await mkdir(root)
    const old = join(root, '.tmp-00000000-0000-0000-0000-000000000001')
    const fresh = join(root, '.tmp-00000000-0000-0000-0000-000000000002')
    const unrelated = join(root, 'not-staging')
    await Promise.all([mkdir(old), mkdir(fresh), mkdir(unrelated)])
    await utimes(old, new Date(0), new Date(0))
    const service = createPresentationAttachmentService({ userDataPath })
    await service(
      { documentId: 'doc', operation: 'attachment_list_assets' },
      new AbortController().signal,
    )
    await expect(access(old)).rejects.toMatchObject({ code: 'ENOENT' })
    await expect(access(fresh)).resolves.toBeUndefined()
    await expect(access(unrelated)).resolves.toBeUndefined()
  })
  it('binds a user license declaration to a ready evidence attachment', async () => {
    const { call, userDataPath } = await setup()
    const imageId = await upload(call)
    await call({ operation: 'attachment_finish', attachmentId: imageId })
    const evidence = Buffer.from('Permission granted for this picture.', 'utf8')
    const evidenceId = await upload(call, evidence, 'permission.txt')
    await call({ operation: 'attachment_finish', attachmentId: evidenceId })
    const declared = (await call({
      operation: 'attachment_attest_license',
      attachmentId: imageId,
      license: 'licensed',
      evidenceAttachmentId: evidenceId,
    })) as { licenseDeclaration: { kind: string; evidenceAttachmentId: string } }
    expect(declared.licenseDeclaration).toMatchObject({
      kind: 'licensed',
      evidenceAttachmentId: evidenceId,
    })
    expect(await call({ operation: 'attachment_asset', attachmentId: imageId })).toMatchObject({
      license: 'licensed',
      licenseEvidence: `attachment:${evidenceId}`,
    })
    const evidencePath = join(
      userDataPath,
      'presentation-attachments',
      hash('doc'),
      evidenceId,
      'text.txt',
    )
    await writeFile(evidencePath, 'tampered')
    await expect(call({ operation: 'attachment_asset', attachmentId: imageId })).rejects.toThrow(
      'invalid_state',
    )
    await writeFile(evidencePath, evidence)
    await expect(
      call({ operation: 'attachment_delete', attachmentId: evidenceId }),
    ).rejects.toThrow('attachment_in_use')
    await call({ operation: 'attachment_revoke_license', attachmentId: imageId })
    expect(await call({ operation: 'attachment_asset', attachmentId: imageId })).not.toHaveProperty(
      'license',
    )
    expect(await call({ operation: 'attachment_delete', attachmentId: evidenceId })).toEqual({
      attachmentId: evidenceId,
      deleted: true,
    })
  })
  it('imports a public image URL once, redacts its query and reuses the PC cache', async () => {
    const userDataPath = await mkdtemp(join(tmpdir(), 'ppt-remote-image-'))
    dirs.push(userDataPath)
    const fetchImage = vi.fn(
      async () => new Response(png, { status: 200, headers: { 'content-type': 'image/png' } }),
    )
    const service = createPresentationAttachmentService({
      userDataPath,
      fetchImage,
      normalizeImage: async () => ({ bytes: png, width: 1, height: 1 }),
    })
    const call = (body: Record<string, unknown>) =>
      service({ documentId: 'doc', ...body }, new AbortController().signal)
    const url = 'https://93.184.216.34/image.png?token=secret'
    const first = (await call({ operation: 'attachment_import_url', url })) as {
      attachmentId: string
      source: string
    }
    expect(first).toMatchObject({ status: 'ready', source: 'https://93.184.216.34/image.png' })
    expect(JSON.stringify(first)).not.toContain('secret')
    expect(await call({ operation: 'attachment_import_url', url })).toEqual(first)
    expect(fetchImage).toHaveBeenCalledTimes(1)
    expect(
      await call({ operation: 'attachment_asset', attachmentId: first.attachmentId }),
    ).toMatchObject({ source: first.source })
    await expect(
      call({ operation: 'attachment_import_url', url: 'http://127.0.0.1/image.png' }),
    ).rejects.toThrow('remote_image_unavailable')
    expect(fetchImage).toHaveBeenCalledTimes(1)
    const path = join(
      userDataPath,
      'presentation-attachments',
      hash('doc'),
      first.attachmentId,
      'metadata.json',
    )
    const tampered = JSON.parse(await readFile(path, 'utf8'))
    tampered.source = 'file:///secret'
    await writeFile(path, JSON.stringify(tampered))
    await expect(call({ operation: 'attachment_list_assets' })).rejects.toThrow('invalid_state')
  })
  it('rejects an oversized remote response without publishing an attachment', async () => {
    const userDataPath = await mkdtemp(join(tmpdir(), 'ppt-remote-large-'))
    dirs.push(userDataPath)
    const service = createPresentationAttachmentService({
      userDataPath,
      fetchImage: async () => new Response(Buffer.alloc(10 * 1024 * 1024 + 1)),
      normalizeImage: async () => ({ bytes: png, width: 1, height: 1 }),
    })
    const call = (body: Record<string, unknown>) =>
      service({ documentId: 'doc', ...body }, new AbortController().signal)
    await expect(
      call({ operation: 'attachment_import_url', url: 'https://93.184.216.34/image.png' }),
    ).rejects.toThrow('quota_exceeded')
    expect(await call({ operation: 'attachment_list_assets' })).toEqual({ attachments: [] })
  })
  it('keeps both public sources when distinct URLs return the same image bytes', async () => {
    const userDataPath = await mkdtemp(join(tmpdir(), 'ppt-remote-sources-'))
    dirs.push(userDataPath)
    const fetchImage = vi.fn(async () => new Response(png))
    const service = createPresentationAttachmentService({
      userDataPath,
      fetchImage,
      normalizeImage: async () => ({ bytes: png, width: 1, height: 1 }),
    })
    const call = (url: string) =>
      service(
        { documentId: 'doc', operation: 'attachment_import_url', url },
        new AbortController().signal,
      )
    const first = (await call('https://93.184.216.34/a.png?private=one')) as {
      attachmentId: string
    }
    const second = (await call('https://93.184.216.34/b.png?private=two')) as {
      attachmentId: string
      sources: string[]
    }
    expect(second.attachmentId).toBe(first.attachmentId)
    expect(second.sources).toEqual(['https://93.184.216.34/a.png', 'https://93.184.216.34/b.png'])
    const asset = (await service(
      { documentId: 'doc', operation: 'attachment_asset', attachmentId: first.attachmentId },
      new AbortController().signal,
    )) as { sources: string[] }
    expect(asset.sources).toEqual(second.sources)
    expect(JSON.stringify(second)).not.toContain('private')
    expect(await call('https://93.184.216.34/b.png?private=two')).toEqual(second)
    expect(fetchImage).toHaveBeenCalledTimes(2)
  })
  it('allows more than 32 images while paging their bounded list', async () => {
    const { call } = await setup()
    for (let i = 0; i < 33; i++) {
      const id = hash(`image-${i}`)
      await call({
        operation: 'attachment_begin',
        attachmentId: id,
        sha256: id,
        name: `image-${i}.png`,
        sizeBytes: 1,
      })
    }
    const first = (await call({ operation: 'attachment_list_assets' })) as {
      attachments: { attachmentId: string }[]
      nextAfter?: string
    }
    expect(first.attachments).toHaveLength(32)
    expect(first.nextAfter).toBe(first.attachments.at(-1)?.attachmentId)
    const second = (await call({
      operation: 'attachment_list_assets',
      after: first.nextAfter,
    })) as { attachments: { attachmentId: string }[]; nextAfter?: string }
    expect(second.attachments).toHaveLength(1)
    expect(second.nextAfter).toBeUndefined()
    expect(second.attachments[0]!.attachmentId > first.nextAfter!).toBe(true)
    await expect(call({ operation: 'attachment_list_assets', after: 'bad' })).rejects.toThrow(
      'invalid_request',
    )
  })
  it('normalizes once, preserves original hash, and restores validated cache after restart', async () => {
    let calls = 0
    const { call, userDataPath } = await setup(async () => {
      calls++
      return { bytes: png, width: 1, height: 1 }
    })
    const id = await upload(call)
    expect(await call({ operation: 'attachment_finish', attachmentId: id })).toEqual({
      attachmentId: id,
      name: 'photo.png',
      sizeBytes: png.length,
      sha256: id,
      receivedBytes: png.length,
      status: 'ready',
      kind: 'image',
      mime: 'image/png',
      width: 1,
      height: 1,
      assetSha256: hash(png),
    })
    const service = createPresentationAttachmentService({
      userDataPath,
      normalizeImage: async () => {
        throw new Error('must not decode twice')
      },
    })
    expect(
      await service(
        { operation: 'attachment_finish', documentId: 'doc', attachmentId: id },
        new AbortController().signal,
      ),
    ).toMatchObject({ kind: 'image' })
    expect(calls).toBe(1)
    expect(await call({ operation: 'attachment_asset', attachmentId: id })).toEqual({
      id,
      mime: 'image/png',
      base64: png.toString('base64'),
      width: 1,
      height: 1,
      source: `attachment:${id}`,
    })
    expect(await call({ operation: 'attachment_list_assets' })).toMatchObject({
      attachments: [{ kind: 'image', width: 1 }],
    })
    await expect(
      call({ operation: 'attachment_asset', documentId: 'different', attachmentId: id }),
    ).rejects.toThrow('not_found')
    await expect(
      call({ operation: 'attachment_read', attachmentId: id, offset: 0, maxChars: 10 }),
    ).rejects.toThrow('invalid_state')
  })
  it('preserves the legacy text list while the new list includes images in every status', async () => {
    const { call } = await setup()
    const text = await upload(call, Buffer.from('evidence'), 'source.txt')
    await call({ operation: 'attachment_finish', attachmentId: text })
    const image = await upload(call)
    expect(await call({ operation: 'attachment_list' })).toMatchObject({
      attachments: [{ attachmentId: text, kind: 'text' }],
    })
    const pending = (await call({ operation: 'attachment_list_assets' })) as {
      attachments: unknown[]
    }
    expect(pending.attachments).toHaveLength(2)
    await call({ operation: 'attachment_finish', attachmentId: image })
    expect(await call({ operation: 'attachment_list' })).toMatchObject({
      attachments: [{ attachmentId: text, kind: 'text' }],
    })
    const ready = (await call({ operation: 'attachment_list_assets' })) as {
      attachments: unknown[]
    }
    expect(ready.attachments).toHaveLength(2)
  })
  it('keeps original JPEG identity when its normalized PNG has different bytes', async () => {
    const { call } = await setup()
    const jpeg = Buffer.from([255, 216, 255, 192, 0, 11, 8, 0, 1, 0, 1, 1, 1, 17, 0, 255, 217])
    const id = await upload(call, jpeg, 'photo.jpeg')
    expect(await call({ operation: 'attachment_finish', attachmentId: id })).toMatchObject({
      sha256: hash(jpeg),
      assetSha256: hash(png),
      mime: 'image/png',
    })
    expect(await call({ operation: 'attachment_asset', attachmentId: id })).toMatchObject({
      id: hash(jpeg),
      base64: png.toString('base64'),
    })
  })
  it('rejects spoofed MIME and pixel-limit headers before invoking normalization', async () => {
    let calls = 0
    const { call } = await setup(async () => {
      calls++
      return { bytes: png, width: 1, height: 1 }
    })
    const id = await upload(call, png, 'spoof.jpg')
    expect(await call({ operation: 'attachment_finish', attachmentId: id })).toMatchObject({
      status: 'failed',
      error: 'parse_failed',
    })
    const large = Buffer.from(png)
    large.writeUInt32BE(8000, 16)
    large.writeUInt32BE(8000, 20)
    const largeId = await upload(call, large, 'large.png')
    expect(await call({ operation: 'attachment_finish', attachmentId: largeId })).toMatchObject({
      status: 'failed',
    })
    expect(calls).toBe(0)
    await expect(
      call({
        operation: 'attachment_begin',
        attachmentId: id,
        sha256: id,
        name: 'huge.png',
        sizeBytes: 10 * 1024 * 1024 + 1,
      }),
    ).rejects.toThrow('invalid_request')
  })
  it('denies tampered or symlinked normalized cache and does not repair silently on repeated finish', async () => {
    const { call, userDataPath } = await setup()
    const id = await upload(call)
    await call({ operation: 'attachment_finish', attachmentId: id })
    const path = join(userDataPath, 'presentation-attachments', hash('doc'), id, 'image.png')
    const bad = Buffer.from(png)
    bad[45] = bad[45]! ^ 1
    await writeFile(path, bad)
    await expect(call({ operation: 'attachment_asset', attachmentId: id })).rejects.toThrow(
      'invalid_state',
    )
    await expect(call({ operation: 'attachment_finish', attachmentId: id })).rejects.toThrow(
      'invalid_state',
    )
    await rm(path)
    await symlink('/etc/passwd', path)
    await expect(call({ operation: 'attachment_asset', attachmentId: id })).rejects.toThrow(
      'invalid_state',
    )
  })
  it('cancels normalization without publishing ready and retries safely', async () => {
    const controller = new AbortController()
    let first = true
    const { call } = await setup(async () => {
      if (first) {
        first = false
        controller.abort()
      }
      return { bytes: png, width: 1, height: 1 }
    })
    const id = await upload(call)
    await expect(
      call({ operation: 'attachment_finish', attachmentId: id }, controller.signal),
    ).rejects.toThrow('aborted')
    expect(await call({ operation: 'attachment_list_assets' })).toMatchObject({
      attachments: [{ status: 'uploading' }],
    })
    expect(await call({ operation: 'attachment_finish', attachmentId: id })).toMatchObject({
      status: 'ready',
    })
  })
  it('rejects oversized or inconsistent normalizer output', async () => {
    for (const output of [
      { bytes: png, width: 2, height: 1 },
      { bytes: Buffer.alloc(4 * 1024 * 1024 + 1), width: 1, height: 1 },
    ]) {
      const { call } = await setup(async () => output)
      const id = await upload(call)
      expect(await call({ operation: 'attachment_finish', attachmentId: id })).toMatchObject({
        status: 'failed',
      })
    }
  })
  it('does not expose text attachments through the image route', async () => {
    const { call } = await setup()
    const id = await upload(call, Buffer.from('text'), 'source.txt')
    await call({ operation: 'attachment_finish', attachmentId: id })
    await expect(call({ operation: 'attachment_asset', attachmentId: id })).rejects.toThrow(
      'invalid_state',
    )
  })
})
