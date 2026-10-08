import { expect, it, vi } from 'vitest'
import { createHash } from 'node:crypto'
import { createPresentationImageBackup } from '../src/skills/powerpoint/presentation-image-backup'
const bytes = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII=',
  'base64',
)
const metadata = {
  attachmentId: createHash('sha256').update(bytes).digest('hex'),
  sizeBytes: bytes.length,
  mime: 'image/png' as const,
}
const good = {
  ...metadata,
  sha256: metadata.attachmentId,
  offset: 0,
  base64: bytes.toString('base64'),
}
it('fails closed for malformed, truncated, wrong identity and corrupted original responses', async () => {
  for (const response of [
    null,
    {},
    { ...good, offset: 1 },
    { ...good, mime: 'image/jpeg' },
    { ...good, sha256: 'a'.repeat(64) },
    { ...good, sizeBytes: bytes.length + 1 },
    { ...good, base64: bytes.subarray(1).toString('base64') },
    { ...good, base64: Buffer.alloc(bytes.length).toString('base64') },
    { ...good, extra: true },
    { error: 'unsupported_operation' },
  ]) {
    const client = createPresentationImageBackup({
      available: () => true,
      request: async () => new Response(JSON.stringify(response)),
    })
    await expect(client.load('doc', metadata)).rejects.toThrow()
  }
})
it('rejects invalid source bytes, oversize bytes and cancellation before contacting PC', async () => {
  const request = vi.fn(async () => new Response('{}'))
  const client = createPresentationImageBackup({ available: () => true, request })
  for (const base64 of [
    '',
    '!!!',
    bytes.toString('base64') + '\n',
    Buffer.alloc(2 * 1024 * 1024 + 1).toString('base64'),
    btoa('text'),
  ])
    await expect(client.save('doc', base64)).rejects.toThrow()
  const c = new AbortController()
  c.abort()
  await expect(client.save('doc', bytes.toString('base64'), c.signal)).rejects.toThrow('cancelled')
  expect(request).not.toHaveBeenCalled()
})
it('cancellation or availability loss during response prevents further processing', async () => {
  for (const abort of [true, false]) {
    const c = new AbortController()
    let available = true
    const client = createPresentationImageBackup({
      available: () => available,
      request: async () => {
        if (abort) c.abort()
        else available = false
        return new Response(JSON.stringify(good))
      },
    })
    await expect(client.load('doc', metadata, c.signal)).rejects.toThrow(
      abort ? 'cancelled' : 'unavailable',
    )
  }
})
it('reads multiple bounded chunks and checks complete digest', async () => {
  const large = Buffer.concat([bytes, Buffer.alloc(300000)])
  const m = {
    ...metadata,
    sizeBytes: large.length,
    attachmentId: createHash('sha256').update(large).digest('hex'),
  }
  const request = vi.fn(async (body: unknown) => {
    const b = body as { offset: number; length: number; documentId: string }
    expect(b.documentId).toBe('doc')
    expect(b.length).toBeLessThanOrEqual(128 * 1024)
    return new Response(
      JSON.stringify({
        ...m,
        sha256: m.attachmentId,
        offset: b.offset,
        base64: large.subarray(b.offset, b.offset + b.length).toString('base64'),
      }),
    )
  })
  expect(
    await createPresentationImageBackup({ available: () => true, request }).load('doc', m),
  ).toBe(large.toString('base64'))
  expect(request).toHaveBeenCalledTimes(3)
})
it('rejects older service responses before returning a usable savepoint', async () => {
  const request = vi.fn(async () => new Response(JSON.stringify({ error: 'invalid_request' })))
  await expect(
    createPresentationImageBackup({ available: () => true, request }).save(
      'doc',
      bytes.toString('base64'),
    ),
  ).rejects.toThrow('unavailable')
  expect(request).toHaveBeenCalledTimes(1)
})
