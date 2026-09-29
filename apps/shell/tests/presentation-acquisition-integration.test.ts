import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { afterEach, expect, it, vi } from 'vitest'
import { createPresentationService } from '../src/main/presentation-service'
import { createPresentationAttachmentSkill } from '../../office-addin/src/skills/powerpoint/presentation-attachments'
import { InMemoryVfs } from '../../office-addin/src/skills/shared/vfs'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const png = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII=',
  'base64',
)
const hash = (value: string | Uint8Array) => createHash('sha256').update(value).digest('hex')
function setup(
  fetchImage: (url: string, signal: AbortSignal) => Promise<Response | null>,
  fetchPage?: (url: string, signal: AbortSignal) => Promise<Response | null>,
) {
  const userDataPath = mkdtempSync(join(tmpdir(), 'ppt-acquisition-cross-'))
  roots.push(userDataPath)
  let documentId = 'document'
  let pc = createPresentationService({
    userDataPath,
    fetchImage,
    fetchPage,
    normalizeImage: async () => ({ bytes: png, width: 1, height: 1 }),
  })
  const request = vi.fn(
    async (body: unknown, signal?: AbortSignal) =>
      new Response(
        Buffer.from(await pc(body, signal ?? new AbortController().signal)).toString('utf8'),
      ),
  )
  const client = () =>
    createPresentationAttachmentSkill({
      available: () => true,
      remoteImagesAvailable: () => true,
      webpagesAvailable: () => true,
      imagesAvailable: () => true,
      documentId: async () => documentId,
      request,
      vfs: new InMemoryVfs(),
    })
  return {
    client,
    request,
    switchDocument: (value: string) => {
      documentId = value
    },
    restart: () => {
      pc = createPresentationService({
        userDataPath,
        fetchImage,
        fetchPage,
        normalizeImage: async () => ({ bytes: png, width: 1, height: 1 }),
      })
    },
  }
}
it('restores actual image fallback and cache attempts across both PC and Taskpane reopening', async () => {
  const primary = 'https://8.8.8.8/primary.png?private=secret#fragment'
  const backup = 'https://8.8.8.8/backup.png?private=secret'
  const fetchImage = vi.fn(async (url: string) =>
    url === primary ? null : new Response(png, { headers: { 'content-type': 'image/png' } }),
  )
  const f = setup(fetchImage)
  const client = f.client()
  const result = await client.importUrls([primary, backup])
  expect(result.attachmentId).toBe(hash(png))
  const before = await client.acquisitionHistory()
  expect(before).toMatchObject({ documentId: 'document', totalAttempts: 2 })
  expect(before!.records.map((record) => [record.attempt, record.state])).toEqual([
    [1, 'rejected'],
    [2, 'ready'],
  ])
  expect(before!.records[0]).toMatchObject({
    source: 'https://8.8.8.8/primary.png',
    sourceUrlHash: hash(new URL(primary).toString()),
    error: 'remote_image_unavailable',
  })
  expect(JSON.stringify(before)).not.toContain('secret')
  f.restart()
  const reopened = f.client()
  expect(await reopened.acquisitionHistory()).toEqual(before)
  await reopened.importUrl(backup)
  expect(fetchImage).toHaveBeenCalledTimes(2)
  const after = await reopened.acquisitionHistory()
  expect(after!.records.slice(0, 2)).toEqual(before!.records)
  expect(after!.records.at(-1)).toMatchObject({
    attempt: 3,
    state: 'ready',
    attachmentId: hash(png),
    assetSha256: hash(png),
  })
  f.switchDocument('another-document')
  expect(await f.client().acquisitionHistory()).toMatchObject({
    documentId: 'another-document',
    totalAttempts: 0,
    records: [],
  })
})
it('preserves the real webpage failure and frozen snapshot identity on a later explicit retry', async () => {
  const raw = '<html><p>原始研究资料与限定语。</p></html>'
  const fetchPage = vi
    .fn()
    .mockResolvedValueOnce(null)
    .mockImplementation(async () => new Response(raw, { headers: { 'content-type': 'text/html' } }))
  const f = setup(async () => null, fetchPage)
  const url = 'https://8.8.8.8/research?private=secret'
  await expect(f.client().importWebpage(url)).rejects.toThrow(
    'presentation_remote_webpage_unavailable',
  )
  const result = await f.client().importWebpage(url)
  f.restart()
  const restored = await f.client().acquisitionHistory()
  expect(restored!.records.map((record) => record.state)).toEqual(['rejected', 'ready'])
  expect(restored!.records[1]).toMatchObject({
    kind: 'webpage',
    attachmentId: result.attachmentId,
    sha256: hash(raw),
    sizeBytes: Buffer.byteLength(raw),
  })
  expect(restored!.records[1]).not.toHaveProperty('assetSha256')
  expect(JSON.stringify(restored)).not.toContain('secret')
})
it('recovers a saved outcome after the import response is lost without repeating its network request', async () => {
  const fetchImage = vi.fn(async () => new Response(png))
  const f = setup(fetchImage)
  const actualRequest = f.request.getMockImplementation()!
  f.request.mockImplementation(async (body, signal) => {
    const response = await actualRequest(body, signal)
    if ((body as { operation: string }).operation === 'attachment_import_url')
      throw new Error('lost-response')
    return response
  })
  await expect(f.client().importUrl('https://8.8.8.8/photo.png')).rejects.toThrow('lost-response')
  const history = await f.client().acquisitionHistory()
  expect(history!.records).toHaveLength(1)
  expect(history!.records[0]).toMatchObject({ state: 'ready', attachmentId: hash(png) })
  expect(fetchImage).toHaveBeenCalledTimes(1)
  expect(
    f.request.mock.calls.filter(
      ([body]) => (body as { operation: string }).operation === 'attachment_import_url',
    ),
  ).toHaveLength(1)
})
it('reads the persisted start while network work is pending and does not invent a terminal outcome', async () => {
  let release!: (value: Response) => void
  let started!: () => void
  const reached = new Promise<void>((resolve) => {
    started = resolve
  })
  const fetchImage = vi.fn(async () => {
    started()
    return new Promise<Response>((resolve) => {
      release = resolve
    })
  })
  const f = setup(fetchImage)
  const client = f.client()
  const pending = client.importUrl('https://8.8.8.8/photo.png')
  await reached
  try {
    const history = await client.acquisitionHistory()
    expect(history!.records[0]).toMatchObject({ state: 'fetching', attempt: 1 })
    expect(history!.records[0]).not.toHaveProperty('finishedAt')
  } finally {
    release(new Response(png))
    await pending
  }
  expect((await client.acquisitionHistory())!.records[0]).toMatchObject({
    state: 'ready',
    attempt: 1,
  })
})
it('preserves an interrupted fetch after clearing the pane without replaying the operation', async () => {
  let started!: () => void
  const reached = new Promise<void>((resolve) => {
    started = resolve
  })
  const fetchImage = vi.fn(async (_url: string, signal: AbortSignal) => {
    started()
    return new Promise<Response>((_resolve, reject) =>
      signal.addEventListener('abort', () => reject(new Error('private download error')), {
        once: true,
      }),
    )
  })
  const f = setup(fetchImage)
  const client = f.client()
  const pending = client.importUrl('https://8.8.8.8/photo.png?private=secret')
  await reached
  client.clear()
  await expect(pending).rejects.toThrow('upload_cancelled')
  f.restart()
  const history = await f.client().acquisitionHistory()
  expect(history!.records[0]).toMatchObject({ state: 'rejected', error: 'aborted', attempt: 1 })
  expect(JSON.stringify(history)).not.toContain('private download error')
  expect(JSON.stringify(history)).not.toContain('secret')
  expect(fetchImage).toHaveBeenCalledTimes(1)
})
