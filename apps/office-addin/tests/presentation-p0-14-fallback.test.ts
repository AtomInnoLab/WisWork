import { readFileSync } from 'node:fs'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { createPresentationAttachmentSkill } from '../src/skills/powerpoint/presentation-attachments.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'
import { createPresentationAttachmentService } from '../../shell/src/main/presentation-attachments.js'
import { createPresentationService } from '../../shell/src/main/presentation-service.js'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark.js'
import type { PresentationDeck } from '../../../packages/pptx-engine/src/presentation.js'

const roots: string[] = []
afterEach(async () => {
  await Promise.all(roots.splice(0).map((path) => rm(path, { recursive: true, force: true })))
})

it('falls back after an image fetch failure, reuses the durable cache, and compiles eight pages', async () => {
  const userDataPath = await mkdtemp(join(tmpdir(), 'ppt-p0-14-fallback-'))
  roots.push(userDataPath)
  const image = readFileSync(
    new URL(
      '../../../docs/product/ppt-benchmark-materials/PPT-P0-13/images/schematic-07.png',
      import.meta.url,
    ),
  )
  const failed = 'https://93.184.216.34/controlled-timeout.png'
  const healthy = 'https://93.184.216.34/owned-illustration.png'
  const fetchImage = vi.fn(async (url: string) => {
    if (url === failed) throw new Error('simulated_timeout')
    if (url === healthy) return new Response(image, { headers: { 'content-type': 'image/png' } })
    throw new Error('unexpected_source')
  })
  const attachments = createPresentationAttachmentService({
    userDataPath,
    fetchImage,
    normalizeImage: async () => ({ bytes: image, width: 960, height: 540 }),
  })
  const skill = createPresentationAttachmentSkill({
    available: () => true,
    remoteImagesAvailable: () => true,
    request: async (body, signal) => {
      try {
        const result = await attachments(
          body as Record<string, unknown>,
          signal ?? new AbortController().signal,
        )
        return new Response(JSON.stringify(result))
      } catch (error) {
        return new Response(JSON.stringify({ error: (error as Error).message }))
      }
    },
    documentId: async () => 'p0-14-document',
    vfs: new InMemoryVfs(),
  })
  const imported = await skill.importUrls([failed, healthy])
  expect(imported).toMatchObject({ status: 'ready', kind: 'image', width: 960, height: 540 })
  expect(imported.source).toBe(healthy)
  expect(fetchImage.mock.calls.map(([url]) => url)).toEqual([failed, healthy])
  expect((await skill.importUrls([healthy])).attachmentId).toBe(imported.attachmentId)
  expect(fetchImage).toHaveBeenCalledTimes(2)

  const deck: PresentationDeck = benchmarkDeck()
  deck.assets = [{ id: deck.assets[0]!.id, attachmentId: imported.attachmentId }]
  const service = createPresentationService({ userDataPath })
  const compiled = JSON.parse(
    Buffer.from(
      await service(
        { operation: 'compile', documentId: 'p0-14-document', requestId: 'fallback', deck },
        new AbortController().signal,
      ),
    ).toString('utf8'),
  )
  expect(compiled.status).toBe('compiled')
  const zip = await JSZip.loadAsync(Buffer.from(compiled.pptxBase64, 'base64'))
  expect(
    Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path)),
  ).toHaveLength(8)
  const media = Object.keys(zip.files).filter(
    (path) => path.startsWith('ppt/media/') && !zip.files[path]!.dir,
  )
  expect(media).toHaveLength(1)
  expect(await zip.file(media[0]!)!.async('nodebuffer')).toEqual(image)
})

it('reports exhausted sources without leaving a ready image attachment', async () => {
  const userDataPath = await mkdtemp(join(tmpdir(), 'ppt-p0-14-exhausted-'))
  roots.push(userDataPath)
  const fetchImage = vi.fn(async () => {
    throw new Error('simulated_timeout')
  })
  const attachments = createPresentationAttachmentService({ userDataPath, fetchImage })
  const skill = createPresentationAttachmentSkill({
    available: () => true,
    remoteImagesAvailable: () => true,
    request: async (body, signal) => {
      try {
        return new Response(
          JSON.stringify(
            await attachments(
              body as Record<string, unknown>,
              signal ?? new AbortController().signal,
            ),
          ),
        )
      } catch (error) {
        return new Response(JSON.stringify({ error: (error as Error).message }))
      }
    },
    documentId: async () => 'p0-14-document',
    vfs: new InMemoryVfs(),
  })
  await expect(
    skill.importUrls(['https://93.184.216.34/failed-a.png', 'https://93.184.216.34/failed-b.png']),
  ).rejects.toThrow('presentation_image_candidates_exhausted')
  expect(fetchImage).toHaveBeenCalledTimes(2)
  const listed = (await attachments(
    { operation: 'attachment_list_assets', documentId: 'p0-14-document' },
    new AbortController().signal,
  )) as { attachments: unknown[] }
  expect(listed.attachments).toEqual([])
})
