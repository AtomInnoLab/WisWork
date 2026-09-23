import { createHash } from 'node:crypto'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { PresentationStore } from '@wiswork/project-store'
import { afterEach, expect, it, vi } from 'vitest'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { createPresentationService } from '../src/main/presentation-service'

const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const signal = () => new AbortController().signal
const decode = (value: Uint8Array) => JSON.parse(Buffer.from(value).toString('utf8'))

it('compiles durable image references into real PPTX media and recovers without decoding again', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'presentation-assets-'))
  roots.push(userDataPath)
  const originalDeck = benchmarkDeck()
  const original = originalDeck.assets[0]!
  if (!('base64' in original)) throw new Error('fixture must be inline')
  const bytes = Buffer.from(original.base64, 'base64')
  const attachmentId = createHash('sha256').update(bytes).digest('hex')
  const normalizeImage = vi.fn(async () => ({ bytes, width: 1, height: 1 }))
  let service = createPresentationService({ userDataPath, normalizeImage })
  const request = async (body: Record<string, unknown>) =>
    decode(await service({ documentId: 'document-images', ...body }, signal()))
  await request({
    operation: 'attachment_begin',
    attachmentId,
    name: 'chart.png',
    sha256: attachmentId,
    sizeBytes: bytes.length,
  })
  await request({
    operation: 'attachment_chunk',
    attachmentId,
    offset: 0,
    base64: bytes.toString('base64'),
  })
  expect(await request({ operation: 'attachment_finish', attachmentId })).toMatchObject({
    status: 'ready',
    kind: 'image',
    width: 1,
    height: 1,
  })
  const deck = { ...originalDeck, assets: [{ id: original.id, attachmentId }] }
  const input = { operation: 'compile', requestId: 'images-first', deck }
  expect(Buffer.byteLength(JSON.stringify(input))).toBeLessThan(256 * 1024)
  const output = await request(input)
  expect(output.status).toBe('compiled')
  const zip = await JSZip.loadAsync(Buffer.from(output.pptxBase64, 'base64'))
  const media = Object.keys(zip.files).filter(
    (path) => path.startsWith('ppt/media/') && !zip.files[path]!.dir,
  )
  expect(media.length).toBeGreaterThan(0)
  expect(await zip.file(media[0]!)!.async('nodebuffer')).toEqual(bytes)
  expect(output.report.checks).toMatchObject({ render: 'not_run', sources: 'not_verified' })
  service = createPresentationService({ userDataPath, normalizeImage })
  expect(await request(input)).toEqual(output)
  expect(await request({ operation: 'attachment_finish', attachmentId })).toMatchObject({
    status: 'ready',
  })
  expect(await request({ ...input, requestId: 'images-second' })).toMatchObject({
    status: 'compiled',
  })
  const saved = new PresentationStore(userDataPath).request(
    deck.id,
    'document-images',
    'images-second',
  )!
  expect((saved.deck as typeof deck).assets).toEqual([{ id: original.id, attachmentId }])
  expect(normalizeImage).toHaveBeenCalledTimes(1)
  expect(
    await request({
      ...input,
      documentId: 'other-document',
      requestId: 'other',
      deck: { ...deck, id: 'another-project' },
    }),
  ).toEqual({ error: 'not_found' })
})
