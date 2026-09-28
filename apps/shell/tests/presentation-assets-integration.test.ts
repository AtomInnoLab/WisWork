import { createHash } from 'node:crypto'
import { mkdtempSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { PresentationStore } from '@wiswork/project-store'
import { afterEach, expect, it, vi } from 'vitest'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { createPresentationService } from '../src/main/presentation-service'
import { createPresentationAttachmentService } from '../src/main/presentation-attachments'

const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const signal = () => new AbortController().signal
const decode = (value: Uint8Array) => JSON.parse(Buffer.from(value).toString('utf8'))

it('rejects inline image evidence that was not resolved against this document', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'presentation-evidence-forgery-'))
  roots.push(userDataPath)
  const deck = benchmarkDeck()
  deck.assets[0]!.license = 'licensed'
  deck.assets[0]!.licenseEvidence = `attachment:${'a'.repeat(64)}`
  const service = createPresentationService({ userDataPath })
  const result = decode(
    await service(
      { operation: 'compile', documentId: 'document-forgery', requestId: 'forged', deck },
      signal(),
    ),
  )
  expect(result).toEqual({ error: 'invalid_deck' })
})

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
  const evidence = Buffer.from('Licensed for this presentation.', 'utf8')
  const evidenceId = createHash('sha256').update(evidence).digest('hex')
  await request({
    operation: 'attachment_begin',
    attachmentId: evidenceId,
    name: 'permission.txt',
    sha256: evidenceId,
    sizeBytes: evidence.length,
  })
  await request({
    operation: 'attachment_chunk',
    attachmentId: evidenceId,
    offset: 0,
    base64: evidence.toString('base64'),
  })
  await request({ operation: 'attachment_finish', attachmentId: evidenceId })
  await request({
    operation: 'attachment_attest_license',
    attachmentId,
    license: 'licensed',
    evidenceAttachmentId: evidenceId,
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
  expect(await zip.file('ppt/notesSlides/notesSlide3.xml')!.async('string')).toContain(
    `attachment:${evidenceId}`,
  )
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

it('compiles an explicitly selected animated first frame into PPTX image media', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'presentation-animation-asset-'))
  roots.push(userDataPath)
  const raw = readFileSync(join(__dirname, 'fixtures/presentation-image/animated.gif'))
  const firstFrame = readFileSync(join(__dirname, 'fixtures/presentation-image/control.png'))
  const attachmentId = createHash('sha256').update(raw).digest('hex')
  const service = createPresentationService({
    userDataPath,
    normalizeFirstFrame: async () => ({ bytes: firstFrame, width: 2, height: 3 }),
  })
  const request = async (body: Record<string, unknown>) =>
    decode(await service({ documentId: 'document-animation', ...body }, signal()))
  await request({
    operation: 'attachment_begin',
    attachmentId,
    name: 'animation.gif',
    sha256: attachmentId,
    sizeBytes: raw.length,
  })
  await request({
    operation: 'attachment_chunk',
    attachmentId,
    offset: 0,
    base64: raw.toString('base64'),
  })
  expect(await request({ operation: 'attachment_finish', attachmentId })).toMatchObject({
    status: 'failed',
    error: 'animated_image_unsupported',
  })
  expect(
    await request({ operation: 'attachment_extract_first_frame', attachmentId }),
  ).toMatchObject({ status: 'ready', animationHandling: 'first_frame' })
  const deck = benchmarkDeck()
  deck.assets = [{ id: deck.assets[0]!.id, attachmentId }]
  const output = await request({ operation: 'compile', requestId: 'first-frame', deck })
  expect(output.status).toBe('compiled')
  const zip = await JSZip.loadAsync(Buffer.from(output.pptxBase64, 'base64'))
  const media = Object.keys(zip.files).filter(
    (path) => path.startsWith('ppt/media/') && !zip.files[path]!.dir,
  )
  expect(media.length).toBeGreaterThan(0)
  expect(await zip.file(media[0]!)!.async('nodebuffer')).toEqual(firstFrame)
})

it('keeps an animated URL unusable until explicit conversion, then compiles its source and first frame', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'presentation-animation-url-'))
  roots.push(userDataPath)
  const raw = readFileSync(join(__dirname, 'fixtures/presentation-image/animated.webp'))
  const firstFrame = readFileSync(join(__dirname, 'fixtures/presentation-image/control.png'))
  const service = createPresentationService({
    userDataPath,
    fetchImage: async () => new Response(raw),
    normalizeFirstFrame: async () => ({ bytes: firstFrame, width: 2, height: 3 }),
  })
  const request = async (body: Record<string, unknown>) =>
    decode(await service({ documentId: 'document-animation-url', ...body }, signal()))
  const imported = await request({
    operation: 'attachment_import_url',
    url: 'https://93.184.216.34/animated.webp?private=secret',
    stageAnimated: true,
  })
  expect(imported).toMatchObject({ status: 'failed', error: 'animated_image_unsupported' })
  const deck = benchmarkDeck()
  deck.assets = [{ id: deck.assets[0]!.id, attachmentId: imported.attachmentId }]
  expect(
    await request({ operation: 'compile', requestId: 'url-before-choice', deck }),
  ).toMatchObject({ error: 'asset_unavailable' })
  expect(
    await request({
      operation: 'attachment_extract_first_frame',
      attachmentId: imported.attachmentId,
    }),
  ).toMatchObject({ status: 'ready', animationHandling: 'first_frame' })
  const output = await request({ operation: 'compile', requestId: 'url-after-choice', deck })
  expect(output.status).toBe('compiled')
  const zip = await JSZip.loadAsync(Buffer.from(output.pptxBase64, 'base64'))
  const media = Object.keys(zip.files).filter(
    (path) => path.startsWith('ppt/media/') && !zip.files[path]!.dir,
  )
  expect(await zip.file(media[0]!)!.async('nodebuffer')).toEqual(firstFrame)
  const notes = await zip.file('ppt/notesSlides/notesSlide3.xml')!.async('string')
  expect(notes).toContain('https://93.184.216.34/animated.webp')
  expect(notes).not.toContain('private=secret')
})

it('keeps all PC image URL sources in compiled PowerPoint notes', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'presentation-image-sources-'))
  roots.push(userDataPath)
  const deck = benchmarkDeck()
  const image = Buffer.from(deck.assets[0]!.base64, 'base64')
  const attachments = createPresentationAttachmentService({
    userDataPath,
    fetchImage: async () => new Response(image),
    normalizeImage: async () => ({ bytes: image, width: 1, height: 1 }),
  })
  const imported = (await attachments(
    {
      documentId: 'document-images',
      operation: 'attachment_import_url',
      url: 'https://93.184.216.34/first.png',
    },
    signal(),
  )) as { attachmentId: string }
  await attachments(
    {
      documentId: 'document-images',
      operation: 'attachment_import_url',
      url: 'https://93.184.216.34/second.png',
    },
    signal(),
  )
  const service = createPresentationService({ userDataPath })
  const output = decode(
    await service(
      {
        documentId: 'document-images',
        operation: 'compile',
        requestId: 'all-sources',
        deck: {
          ...deck,
          assets: [{ id: deck.assets[0]!.id, attachmentId: imported.attachmentId }],
        },
      },
      signal(),
    ),
  )
  expect(output.status).toBe('compiled')
  const zip = await JSZip.loadAsync(Buffer.from(output.pptxBase64, 'base64'))
  const notes = await zip.file('ppt/notesSlides/notesSlide3.xml')!.async('string')
  expect(notes).toContain('https://93.184.216.34/first.png')
  expect(notes).toContain('https://93.184.216.34/second.png')
  expect(output.report.checks.sources).toBe('not_verified')
})
