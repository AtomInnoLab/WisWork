import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, expect, it } from 'vitest'
import JSZip from 'jszip'
import { readFileSync } from 'node:fs'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkPlannedDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { replacePowerPointPictureMediaPackage } from '../../office-addin/src/skills/powerpoint/presentation-picture-package'
import { createPresentationService } from '../src/main/presentation-service'

const roots: string[] = []
afterEach(() => roots.splice(0).forEach((root) => rmSync(root, { recursive: true, force: true })))
const sha = (value: string | Buffer) => createHash('sha256').update(value).digest('hex')
async function fixture() {
  const userDataPath = mkdtempSync(join(tmpdir(), 'wiswork-existing-page-'))
  roots.push(userDataPath)
  const deck = benchmarkPlannedDeck()
  const raw = Buffer.from(
    (await compilePresentationDeck({ ...deck, slides: [deck.slides[0]!] })).bytes,
  )
  const begin = {
    backupId: 'native-backup-1',
    hostSlideId: 'host-1',
    slideIds: ['host-1', 'host-2'],
    sha256: sha(raw),
    sizeBytes: raw.length,
  }
  const call = async (
    operation: string,
    extra: Record<string, unknown> = {},
    documentId = 'document-1',
    service = createPresentationService({ userDataPath }),
  ) =>
    JSON.parse(
      Buffer.from(
        await service({ operation, documentId, ...extra }, new AbortController().signal),
      ).toString(),
    ) as Record<string, unknown>
  return { userDataPath, raw, begin, call }
}
it('persists exact single-page package, resumes chunks and reads after restart', async () => {
  const f = await fixture()
  expect(await f.call('existing_page_backup_begin', f.begin)).toMatchObject({
    ...f.begin,
    documentId: 'document-1',
    status: 'uploading',
    receivedBytes: 0,
  })
  const half = Math.floor(f.raw.length / 2)
  await f.call('existing_page_backup_chunk', {
    backupId: f.begin.backupId,
    offset: 0,
    base64: f.raw.subarray(0, half).toString('base64'),
  })
  expect(await f.call('existing_page_backup_status', { backupId: f.begin.backupId })).toMatchObject(
    {
      receivedBytes: half,
      status: 'uploading',
    },
  )
  await f.call('existing_page_backup_chunk', {
    backupId: f.begin.backupId,
    offset: half,
    base64: f.raw.subarray(half).toString('base64'),
  })
  expect(await f.call('existing_page_backup_finish', { backupId: f.begin.backupId })).toMatchObject(
    {
      status: 'ready',
      sha256: f.begin.sha256,
      receivedBytes: f.raw.length,
    },
  )
  expect(await f.call('existing_page_backup_begin', f.begin)).toMatchObject({ status: 'ready' })
  const read = await f.call('existing_page_backup_read', {
    backupId: f.begin.backupId,
    offset: 0,
    length: 131072,
  })
  expect(Buffer.from(read.base64 as string, 'base64')).toEqual(f.raw)
})

it('keeps frozen P0-13 original and picture revision as distinct durable PC savepoints', async () => {
  const material = new URL(
    '../../../docs/product/ppt-benchmark-materials/PPT-P0-13/',
    import.meta.url,
  )
  const zip = await JSZip.loadAsync(
    readFileSync(new URL('wiswork-image-dense-research-draft.pptx', material)),
  )
  for (let page = 1; page <= 8; page++)
    if (page !== 4) {
      zip.remove(`ppt/slides/slide${page}.xml`)
      zip.remove(`ppt/slides/_rels/slide${page}.xml.rels`)
    }
  zip.file(
    'ppt/presentation.xml',
    (await zip.file('ppt/presentation.xml')!.async('string')).replace(
      /<p:sldId\b[^>]*\/>/g,
      (item) => (item.includes('r:id="rId5"') ? item : ''),
    ),
  )
  const original = Buffer.from(await zip.generateAsync({ type: 'uint8array' }))
  const slide = await zip.file('ppt/slides/slide4.xml')!.async('string')
  const pictureId = /<p:pic\b[\s\S]*?<p:cNvPr id="(\d+)"/.exec(slide)![1]
  const replacement = readFileSync(new URL('images/schematic-12.png', material))
  const revised = Buffer.from(
    (
      await replacePowerPointPictureMediaPackage(original.toString('base64'), pictureId, {
        mime: 'image/png',
        base64: replacement.toString('base64'),
      })
    ).base64,
    'base64',
  )
  expect(revised.equals(original)).toBe(false)
  const userDataPath = mkdtempSync(join(tmpdir(), 'wiswork-p0-13-savepoints-'))
  roots.push(userDataPath)
  const request = async (
    operation: string,
    body: Record<string, unknown>,
    service = createPresentationService({ userDataPath }),
  ) =>
    JSON.parse(
      Buffer.from(
        await service(
          { operation, documentId: 'P0-13-document', ...body },
          new AbortController().signal,
        ),
      ).toString(),
    ) as Record<string, unknown>
  for (const [backupId, value] of [
    ['p013-original', original],
    ['p013-revised', revised],
  ] as const) {
    const scope = {
      backupId,
      hostSlideId: 'host-page-4',
      slideIds: ['host-page-1', 'host-page-4', 'host-page-8'],
      sha256: sha(value),
      sizeBytes: value.length,
    }
    expect(await request('existing_page_backup_begin', scope)).toMatchObject({
      status: 'uploading',
      receivedBytes: 0,
    })
    for (let offset = 0; offset < value.length; offset += 128 * 1024)
      await request('existing_page_backup_chunk', {
        backupId,
        offset,
        base64: value.subarray(offset, offset + 128 * 1024).toString('base64'),
      })
    expect(await request('existing_page_backup_finish', { backupId })).toMatchObject({
      status: 'ready',
      sha256: sha(value),
    })
  }
  const restarted = createPresentationService({ userDataPath })
  for (const [backupId, value] of [
    ['p013-original', original],
    ['p013-revised', revised],
  ] as const) {
    const chunks: Buffer[] = []
    for (let offset = 0; offset < value.length; offset += 128 * 1024) {
      const read = await request(
        'existing_page_backup_read',
        {
          backupId,
          offset,
          length: Math.min(128 * 1024, value.length - offset),
        },
        restarted,
      )
      chunks.push(Buffer.from(read.base64 as string, 'base64'))
    }
    expect(Buffer.concat(chunks).equals(value)).toBe(true)
  }
})
it('renders only a ready, document-scoped page package and returns its identity', async () => {
  const f = await fixture()
  const png = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII=',
    'base64',
  )
  const renderPage = async (raw: Uint8Array) => {
    expect(Buffer.from(raw)).toEqual(f.raw)
    return png
  }
  const service = createPresentationService({ userDataPath: f.userDataPath, renderPage })
  const call = (operation: string, body: Record<string, unknown> = {}, documentId = 'document-1') =>
    f.call(operation, body, documentId, service)
  await call('existing_page_backup_begin', f.begin)
  expect(await call('existing_page_backup_render', { backupId: f.begin.backupId })).toEqual({
    error: 'page_not_ready',
  })
  await call('existing_page_backup_chunk', {
    backupId: f.begin.backupId,
    offset: 0,
    base64: f.raw.toString('base64'),
  })
  await call('existing_page_backup_finish', { backupId: f.begin.backupId })
  expect(
    await call('existing_page_backup_render', { backupId: f.begin.backupId }, 'other'),
  ).toEqual({ error: 'not_found' })
  expect(await call('existing_page_backup_render', { backupId: f.begin.backupId })).toEqual({
    backupId: f.begin.backupId,
    hostSlideId: f.begin.hostSlideId,
    sha256: f.begin.sha256,
    renderer: 'libreoffice',
    mime: 'image/png',
    base64: png.toString('base64'),
  })
  await call('existing_page_backup_release', f.begin)
  expect(await call('existing_page_backup_render', { backupId: f.begin.backupId })).toEqual({
    error: 'not_found',
  })
})

it('cancels a queued page-savepoint read without overtaking an active render', async () => {
  const f = await fixture()
  await f.call('existing_page_backup_begin', f.begin)
  await f.call('existing_page_backup_chunk', {
    backupId: f.begin.backupId,
    offset: 0,
    base64: f.raw.toString('base64'),
  })
  await f.call('existing_page_backup_finish', { backupId: f.begin.backupId })
  let rendering!: () => void
  let finish!: (image: Uint8Array) => void
  const started = new Promise<void>((resolve) => {
    rendering = resolve
  })
  const image = new Promise<Uint8Array>((resolve) => {
    finish = resolve
  })
  const service = createPresentationService({
    userDataPath: f.userDataPath,
    renderPage: async () => {
      rendering()
      return image
    },
  })
  const body = (operation: string) => ({
    operation,
    documentId: 'document-1',
    backupId: f.begin.backupId,
  })
  const first = service(body('existing_page_backup_render'), new AbortController().signal)
  await started
  const controller = new AbortController()
  const second = service(body('existing_page_backup_status'), controller.signal)
  controller.abort()
  expect(JSON.parse(Buffer.from(await second).toString())).toEqual({ error: 'aborted' })
  const third = service(body('existing_page_backup_status'), new AbortController().signal)
  finish(
    Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII=',
      'base64',
    ),
  )
  expect(JSON.parse(Buffer.from(await first).toString())).toMatchObject({ mime: 'image/png' })
  expect(JSON.parse(Buffer.from(await third).toString())).toMatchObject({
    status: 'ready',
    sha256: f.begin.sha256,
  })
})
it('abandons only a matching incomplete fallback upload without releasing a ready savepoint', async () => {
  const f = await fixture()
  await f.call('existing_page_backup_begin', f.begin)
  await f.call('existing_page_backup_chunk', {
    backupId: f.begin.backupId,
    offset: 0,
    base64: f.raw.subarray(0, 1024).toString('base64'),
  })
  expect(
    await f.call('existing_page_backup_abandon', { ...f.begin, hostSlideId: 'other' }),
  ).toEqual({ error: 'invalid_request' })
  expect(await f.call('existing_page_backup_abandon', f.begin)).toEqual({
    backupId: f.begin.backupId,
    documentId: 'document-1',
    status: 'abandoned',
  })
  expect(await f.call('existing_page_backup_status', { backupId: f.begin.backupId })).toEqual({
    error: 'not_found',
  })
  await f.call('existing_page_backup_begin', f.begin)
  await f.call('existing_page_backup_chunk', {
    backupId: f.begin.backupId,
    offset: 0,
    base64: f.raw.toString('base64'),
  })
  await f.call('existing_page_backup_finish', { backupId: f.begin.backupId })
  expect(await f.call('existing_page_backup_abandon', f.begin)).toEqual({
    error: 'request_conflict',
  })
})
it('requires scoped identity and refuses conflicting page/order, document and size', async () => {
  const f = await fixture()
  expect(await f.call('existing_page_backup_begin', { ...f.begin, projectId: 'foreign' })).toEqual({
    error: 'invalid_request',
  })
  expect(
    await f.call('existing_page_backup_begin', { ...f.begin, slideIds: ['host-1', 'host-1'] }),
  ).toEqual({ error: 'invalid_request' })
  expect(
    await f.call('existing_page_backup_begin', { ...f.begin, hostSlideId: 'missing' }),
  ).toEqual({ error: 'invalid_request' })
  expect(
    await f.call('existing_page_backup_begin', { ...f.begin, sizeBytes: 8 * 1024 * 1024 + 1 }),
  ).toEqual({ error: 'invalid_request' })
  await f.call('existing_page_backup_begin', f.begin)
  expect(
    await f.call('existing_page_backup_begin', { ...f.begin, slideIds: ['host-2', 'host-1'] }),
  ).toEqual({ error: 'request_conflict' })
  expect(
    await f.call('existing_page_backup_status', { backupId: f.begin.backupId }, 'another-document'),
  ).toEqual({ error: 'not_found' })
  expect(
    await f.call('existing_page_backup_chunk', {
      backupId: f.begin.backupId,
      offset: 1,
      base64: 'YQ==',
    }),
  ).toEqual({ error: 'request_conflict' })
})
it('rejects package corruption and rechecks persisted ready bytes', async () => {
  const f = await fixture()
  await f.call('existing_page_backup_begin', f.begin)
  await f.call('existing_page_backup_chunk', {
    backupId: f.begin.backupId,
    offset: 0,
    base64: f.raw.toString('base64'),
  })
  await f.call('existing_page_backup_finish', { backupId: f.begin.backupId })
  const path = join(
    f.userDataPath,
    'presentation-existing-page-backups',
    sha('document-1'),
    sha(f.begin.backupId),
    'raw.pptx',
  )
  writeFileSync(path, Buffer.alloc(f.raw.length))
  expect(
    await f.call('existing_page_backup_read', { backupId: f.begin.backupId, offset: 0, length: 1 }),
  ).toEqual({ error: 'digest_mismatch' })
})
it('resumes a partial upload after service restart and rejects a non-PPTX package', async () => {
  const f = await fixture()
  const invalid = Buffer.from('not-a-pptx')
  const begin = {
    ...f.begin,
    backupId: 'invalid-package',
    sha256: sha(invalid),
    sizeBytes: invalid.length,
  }
  await f.call('existing_page_backup_begin', begin)
  await f.call('existing_page_backup_chunk', {
    backupId: begin.backupId,
    offset: 0,
    base64: invalid.subarray(0, 3).toString('base64'),
  })
  expect(await f.call('existing_page_backup_begin', begin)).toMatchObject({ receivedBytes: 3 })
  await f.call('existing_page_backup_chunk', {
    backupId: begin.backupId,
    offset: 3,
    base64: invalid.subarray(3).toString('base64'),
  })
  expect(await f.call('existing_page_backup_finish', { backupId: begin.backupId })).toEqual({
    error: 'unsupported_file',
  })
})
it('reserves room for eight original pages and eight restore transactions, then reuses quota', async () => {
  const f = await fixture()
  for (let i = 0; i < 16; i++) {
    const begin = { ...f.begin, backupId: `backup-${i}` }
    expect(await f.call('existing_page_backup_begin', begin)).toMatchObject({ status: 'uploading' })
    await f.call('existing_page_backup_chunk', {
      backupId: begin.backupId,
      offset: 0,
      base64: f.raw.toString('base64'),
    })
    expect(await f.call('existing_page_backup_finish', { backupId: begin.backupId })).toMatchObject(
      {
        status: 'ready',
      },
    )
  }
  expect(
    await f.call('existing_page_backup_begin', { ...f.begin, backupId: 'seventeenth' }),
  ).toEqual({ error: 'quota_exceeded' })
  const scope = { ...f.begin, backupId: 'backup-0' }
  expect(await f.call('existing_page_backup_release', { ...scope, sha256: sha('wrong') })).toEqual({
    error: 'request_conflict',
  })
  expect(
    await f.call('existing_page_backup_release', { ...scope, slideIds: ['host-2', 'host-1'] }),
  ).toEqual({ error: 'request_conflict' })
  expect(await f.call('existing_page_backup_release', scope)).toMatchObject({
    backupId: 'backup-0',
    status: 'released',
  })
  const residual = join(
    f.userDataPath,
    'presentation-existing-page-backups',
    sha('document-1'),
    sha('backup-0'),
  )
  mkdirSync(residual)
  expect(await f.call('existing_page_backup_release', scope)).toMatchObject({
    backupId: 'backup-0',
    status: 'released',
  })
  expect(existsSync(residual)).toBe(false)
  expect(await f.call('existing_page_backup_release', scope)).toMatchObject({
    backupId: 'backup-0',
    status: 'released',
  })
  expect(
    await f.call('existing_page_backup_release', { ...scope, sizeBytes: scope.sizeBytes - 1 }),
  ).toEqual({ error: 'request_conflict' })
  expect(await f.call('existing_page_backup_status', { backupId: 'backup-0' })).toEqual({
    error: 'not_found',
  })
  expect(await f.call('existing_page_backup_begin', scope)).toEqual({ error: 'request_conflict' })
  expect(
    await f.call('existing_page_backup_begin', { ...f.begin, backupId: 'seventeenth' }),
  ).toMatchObject({ status: 'uploading' })
})
it('refuses unready, missing, foreign and damaged backup release', async () => {
  const f = await fixture()
  expect(await f.call('existing_page_backup_release', f.begin)).toEqual({ error: 'not_found' })
  await f.call('existing_page_backup_begin', f.begin)
  expect(await f.call('existing_page_backup_release', f.begin)).toEqual({ error: 'page_not_ready' })
  await f.call('existing_page_backup_chunk', {
    backupId: f.begin.backupId,
    offset: 0,
    base64: f.raw.toString('base64'),
  })
  await f.call('existing_page_backup_finish', { backupId: f.begin.backupId })
  expect(await f.call('existing_page_backup_release', f.begin, 'foreign')).toEqual({
    error: 'not_found',
  })
  const path = join(
    f.userDataPath,
    'presentation-existing-page-backups',
    sha('document-1'),
    sha(f.begin.backupId),
    'raw.pptx',
  )
  writeFileSync(path, Buffer.alloc(f.raw.length))
  expect(await f.call('existing_page_backup_release', f.begin)).toEqual({
    error: 'digest_mismatch',
  })
})
it('lists only active backup metadata for the requested document', async () => {
  const f = await fixture()
  expect(await f.call('existing_page_backup_list')).toEqual({
    documentId: 'document-1',
    backups: [],
  })
  await f.call('existing_page_backup_begin', f.begin)
  expect(await f.call('existing_page_backup_list')).toMatchObject({
    backups: [{ backupId: f.begin.backupId, status: 'uploading' }],
  })
  expect(await f.call('existing_page_backup_list', {}, 'another-document')).toEqual({
    documentId: 'another-document',
    backups: [],
  })
  expect(await f.call('existing_page_backup_list', { backupId: f.begin.backupId })).toEqual({
    error: 'invalid_request',
  })
})
