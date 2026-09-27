import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, expect, it } from 'vitest'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkPlannedDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
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
