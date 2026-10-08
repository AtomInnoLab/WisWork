import { createHash } from 'node:crypto'
import { mkdtempSync, rmSync, writeFileSync, symlinkSync, unlinkSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, expect, it } from 'vitest'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'
import { createPresentationPageBackupService } from '../src/main/presentation-page-backups'
const roots: string[] = []
afterEach(() => roots.splice(0).forEach((root) => rmSync(root, { recursive: true, force: true })))
const hash = (bytes: string | Buffer) => createHash('sha256').update(bytes).digest('hex')
async function setup() {
  const userDataPath = mkdtempSync(join(tmpdir(), 'wiswork-page-backup-'))
  roots.push(userDataPath)
  const deck = benchmarkPlannedDeck(),
    pageId = deck.slides[0]!.id
  const service = createPresentationService({ userDataPath })
  const run = async (operation: string, extra = {}) =>
    JSON.parse(
      Buffer.from(
        await service(
          { operation, documentId: 'doc', projectId: deck.id, ...extra },
          new AbortController().signal,
        ),
      ).toString(),
    )
  await run('save_plan', { expectedRevision: 0, plan: benchmarkPlan() })
  await run('production_begin', { requestId: 'parent', planRevision: 1, deck })
  await run('production_run', { requestId: 'parent' })
  await run('production_rebuild_page', {
    parentRequestId: 'parent',
    requestId: 'child',
    pageId,
    slide: { ...deck.slides[0], notes: 'revision' },
  })
  await run('production_run', { requestId: 'child' })
  const raw = Buffer.from(
    (await compilePresentationDeck({ ...deck, slides: [deck.slides[0]] })).bytes,
  )
  const body = {
    backupId: 'backup-1',
    requestId: 'child',
    pageId,
    hostSlideId: '256',
    slideIds: ['256', '257'],
    sha256: hash(raw),
    sizeBytes: raw.length,
  }
  const backup = createPresentationPageBackupService({ userDataPath })
  const call = (operation: string, extra = {}, signal = new AbortController().signal) =>
    backup({ operation, documentId: 'doc', projectId: deck.id, ...extra }, signal)
  const dir = join(userDataPath, 'presentation-page-backups', hash(deck.id), hash(body.backupId))
  return { userDataPath, run, raw, body, call, dir, projectId: deck.id }
}
it('persists original single-page bytes, replays overlapping chunks and survives restart', async () => {
  const f = await setup()
  expect(await f.call('page_backup_begin', f.body)).toMatchObject({
    status: 'uploading',
    receivedBytes: 0,
    parentRequestId: 'parent',
  })
  const half = Math.floor(f.raw.length / 2)
  await f.call('page_backup_chunk', {
    backupId: f.body.backupId,
    offset: 0,
    base64: f.raw.subarray(0, half).toString('base64'),
  })
  await f.call('page_backup_chunk', {
    backupId: f.body.backupId,
    offset: 0,
    base64: f.raw.toString('base64'),
  })
  expect(await f.call('page_backup_finish', { backupId: f.body.backupId })).toMatchObject({
    status: 'ready',
    receivedBytes: f.raw.length,
  })
  const restarted = createPresentationPageBackupService({ userDataPath: f.userDataPath })
  const result = await restarted(
    {
      operation: 'page_backup_read',
      documentId: 'doc',
      projectId: f.projectId,
      backupId: f.body.backupId,
      offset: 0,
      length: 131072,
    },
    new AbortController().signal,
  )
  expect(result).toEqual({
    backupId: f.body.backupId,
    offset: 0,
    sizeBytes: f.raw.length,
    sha256: hash(f.raw),
    base64: f.raw.toString('base64'),
  })
  expect(await f.call('page_backup_begin', f.body)).toMatchObject({ status: 'ready' })
})
it('rejects identity conflicts, unexpected keys, unready tasks and document changes', async () => {
  const f = await setup()
  await expect(f.call('page_backup_begin', { ...f.body, requestId: 'parent' })).rejects.toThrow(
    'invalid_request',
  )
  await expect(
    f.call('page_backup_begin', { ...f.body, slideIds: ['256', '256'] }),
  ).rejects.toThrow('invalid_request')
  await expect(f.call('page_backup_begin', { ...f.body, hostSlideId: 'missing' })).rejects.toThrow(
    'invalid_request',
  )
  await expect(f.call('page_backup_begin', { ...f.body, extra: 1 })).rejects.toThrow(
    'invalid_request',
  )
  await f.call('page_backup_begin', f.body)
  await expect(
    f.call('page_backup_begin', { ...f.body, slideIds: ['257', '256'] }),
  ).rejects.toThrow('request_conflict')
  await expect(
    f.call('page_backup_status', { backupId: f.body.backupId, documentId: 'other' }),
  ).rejects.toThrow('document_mismatch')
  await expect(
    f.call('page_backup_chunk', { backupId: f.body.backupId, offset: 1, base64: 'YWJj' }),
  ).rejects.toThrow('request_conflict')
  await expect(
    f.call('page_backup_read', { backupId: f.body.backupId, offset: 0, length: 1 }),
  ).rejects.toThrow('page_not_ready')
  await expect(f.call('page_backup_finish', { backupId: f.body.backupId })).rejects.toThrow(
    'invalid_state',
  )
  const aborted = new AbortController()
  aborted.abort()
  await expect(
    f.call('page_backup_status', { backupId: f.body.backupId }, aborted.signal),
  ).rejects.toThrow('aborted')
})
it('caps pending backups and rejects damaged, linked or non-PPTX bytes', async () => {
  const f = await setup()
  for (let i = 0; i < 8; i++)
    await f.call('page_backup_begin', { ...f.body, backupId: `backup-${i}` })
  await expect(f.call('page_backup_begin', { ...f.body, backupId: 'ninth' })).rejects.toThrow(
    'quota_exceeded',
  )
  await f.call('page_backup_chunk', {
    backupId: f.body.backupId,
    offset: 0,
    base64: f.raw.toString('base64'),
  })
  await f.call('page_backup_finish', { backupId: f.body.backupId })
  writeFileSync(join(f.dir, 'raw.pptx'), Buffer.alloc(f.raw.length))
  await expect(
    f.call('page_backup_read', { backupId: f.body.backupId, offset: 0, length: 1 }),
  ).rejects.toThrow('digest_mismatch')
  unlinkSync(join(f.dir, 'raw.pptx'))
  symlinkSync('/etc/passwd', join(f.dir, 'raw.pptx'))
  await expect(f.call('page_backup_status', { backupId: f.body.backupId })).rejects.toThrow(
    'invalid_state',
  )
})
it('requires a real single-page PPTX and rejects bomb central directory sizes', async () => {
  const f = await setup()
  const central = f.raw.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]))
  const extra = central + 46 + f.raw.readUInt16LE(central + 28)
  const zip64 = Buffer.concat([
    f.raw.subarray(0, extra),
    Buffer.from([1, 0, 0, 0]),
    f.raw.subarray(extra),
  ])
  zip64.writeUInt16LE(f.raw.readUInt16LE(central + 30) + 4, central + 30)
  const end = zip64.length - 22
  zip64.writeUInt32LE(zip64.readUInt32LE(end + 12) + 4, end + 12)
  const multiDisk = Buffer.from(f.raw)
  multiDisk.writeUInt16LE(1, multiDisk.length - 22 + 6)
  for (const [id, raw] of [
    ['zip64', zip64],
    ['multidisk', multiDisk],
    ['text', Buffer.from('not pptx')],
    ['multi', Buffer.from((await compilePresentationDeck(benchmarkPlannedDeck())).bytes)],
    ['bomb', Buffer.from(f.raw)],
  ] as const) {
    if (id === 'bomb') {
      const position = raw.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]))
      raw.writeUInt32LE(0x7fffffff, position + 24)
    }
    await f.call('page_backup_begin', {
      ...f.body,
      backupId: id,
      sizeBytes: raw.length,
      sha256: hash(raw),
    })
    for (let offset = 0; offset < raw.length; offset += 131072)
      await f.call('page_backup_chunk', {
        backupId: id,
        offset,
        base64: raw.subarray(offset, offset + 131072).toString('base64'),
      })
    await expect(f.call('page_backup_finish', { backupId: id })).rejects.toThrow('unsupported_file')
  }
})
it('resumes an interrupted upload after restart and rejects chunk mutation and metadata corruption', async () => {
  const f = await setup()
  await f.call('page_backup_begin', f.body)
  const half = Math.floor(f.raw.length / 2)
  await f.call('page_backup_chunk', {
    backupId: f.body.backupId,
    offset: 0,
    base64: f.raw.subarray(0, half).toString('base64'),
  })
  const restart = createPresentationPageBackupService({ userDataPath: f.userDataPath })
  const call = (operation: string, extra = {}) =>
    restart(
      { operation, projectId: f.projectId, documentId: 'doc', backupId: f.body.backupId, ...extra },
      new AbortController().signal,
    )
  expect(await call('page_backup_status')).toMatchObject({
    receivedBytes: half,
    status: 'uploading',
  })
  await expect(
    call('page_backup_chunk', { offset: 0, base64: Buffer.alloc(half).toString('base64') }),
  ).rejects.toThrow('request_conflict')
  await expect(call('page_backup_chunk', { offset: half, base64: 'abc!' })).rejects.toThrow(
    'invalid_request',
  )
  await call('page_backup_chunk', { offset: half, base64: f.raw.subarray(half).toString('base64') })
  await call('page_backup_finish')
  await expect(call('page_backup_read', { offset: 0, length: 131073 })).rejects.toThrow(
    'invalid_request',
  )
  writeFileSync(
    join(f.dir, 'metadata.json'),
    JSON.stringify({
      ...f.body,
      projectId: f.projectId,
      documentId: 'doc',
      parentRequestId: 'parent',
      parentInputDigest: '0'.repeat(64),
      inputDigest: '0'.repeat(64),
      status: 'ready',
    }),
  )
  await expect(call('page_backup_status')).rejects.toThrow('invalid_state')
})
it('rejects uncompiled derived tasks, oversized backups and linked project directories', async () => {
  const f = await setup()
  const deck = benchmarkPlannedDeck()
  await f.run('production_rebuild_page', {
    parentRequestId: 'parent',
    requestId: 'pending-child',
    pageId: f.body.pageId,
    slide: { ...deck.slides[0], notes: 'still pending' },
  })
  await expect(
    f.call('page_backup_begin', { ...f.body, requestId: 'pending-child' }),
  ).rejects.toThrow('page_not_ready')
  await expect(
    f.call('page_backup_begin', { ...f.body, sizeBytes: 8 * 1024 * 1024 + 1 }),
  ).rejects.toThrow('invalid_request')
  await f.call('page_backup_begin', f.body)
  const project = join(f.userDataPath, 'presentation-page-backups', hash(f.projectId))
  rmSync(project, { recursive: true, force: true })
  symlinkSync(tmpdir(), project)
  await expect(f.call('page_backup_status', { backupId: f.body.backupId })).rejects.toThrow(
    'invalid_state',
  )
})
