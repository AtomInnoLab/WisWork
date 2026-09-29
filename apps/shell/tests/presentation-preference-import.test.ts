import { afterEach, expect, it } from 'vitest'
import { mkdtempSync, rmSync, readFileSync, writeFileSync, symlinkSync, readdirSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { createHash } from 'node:crypto'
import { createPresentationService } from '../src/main/presentation-service'
import { PresentationPreferenceLibrary } from '../src/main/presentation-preferences'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const digest = (text: string) => createHash('sha256').update(text).digest('hex')
const source = { documentId: 'source-doc', projectId: 'source-project', changeId: 'original-edit' }
const original = { projectId: source.projectId, changeId: source.changeId, text: 'Short titles' }
const approvalId = '12345678-1234-4123-8123-123456789abc'
async function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'pref-import-'))
  roots.push(root)
  let service = createPresentationService({ userDataPath: root })
  const call = async (body: unknown) =>
    JSON.parse(Buffer.from(await service(body, new AbortController().signal)).toString('utf8'))
  expect(
    await call({
      operation: 'preference_save',
      documentId: source.documentId,
      preference: original,
    }),
  ).toEqual({ preference: original })
  const request = {
    operation: 'preference_import',
    documentId: 'target-doc',
    projectId: 'target-project',
    source,
    expectedTextDigest: digest(original.text),
    approvalId,
  }
  return {
    root,
    call,
    request,
    restart: () => {
      service = createPresentationService({ userDataPath: root })
    },
  }
}
it('imports exact selected source with durable first approval, separate target deletion and no global listing', async () => {
  const f = await fixture()
  expect(await f.call({ operation: 'preference_get', ...source })).toEqual({ preference: original })
  const result = await f.call(f.request)
  expect(result.preference).toMatchObject({
    projectId: f.request.projectId,
    text: original.text,
    reuse: { version: 1, source, sourceTextDigest: digest(original.text), approvalId },
  })
  expect(result.preference.changeId).toBe(
    'reuse_' +
      digest(
        JSON.stringify([
          f.request.documentId,
          f.request.projectId,
          source.documentId,
          source.projectId,
          source.changeId,
          digest(original.text),
        ]),
      ),
  )
  expect(new Date(result.preference.reuse.approvedAt).toISOString()).toBe(
    result.preference.reuse.approvedAt,
  )
  f.restart()
  expect(
    await f.call({ ...f.request, approvalId: '22345678-1234-4123-8123-123456789abc' }),
  ).toEqual(result)
  expect(
    await f.call({
      operation: 'preference_list',
      documentId: 'other-doc',
      projectId: f.request.projectId,
    }),
  ).toEqual({ preferences: [] })
  expect(
    await f.call({
      operation: 'preference_delete',
      documentId: f.request.documentId,
      projectId: f.request.projectId,
      changeId: result.preference.changeId,
    }),
  ).toEqual({ deleted: true })
  expect(await f.call({ operation: 'preference_get', ...source })).toEqual({ preference: original })
})
it('rejects wrong snapshot, same scope, forged metadata and copying an imported source', async () => {
  const f = await fixture()
  expect(await f.call({ ...f.request, expectedTextDigest: 'a'.repeat(64) })).toEqual({
    error: 'revision_conflict',
  })
  expect(
    await f.call({ ...f.request, documentId: source.documentId, projectId: source.projectId }),
  ).toEqual({ error: 'invalid_request' })
  expect(await f.call({ ...f.request, approvedAt: '2026-09-29T00:00:00.000Z' })).toEqual({
    error: 'invalid_request',
  })
  const result = await f.call(f.request)
  expect(
    await f.call({
      operation: 'preference_save',
      documentId: f.request.documentId,
      preference: result.preference,
    }),
  ).toEqual({ error: 'invalid_request' })
  expect(
    await f.call({
      ...f.request,
      documentId: 'third-doc',
      source: {
        documentId: f.request.documentId,
        projectId: f.request.projectId,
        changeId: result.preference.changeId,
      },
    }),
  ).toEqual({ error: 'invalid_request' })
})
it('requires a fresh existing unchanged source even for idempotent receipt retries', async () => {
  const f = await fixture()
  await f.call(f.request)
  await f.call({ operation: 'preference_delete', ...source })
  expect(await f.call(f.request)).toEqual({ error: 'not_found' })
  await f.call({
    operation: 'preference_save',
    documentId: source.documentId,
    preference: { ...original, text: 'Changed' },
  })
  expect(await f.call(f.request)).toEqual({ error: 'revision_conflict' })
})
it('preserves V1 disk shape, rejects corruption/symlinks and keeps original quotas', async () => {
  const f = await fixture(),
    library = new PresentationPreferenceLibrary(f.root)
  const filename = readdirSync(join(f.root, 'presentation-preferences'))[0]!,
    path = join(f.root, 'presentation-preferences', filename)
  expect(JSON.parse(readFileSync(path, 'utf8'))).toEqual({
    version: 1,
    documentId: source.documentId,
    projectId: source.projectId,
    preferences: [original],
  })
  for (let i = 1; i < 64; i++)
    library.save(source.documentId, { ...original, changeId: 'edit_' + i })
  expect(() => library.save(source.documentId, { ...original, changeId: 'over' })).toThrow(
    'quota_exceeded',
  )
  const saved = readFileSync(path)
  writeFileSync(path, 'bad json')
  expect(await f.call(f.request)).toEqual({ error: 'invalid_state' })
  rmSync(path)
  writeFileSync(path + '.real', saved)
  symlinkSync(path + '.real', path)
  expect(await f.call(f.request)).toEqual({ error: 'invalid_state' })
})
it('retries after a lost service response and source key reordering without replacing the first approval', async () => {
  const f = await fixture()
  const lost = async () => {
    await f.call(f.request)
    throw Error('lost response')
  }
  await expect(lost()).rejects.toThrow('lost response')
  f.restart()
  const retry = await f.call({
    ...f.request,
    source: {
      changeId: source.changeId,
      projectId: source.projectId,
      documentId: source.documentId,
    },
    approvalId: '22345678-1234-4123-8123-123456789abc',
  })
  expect(retry.preference.reuse.approvalId).toBe(approvalId)
  expect(
    (
      await f.call({
        operation: 'preference_list',
        documentId: f.request.documentId,
        projectId: f.request.projectId,
      })
    ).preferences,
  ).toHaveLength(1)
})
it('bounds imported metadata by the unchanged 64KiB quota before reaching 64 records', async () => {
  const f = await fixture(),
    library = new PresentationPreferenceLibrary(f.root),
    longSource = { ...source, projectId: 'long-source-project', documentId: 'd'.repeat(2048) }
  let accepted = 0
  for (let i = 0; i < 64; i++) {
    const changeId = 'item_' + i
    library.save(longSource.documentId, { ...original, projectId: longSource.projectId, changeId })
    const result = await f.call({ ...f.request, source: { ...longSource, changeId } })
    if (result.error) {
      expect(result.error).toBe('quota_exceeded')
      break
    }
    accepted++
  }
  expect(accepted).toBeGreaterThan(0)
  expect(accepted).toBeLessThan(64)
  const items = library.list(f.request.documentId, f.request.projectId)
  expect(items).toHaveLength(accepted)
})
it('rejects corrupt target provenance identity and unknown get/import keys', async () => {
  const f = await fixture()
  await f.call(f.request)
  const filename = digest(JSON.stringify([f.request.documentId, f.request.projectId])) + '.json',
    path = join(f.root, 'presentation-preferences', filename)
  const stored = JSON.parse(readFileSync(path, 'utf8'))
  stored.preferences[0].reuse.sourceTextDigest = '0'.repeat(64)
  writeFileSync(path, JSON.stringify(stored))
  expect(
    await f.call({
      operation: 'preference_list',
      documentId: f.request.documentId,
      projectId: f.request.projectId,
    }),
  ).toEqual({ error: 'invalid_state' })
  expect(await f.call({ operation: 'preference_get', ...source, approved: true })).toEqual({
    error: 'invalid_request',
  })
  expect(await f.call({ ...f.request, source: { ...source, extra: 'private-value' } })).toEqual({
    error: 'invalid_request',
  })
  expect(await f.call({ operation: 'preference_get', ...source, changeId: 'missing' })).toEqual({})
})
it('rejects a legacy identity collision and ordinary save masquerading as an imported preference', async () => {
  const f = await fixture(),
    library = new PresentationPreferenceLibrary(f.root)
  const changeId =
    'reuse_' +
    digest(
      JSON.stringify([
        f.request.documentId,
        f.request.projectId,
        source.documentId,
        source.projectId,
        source.changeId,
        digest(original.text),
      ]),
    )
  library.save(f.request.documentId, {
    projectId: f.request.projectId,
    changeId,
    text: original.text,
  })
  expect(await f.call(f.request)).toEqual({ error: 'revision_conflict' })
  library.delete(f.request.documentId, f.request.projectId, changeId)
  const imported = await f.call(f.request)
  expect(
    await f.call({
      operation: 'preference_save',
      documentId: f.request.documentId,
      preference: {
        projectId: f.request.projectId,
        changeId: imported.preference.changeId,
        text: original.text,
      },
    }),
  ).toEqual({ error: 'revision_conflict' })
  expect(library.get(f.request.documentId, f.request.projectId, changeId)).toEqual(
    imported.preference,
  )
})
