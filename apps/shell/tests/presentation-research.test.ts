import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { afterEach, expect, it, vi } from 'vitest'
import { createPresentationAttachmentService } from '../src/main/presentation-attachments'
import { createPresentationResearchService } from '../src/main/presentation-research'
import { createPresentationService } from '../src/main/presentation-service'
import {
  parsePresentationResearchRecord,
  parsePresentationResearchSummary,
} from '@wiswork/project-store/presentation-research'
import { researchDraft } from '../../../packages/project-store/tests/fixtures/presentation-research'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const hash = (v: string | Uint8Array) => createHash('sha256').update(v).digest('hex')
function setup() {
  const root = mkdtempSync(join(tmpdir(), 'research-pc-'))
  roots.push(root)
  const attachments = createPresentationAttachmentService({
    userDataPath: root,
    fetchPage: async () =>
      new Response('<html><p>收入增长仅是管理层预测。</p><p>另一个段落。</p></html>', {
        headers: { 'content-type': 'text/html' },
      }),
  })
  const service = createPresentationResearchService({ userDataPath: root, attachments })
  const call = (
    operation: string,
    fields: Record<string, unknown> = {},
    signal = new AbortController().signal,
  ) =>
    service(
      {
        operation: 'research_' + operation,
        documentId: 'doc',
        ...(operation === 'capabilities' ? {} : { projectId: 'research-project' }),
        ...fields,
      },
      signal,
    )
  return { root, attachments, call }
}
async function upload(
  attachments: ReturnType<typeof createPresentationAttachmentService>,
  text: string,
) {
  const raw = Buffer.from(text),
    id = hash(raw),
    call = (operation: string, body: Record<string, unknown>) =>
      attachments(
        { operation, documentId: 'doc', attachmentId: id, ...body },
        new AbortController().signal,
      )
  await call('attachment_begin', { name: 'original.txt', sizeBytes: raw.length, sha256: id })
  await call('attachment_chunk', { offset: 0, base64: raw.toString('base64') })
  await call('attachment_finish', {})
  return id
}
const record = (response: unknown) =>
  parsePresentationResearchRecord((response as { record: unknown }).record)
it('builds before any saved plan from actual uploaded originals and preserves inference/conflict boundaries on restart', async () => {
  const f = setup()
  const id = await upload(f.attachments, '收入增长仅是管理层预测。')
  const draft = researchDraft(id)
  const result = await f.call('build', { ledgerId: 'ledger', expectedRevision: 0, draft })
  const r = record(result)
  expect(r).toMatchObject({
    state: 'completed',
    checks: {
      support: 'not_verified',
      sourceAuthority: 'not_verified',
      timeliness: 'not_verified',
    },
    sources: [
      {
        sourceId: 'original',
        attachmentId: id,
        status: 'found',
        offset: 0,
        provenance: 'user_supplied',
        sha256: id,
      },
    ],
  })
  expect(r.sources![0]!.retrievedAt).toBeUndefined()
  expect(r.draft.facts).toEqual(draft.facts)
  const restart = createPresentationResearchService({ userDataPath: f.root })
  expect(
    await restart(
      {
        operation: 'research_read',
        documentId: 'doc',
        projectId: 'research-project',
        ledgerId: 'ledger',
      },
      new AbortController().signal,
    ),
  ).toEqual(r)
  expect(await f.call('build', { ledgerId: 'ledger', expectedRevision: 0, draft })).toEqual(result)
  expect(parsePresentationResearchSummary(await f.call('list')).revision).toBe(2)
  await expect(f.call('read', { ledgerId: 'ledger', documentId: 'other' })).rejects.toThrow(
    'not_found',
  )
})
it('uses exact fetched URL provenance and records snapshot, excerpt and locator gaps', async () => {
  const f = setup()
  const uri = 'https://8.8.8.8/research?edition=1'
  const imported = (await f.attachments(
    { operation: 'attachment_import_webpage', documentId: 'doc', url: uri },
    new AbortController().signal,
  )) as { attachmentId: string; retrievedAt: number }
  const draft = researchDraft()
  draft.sources = [
    {
      id: 'original',
      title: '网页原文',
      uri,
      snapshotAttachmentId: imported.attachmentId,
      excerpt: '收入增长仅是管理层预测。',
      locator: '第 1 段',
    },
    { id: 'missing', title: '缺失快照', uri, excerpt: '原文' },
    {
      id: 'wrong-url',
      title: '错版',
      uri: uri.replace('1', '2'),
      snapshotAttachmentId: imported.attachmentId,
      excerpt: '收入增长仅是管理层预测。',
    },
    {
      id: 'wrong-locator',
      title: '错段',
      uri,
      snapshotAttachmentId: imported.attachmentId,
      excerpt: '收入增长仅是管理层预测。',
      locator: '第 2 段',
    },
    {
      id: 'not-found',
      title: '摘要不得作为原文',
      uri,
      snapshotAttachmentId: imported.attachmentId,
      excerpt: '已验证增长',
    },
    { id: 'empty', title: '无摘录', uri, snapshotAttachmentId: imported.attachmentId, excerpt: '' },
  ]
  const r = record(await f.call('build', { ledgerId: 'ledger', expectedRevision: 0, draft }))
  expect(r.state).toBe('completed')
  expect(r.sources?.map((s) => s.status)).toEqual([
    'found',
    'missing',
    'source_mismatch',
    'not_found',
    'not_found',
    'empty_excerpt',
  ])
  expect(r.sources![0]).toMatchObject({
    provenance: 'fetched_url_matched',
    retrievedAt: new Date(imported.retrievedAt).toISOString(),
    sha256: imported.attachmentId,
  })
  expect(r.sources![2]?.provenance).toBe('unavailable')
})
it('publishes running before work, permits read while pending, persists cancellation and never replays the same attempt', async () => {
  const f = setup(),
    id = await upload(f.attachments, '收入增长仅是管理层预测。')
  let release!: () => void, entered!: () => void
  const gate = new Promise<void>((r) => {
      release = r
    }),
    started = new Promise<void>((r) => {
      entered = r
    })
  const attachments = vi.fn(async (body: Record<string, unknown>, signal: AbortSignal) => {
    if (body.operation === 'attachment_match_excerpt') {
      entered()
      await gate
    }
    return f.attachments(body, signal)
  })
  const service = createPresentationResearchService({ userDataPath: f.root, attachments })
  const controller = new AbortController()
  const request = {
    operation: 'research_build',
    documentId: 'doc',
    projectId: 'research-project',
    ledgerId: 'ledger',
    expectedRevision: 0,
    draft: researchDraft(id),
  }
  const build = service(request, controller.signal)
  await started
  const read = await service(
    {
      operation: 'research_read',
      documentId: 'doc',
      projectId: 'research-project',
      ledgerId: 'ledger',
    },
    new AbortController().signal,
  )
  expect(parsePresentationResearchRecord(read).state).toBe('running')
  const duplicate = await service(request, new AbortController().signal)
  expect(record(duplicate).state).toBe('running')
  expect(attachments).toHaveBeenCalledTimes(2)
  controller.abort()
  release()
  expect(record(await build)).toMatchObject({ state: 'failed', error: 'aborted' })
  expect(record(await service(request, new AbortController().signal)).state).toBe('failed')
  expect(attachments).toHaveBeenCalledTimes(2)
})
it('exposes capabilities/list through actual presentation service without requiring a project', async () => {
  const f = setup(),
    pc = createPresentationService({ userDataPath: f.root })
  const decode = (v: Uint8Array) => JSON.parse(Buffer.from(v).toString())
  expect(
    decode(
      await pc(
        { operation: 'research_capabilities', documentId: 'doc' },
        new AbortController().signal,
      ),
    ),
  ).toEqual({ version: 1, available: true })
  expect(
    parsePresentationResearchSummary(
      decode(
        await pc(
          { operation: 'research_list', documentId: 'doc', projectId: 'new-project' },
          new AbortController().signal,
        ),
      ),
    ).totalRecords,
  ).toBe(0)
})

it('records uploading sources as gaps and safely persists matcher corruption', async () => {
  const f = setup(),
    raw = Buffer.from('not uploaded yet'),
    id = hash(raw)
  await f.attachments(
    {
      operation: 'attachment_begin',
      documentId: 'doc',
      attachmentId: id,
      sha256: id,
      name: 'pending.txt',
      sizeBytes: raw.length,
    },
    new AbortController().signal,
  )
  const draft = researchDraft(id)
  expect(
    record(await f.call('build', { ledgerId: 'pending', expectedRevision: 0, draft })).sources,
  ).toMatchObject([{ status: 'not_ready', provenance: 'unavailable' }])
  const readyId = await upload(f.attachments, '收入增长仅是管理层预测。')
  const corrupted = createPresentationResearchService({
    userDataPath: f.root,
    attachments: async (body, signal) => {
      if (body.operation === 'attachment_match_excerpt')
        throw new Error('private stack with secrets')
      return f.attachments(body, signal)
    },
  })
  const result = await corrupted(
    {
      operation: 'research_build',
      documentId: 'doc',
      projectId: 'research-project',
      ledgerId: 'broken',
      expectedRevision: 2,
      draft: researchDraft(readyId),
    },
    new AbortController().signal,
  )
  expect(record(result)).toMatchObject({ state: 'failed', error: 'invalid_state' })
  expect(JSON.stringify(result)).not.toContain('secrets')
})
it('reads latest completed archived research through the independent pre-plan route', async () => {
  const f = setup(),
    pc = createPresentationService({ userDataPath: f.root }),
    call = async (body: Record<string, unknown>) =>
      JSON.parse(
        Buffer.from(
          await pc(
            { documentId: 'doc', projectId: 'research-project', ...body },
            new AbortController().signal,
          ),
        ).toString(),
      )
  expect(await call({ operation: 'research_latest' })).toEqual({ record: null })
  const result = await f.call('build', {
    ledgerId: 'first',
    expectedRevision: 0,
    draft: researchDraft(),
  })
  expect(await call({ operation: 'research_latest' })).toEqual({ record: record(result) })
  expect(parsePresentationResearchSummary(await f.call('list')).revision).toBe(2)
})
