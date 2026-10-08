import { PresentationStore } from '@wiswork/project-store'
import { handlePresentationProduction } from '../src/main/presentation-production'
import { createHash } from 'node:crypto'
import { mkdtempSync, readdirSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, expect, it, vi } from 'vitest'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { presentationPlanClaims } from '@wiswork/pptx-engine/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const decode = (v: Uint8Array) => JSON.parse(Buffer.from(v).toString())
const files = (root: string) =>
  readdirSync(root, { recursive: true, withFileTypes: true })
    .filter((e) => e.isFile())
    .map((e) => [
      join(e.parentPath, e.name),
      readFileSync(join(e.parentPath, e.name)).toString('base64'),
    ])
async function setup(uri?: string, snapshot = false) {
  const userDataPath = mkdtempSync(join(tmpdir(), 'wiswork-evidence-'))
  roots.push(userDataPath)
  const compile = vi.fn()
  const service = createPresentationService({ userDataPath, compile })
  const plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  const raw = Buffer.from('😀before 原文 after\f')
  const attachmentId = createHash('sha256').update(raw).digest('hex')
  plan.sources.push({ ...plan.sources[0]!, id: 'unrelated-source' })
  plan.claims.push({ ...plan.claims[0]!, id: 'unrelated-claim' })
  plan.sources[0]!.uri = uri ?? `attachment:${attachmentId}`
  if (snapshot) plan.sources[0]!.snapshotAttachmentId = attachmentId
  plan.sources[0]!.excerpt = '原文'
  deck.claims = presentationPlanClaims(plan)
  const call = async (operation: string, extra: Record<string, unknown> = {}) =>
    decode(await service({ operation, documentId: 'doc', ...extra }, new AbortController().signal))
  expect(
    await call('attachment_begin', {
      attachmentId,
      sha256: attachmentId,
      name: 'evidence.txt',
      sizeBytes: raw.length,
    }),
  ).not.toHaveProperty('error')
  expect(
    await call('attachment_chunk', { attachmentId, offset: 0, base64: raw.toString('base64') }),
  ).not.toHaveProperty('error')
  expect(await call('attachment_finish', { attachmentId })).not.toHaveProperty('error')
  await call('save_plan', { projectId: deck.id, expectedRevision: 0, plan })
  await call('production_begin', { projectId: deck.id, requestId: 'run', planRevision: 1, deck })
  const request = {
    operation: 'production_claim_evidence',
    documentId: 'doc',
    projectId: deck.id,
    requestId: 'run',
    pageId: deck.slides[0]!.id,
    claimId: 'source-1',
    sourceId: 'source',
    offset: 2,
    maxChars: 8000,
  }
  return { userDataPath, compile, service, call, plan, deck, request }
}
it('reads a document-bound snapshot while preserving the original URL in frozen evidence', async () => {
  const originalUri = 'https://example.com/research'
  const f = await setup(originalUri, true)
  const result = await f.call('production_claim_evidence', f.request)
  expect(result).toMatchObject({
    source: { uri: originalUri, snapshotAttachmentId: f.plan.sources[0]!.snapshotAttachmentId },
    attachment: { text: 'before 原文 after\f', provenance: { binding: 'user_supplied' } },
    excerptMatch: { status: 'found' },
  })
  expect(f.compile).not.toHaveBeenCalled()
})
it('rejects a fetched snapshot with a different plan URL and records the matched fetch time', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'wiswork-fetched-evidence-'))
  roots.push(userDataPath)
  const service = createPresentationService({
    userDataPath,
    fetchPage: async () =>
      new Response('<html><body><p>Original finding</p></body></html>', {
        headers: { 'content-type': 'text/html' },
      }),
  })
  const call = async (body: Record<string, unknown>) =>
    decode(await service({ documentId: 'doc', ...body }, new AbortController().signal))
  const url = 'https://8.8.8.8/research?revision=1'
  const snapshot = await call({ operation: 'attachment_import_webpage', url })
  const plan = benchmarkPlan()
  plan.sources[0]!.uri = 'https://8.8.8.8/research?revision=2'
  plan.sources[0]!.snapshotAttachmentId = snapshot.attachmentId
  plan.sources[0]!.excerpt = 'Original finding'
  plan.sources[0]!.locator = '第 1 段'
  const deck = benchmarkPlannedDeck()
  deck.claims = presentationPlanClaims(plan)
  await call({ operation: 'save_plan', projectId: deck.id, expectedRevision: 0, plan })
  await call({
    operation: 'production_begin',
    projectId: deck.id,
    requestId: 'wrong',
    planRevision: 1,
    deck,
  })
  const evidence = (requestId: string) =>
    call({
      operation: 'production_claim_evidence',
      projectId: deck.id,
      requestId,
      pageId: deck.slides[0]!.id,
      claimId: 'source-1',
      sourceId: 'source',
      offset: 0,
      maxChars: 8000,
    })
  expect(await evidence('wrong')).toEqual({ error: 'evidence_source_mismatch' })
  plan.sources[0]!.uri = url
  deck.claims = presentationPlanClaims(plan)
  await call({ operation: 'save_plan', projectId: deck.id, expectedRevision: 1, plan })
  await call({
    operation: 'production_begin',
    projectId: deck.id,
    requestId: 'correct',
    planRevision: 2,
    deck,
  })
  expect(await evidence('correct')).toMatchObject({
    attachment: {
      provenance: { binding: 'fetched_url_matched', retrievedAt: snapshot.retrievedAt },
    },
    excerptMatch: { status: 'found', locator: '第 1 段' },
  })
})
it('reads frozen attachment evidence before compilation without disk writes, including restart and exact windows', async () => {
  const f = await setup()
  const before = files(f.userDataPath)
  const read = (request = f.request) =>
    f.service(request, new AbortController().signal).then(decode)
  const result = await read()
  expect(result).toMatchObject({
    version: 1,
    planRevision: 1,
    source: { excerpt: '原文' },
    attachment: { offset: 2, text: 'before 原文 after\f', offsetUnit: 'utf16_code_unit' },
    excerptMatch: { status: 'found', offset: 9 },
    checks: { support: 'not_verified' },
  })
  expect(files(f.userDataPath)).toEqual(before)
  expect(await read({ ...f.request, maxChars: 8 })).toMatchObject({
    excerptMatch: { status: 'not_found_in_window' },
  })
  f.plan.sources[0]!.excerpt = 'changed'
  await f.call('save_plan', { projectId: f.deck.id, expectedRevision: 1, plan: f.plan })
  const restarted = createPresentationService({ userDataPath: f.userDataPath, compile: f.compile })
  const saved = files(f.userDataPath)
  expect(decode(await restarted(f.request, new AbortController().signal))).toEqual(result)
  expect(files(f.userDataPath)).toEqual(saved)
  expect(f.compile).not.toHaveBeenCalled()
})
it('rejects incomplete, invalid, unrelated, foreign and cancelled requests without writes', async () => {
  const f = await setup()
  const before = files(f.userDataPath)
  for (const key of Object.keys(f.request)) {
    const request: Record<string, unknown> = { ...f.request }
    delete request[key]
    expect(decode(await f.service(request, new AbortController().signal))).toEqual({
      error: 'invalid_request',
    })
  }
  for (const [patch, error] of [
    [{ unknown: true }, 'invalid_request'],
    [{ offset: -1 }, 'invalid_request'],
    [{ offset: 1.5 }, 'invalid_request'],
    [{ maxChars: 8001 }, 'invalid_request'],
    [{ claimId: 'missing' }, 'not_found'],
    [{ sourceId: 'missing' }, 'not_found'],
    [{ sourceId: 'unrelated-source' }, 'not_found'],
    [{ claimId: 'unrelated-claim' }, 'not_found'],
    [{ pageId: 'missing' }, 'not_found'],
    [{ documentId: 'foreign' }, 'document_mismatch'],
  ] as const)
    expect(
      decode(await f.service({ ...f.request, ...patch }, new AbortController().signal)),
    ).toEqual({ error })
  const c = new AbortController()
  c.abort()
  expect(decode(await f.service(f.request, c.signal))).toEqual({ error: 'aborted' })
  expect(files(f.userDataPath)).toEqual(before)
})
it.each(['https://example.com', 'file:///tmp/x', `attachment:${'A'.repeat(64)}`])(
  'rejects unsupported URI %s',
  async (uri) => {
    const f = await setup(uri)
    expect(decode(await f.service(f.request, new AbortController().signal))).toEqual({
      error: 'evidence_source_unsupported',
    })
  },
)

it('does not create attachment directories while reading a source uploaded only in another document', async () => {
  const f = await setup()
  expect(
    await f.call('save_plan', {
      documentId: 'other',
      projectId: 'other-project',
      expectedRevision: 0,
      plan: { ...f.plan, projectId: 'other-project' },
    }),
  ).not.toHaveProperty('error')
  expect(
    await f.call('production_begin', {
      documentId: 'other',
      projectId: 'other-project',
      requestId: 'run',
      planRevision: 1,
      deck: { ...f.deck, id: 'other-project' },
    }),
  ).not.toHaveProperty('error')
  const before = readdirSync(f.userDataPath, { recursive: true })
  expect(
    decode(
      await f.service(
        { ...f.request, documentId: 'other', projectId: 'other-project' },
        new AbortController().signal,
      ),
    ),
  ).toEqual({ error: 'not_found' })
  expect(readdirSync(f.userDataPath, { recursive: true })).toEqual(before)
})

it('honors cancellation after the attachment read completes without publishing or writing', async () => {
  const f = await setup(),
    controller = new AbortController()
  const before = files(f.userDataPath)
  const attachments = vi.fn(async (body: Record<string, unknown>) => {
    if (body.operation === 'attachment_metadata')
      return { attachmentId: f.plan.sources[0]!.uri.slice(11), status: 'ready', kind: 'text' }
    controller.abort()
    return {}
  })
  await expect(
    handlePresentationProduction(
      f.request,
      { store: new PresentationStore(f.userDataPath), compile: f.compile, attachments },
      controller.signal,
    ),
  ).rejects.toThrow('aborted')
  expect(attachments).toHaveBeenCalledWith(
    {
      operation: 'attachment_read',
      documentId: 'doc',
      attachmentId: f.plan.sources[0]!.uri.slice(11),
      offset: 2,
      maxChars: 8000,
    },
    controller.signal,
  )
  expect(files(f.userDataPath)).toEqual(before)
  expect(f.compile).not.toHaveBeenCalled()
})
