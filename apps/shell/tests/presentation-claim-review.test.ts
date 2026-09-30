import { PresentationStore } from '@wiswork/project-store'
import { handlePresentationProduction } from '../src/main/presentation-production'
import { presentationClaimEvidenceContent } from '@wiswork/pptx-engine/presentation-claim-review'
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
import {
  buildImageBackedPdfFixture,
  buildPdfFixture,
} from '../../../packages/file-parse/tests/helpers/fixtures'
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
async function setup(
  uri?: string,
  file?: { raw: Buffer; name: string; excerpt: string; locator: string },
) {
  const userDataPath = mkdtempSync(join(tmpdir(), 'wiswork-evidence-'))
  roots.push(userDataPath)
  const compile = vi.fn()
  const service = createPresentationService({ userDataPath, compile })
  const plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  const raw = file?.raw ?? Buffer.from('😀before 原文 after\f')
  const attachmentId = createHash('sha256').update(raw).digest('hex')
  plan.sources.push({ ...plan.sources[0]!, id: 'unrelated-source' })
  plan.claims.push({ ...plan.claims[0]!, id: 'unrelated-claim' })
  plan.sources[0]!.uri = uri ?? `attachment:${attachmentId}`
  plan.sources[0]!.excerpt = file?.excerpt ?? '原文'
  if (file) plan.sources[0]!.locator = file.locator
  deck.claims = presentationPlanClaims(plan)
  const call = async (operation: string, extra: Record<string, unknown> = {}) =>
    decode(await service({ operation, documentId: 'doc', ...extra }, new AbortController().signal))
  expect(
    await call('attachment_begin', {
      attachmentId,
      sha256: attachmentId,
      name: file?.name ?? 'evidence.txt',
      sizeBytes: raw.length,
    }),
  ).not.toHaveProperty('error')
  for (let offset = 0; offset < raw.length; offset += 128 * 1024)
    expect(
      await call('attachment_chunk', {
        attachmentId,
        offset,
        base64: raw.subarray(offset, offset + 128 * 1024).toString('base64'),
      }),
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
it('records agent judgments, retries immutably, and reads history without an attachment', async () => {
  const f = await setup()
  const evidence = decode(await f.service(f.request, new AbortController().signal))
  const request = {
    ...f.request,
    operation: 'production_record_claim_review',
    reviewId: 'review',
    evidenceDigest: createHash('sha256')
      .update(presentationClaimEvidenceContent(evidence))
      .digest('hex'),
    outcome: 'supported',
    notes: 'Original text supports this claim',
  }
  const write = (patch = {}) =>
    f.service({ ...request, ...patch }, new AbortController().signal).then(decode)
  const result = await write()
  expect(result).toMatchObject({
    reviewId: 'review',
    reviewer: 'agent',
    outcome: 'supported',
    notes: request.notes,
    checks: {
      support: 'agent_reviewed',
      sourceAuthority: 'not_verified',
      timeliness: 'not_verified',
      host: 'not_checked',
    },
  })
  expect(await write()).toEqual(result)
  expect(await write({ notes: 'changed' })).toEqual({ error: 'request_conflict' })
  const before = files(f.userDataPath)
  const readRequest = {
    operation: 'production_read_claim_review',
    documentId: 'doc',
    projectId: request.projectId,
    requestId: 'run',
    reviewId: 'review',
  }
  const attachments = vi.fn(async () => {
    throw new Error('not_found')
  })
  expect(
    await handlePresentationProduction(
      readRequest,
      { store: new PresentationStore(f.userDataPath), compile: f.compile, attachments },
      new AbortController().signal,
    ),
  ).toEqual(result)
  expect(attachments).not.toHaveBeenCalled()
  expect(files(f.userDataPath)).toEqual(before)
  expect(
    decode(
      await createPresentationService({ userDataPath: f.userDataPath })(
        readRequest,
        new AbortController().signal,
      ),
    ),
  ).toEqual(result)
})
it('requires a literal excerpt in the reviewed window before saving supported', async () => {
  const f = await setup()
  const request = { ...f.request, offset: 2, maxChars: 4 }
  const evidence = decode(await f.service(request, new AbortController().signal))
  expect(evidence.excerptMatch.status).toBe('not_found_in_window')
  const body = {
    ...request,
    operation: 'production_record_claim_review',
    reviewId: 'missing-excerpt',
    evidenceDigest: createHash('sha256')
      .update(presentationClaimEvidenceContent(evidence))
      .digest('hex'),
    outcome: 'supported',
    notes: 'I think this supports the claim',
  }
  expect(decode(await f.service(body, new AbortController().signal))).toEqual({
    error: 'evidence_excerpt_not_found',
  })
  expect(
    decode(
      await f.service({ ...body, outcome: 'insufficient_evidence' }, new AbortController().signal),
    ),
  ).toMatchObject({
    outcome: 'insufficient_evidence',
  })
})
it('binds a supported PDF review to the planned page locator', async () => {
  const f = await setup(undefined, {
    raw: Buffer.from(buildPdfFixture(['First page', 'Target evidence'])),
    name: 'evidence.pdf',
    excerpt: 'Target evidence',
    locator: '第 1 页',
  })
  const request = { ...f.request, offset: 0 }
  const evidence = decode(await f.service(request, new AbortController().signal))
  expect(evidence.excerptMatch).toMatchObject({ status: 'found', locator: '第 2 页' })
  const body = {
    ...request,
    operation: 'production_record_claim_review',
    reviewId: 'wrong-page',
    evidenceDigest: createHash('sha256')
      .update(presentationClaimEvidenceContent(evidence))
      .digest('hex'),
    outcome: 'supported',
    notes: 'Evidence on the second page',
  }
  expect(decode(await f.service(body, new AbortController().signal))).toEqual({
    error: 'evidence_locator_mismatch',
  })
})
it('allows supported review for a matching ordinary text PDF page', async () => {
  const f = await setup(undefined, {
    raw: Buffer.from(buildPdfFixture(['First page', 'Target evidence'])),
    name: 'evidence.pdf',
    excerpt: 'Target evidence',
    locator: '第 2 页',
  })
  const request = { ...f.request, offset: 0 }
  const evidence = decode(await f.service(request, new AbortController().signal))
  expect(evidence.excerptMatch).toMatchObject({ status: 'found', locator: '第 2 页' })
  expect(evidence.attachment.locatorSpans).not.toEqual(
    expect.arrayContaining([expect.objectContaining({ imageBacked: true })]),
  )
  expect(
    decode(
      await f.service(
        {
          ...request,
          operation: 'production_record_claim_review',
          reviewId: 'ordinary-pdf',
          evidenceDigest: createHash('sha256')
            .update(presentationClaimEvidenceContent(evidence))
            .digest('hex'),
          outcome: 'supported',
          notes: 'The cited page contains the literal source excerpt.',
        },
        new AbortController().signal,
      ),
    ),
  ).toMatchObject({ outcome: 'supported' })
})
it('allows visible native text over a full-page PDF image after literal review', async () => {
  const f = await setup(undefined, {
    raw: Buffer.from(buildImageBackedPdfFixture('Visible evidence')),
    name: 'image-background.pdf',
    excerpt: 'Visible evidence',
    locator: '第 1 页',
  })
  const request = { ...f.request, offset: 0 }
  const evidence = decode(await f.service(request, new AbortController().signal))
  expect(evidence.attachment.locatorSpans).toEqual(
    expect.arrayContaining([expect.objectContaining({ imageBacked: true })]),
  )
  expect(evidence.attachment.locatorSpans[0].invisibleTextLayer).toBeUndefined()
  expect(
    decode(
      await f.service(
        {
          ...request,
          operation: 'production_record_claim_review',
          reviewId: 'visible-over-image',
          evidenceDigest: createHash('sha256')
            .update(presentationClaimEvidenceContent(evidence))
            .digest('hex'),
          outcome: 'supported',
          notes: 'The visible text contains the cited excerpt.',
        },
        new AbortController().signal,
      ),
    ),
  ).toMatchObject({ outcome: 'supported' })
})
it('does not mark unverified image-backed scan text as supported', async () => {
  const raw = readFileSync(
    new URL(
      '../../../docs/product/ppt-benchmark-materials/PPT-P0-11/naca-rm-l50b01-1950-real-scan.pdf',
      import.meta.url,
    ),
  )
  const f = await setup(undefined, {
    raw,
    name: 'naca-scan.pdf',
    excerpt: 'NATIONAL ADVISORY COMMITTEE FOR AERONAUTICS',
    locator: '第 3 页',
  })
  const request = { ...f.request, offset: 0, maxChars: 8000 }
  const evidence = decode(await f.service(request, new AbortController().signal))
  expect(evidence.excerptMatch).toMatchObject({ status: 'found', locator: '第 3 页' })
  expect(evidence.attachment.locatorSpans).toEqual(
    expect.arrayContaining([
      expect.objectContaining({
        locator: '第 3 页',
        imageBacked: true,
        invisibleTextLayer: true,
      }),
    ]),
  )
  const body = {
    ...request,
    operation: 'production_record_claim_review',
    reviewId: 'scan-review',
    evidenceDigest: createHash('sha256')
      .update(presentationClaimEvidenceContent(evidence))
      .digest('hex'),
    outcome: 'supported',
    notes: 'OCR text alone is not sufficient',
  }
  expect(decode(await f.service(body, new AbortController().signal))).toEqual({
    error: 'evidence_image_backed_unverified',
  })
  expect(
    decode(
      await f.service({ ...body, outcome: 'insufficient_evidence' }, new AbortController().signal),
    ),
  ).toMatchObject({ outcome: 'insufficient_evidence' })

  const auxiliary = Buffer.from('NATIONAL ADVISORY COMMITTEE FOR AERONAUTICS')
  const auxiliaryId = createHash('sha256').update(auxiliary).digest('hex')
  expect(
    await f.call('attachment_begin', {
      attachmentId: auxiliaryId,
      sha256: auxiliaryId,
      name: 'user-supplied-text.txt',
      sizeBytes: auxiliary.length,
    }),
  ).not.toHaveProperty('error')
  expect(
    await f.call('attachment_chunk', {
      attachmentId: auxiliaryId,
      offset: 0,
      base64: auxiliary.toString('base64'),
    }),
  ).not.toHaveProperty('error')
  expect(await f.call('attachment_finish', { attachmentId: auxiliaryId })).toMatchObject({
    status: 'ready',
  })
  f.plan.sources[0]!.uri = `attachment:${auxiliaryId}`
  delete f.plan.sources[0]!.locator
  expect(
    await f.call('save_plan', {
      projectId: f.deck.id,
      expectedRevision: 1,
      plan: f.plan,
    }),
  ).toMatchObject({ revision: 2 })
  f.deck.claims = presentationPlanClaims(f.plan)
  expect(
    await f.call('production_begin', {
      projectId: f.deck.id,
      requestId: 'run-with-auxiliary',
      planRevision: 2,
      deck: f.deck,
    }),
  ).not.toHaveProperty('error')
  const retry = { ...request, requestId: 'run-with-auxiliary' }
  const supplemented = decode(await f.service(retry, new AbortController().signal))
  expect(supplemented.attachment).toMatchObject({ id: auxiliaryId, text: auxiliary.toString() })
  expect(supplemented.attachment.locatorSpans).toBeUndefined()
  expect(supplemented.excerptMatch).toMatchObject({ status: 'found' })
  expect(
    decode(
      await f.service(
        {
          ...retry,
          operation: 'production_record_claim_review',
          reviewId: 'supplemented-review',
          evidenceDigest: createHash('sha256')
            .update(presentationClaimEvidenceContent(supplemented))
            .digest('hex'),
          outcome: 'supported',
          notes:
            'Test-supplied text contains the literal excerpt; content is not independently verified.',
        },
        new AbortController().signal,
      ),
    ),
  ).toMatchObject({ outcome: 'supported', requestId: 'run-with-auxiliary' })
  expect(
    await f.call('production_read_claim_review', {
      projectId: f.deck.id,
      requestId: 'run',
      reviewId: 'scan-review',
    }),
  ).toMatchObject({ outcome: 'insufficient_evidence', requestId: 'run' })
  expect(
    decode(
      await f.service(
        { ...body, reviewId: 'old-run-after-auxiliary' },
        new AbortController().signal,
      ),
    ),
  ).toEqual({ error: 'evidence_image_backed_unverified' })
})
it('blocks an unsampled scan page in a long PDF after attachment upload', async () => {
  const pages = Array.from({ length: 65 }, (_, index) => `Page ${index + 1} text`)
  pages[1] = 'OCR evidence on page two'
  const f = await setup(undefined, {
    raw: Buffer.from(buildPdfFixture(pages, { imagePage: 2, invisibleText: true })),
    name: 'mixed-long.pdf',
    excerpt: pages[1],
    locator: '第 2 页',
  })
  const request = { ...f.request, offset: 0, maxChars: 8000 }
  const evidence = decode(await f.service(request, new AbortController().signal))
  expect(evidence.excerptMatch).toMatchObject({ status: 'found', locator: '第 2 页' })
  expect(evidence.attachment.locatorSpans).toEqual(
    expect.arrayContaining([
      expect.objectContaining({
        locator: '第 2 页',
        imageBacked: true,
        invisibleTextLayer: true,
      }),
    ]),
  )
  expect(
    decode(
      await f.service(
        {
          ...request,
          operation: 'production_record_claim_review',
          reviewId: 'long-pdf-scan-review',
          evidenceDigest: createHash('sha256')
            .update(presentationClaimEvidenceContent(evidence))
            .digest('hex'),
          outcome: 'supported',
          notes: 'Unverified OCR cannot support this claim',
        },
        new AbortController().signal,
      ),
    ),
  ).toEqual({ error: 'evidence_image_backed_unverified' })
}, 20_000)
it('rejects changed evidence, invalid requests and cancellation without writing', async () => {
  const f = await setup()
  const evidence = decode(await f.service(f.request, new AbortController().signal))
  const request = {
    ...f.request,
    operation: 'production_record_claim_review',
    reviewId: 'review',
    evidenceDigest: createHash('sha256')
      .update(presentationClaimEvidenceContent(evidence))
      .digest('hex'),
    outcome: 'insufficient_evidence',
    notes: 'Cannot establish support',
  }
  const before = files(f.userDataPath)
  for (const [patch, error] of [
    [{ evidenceDigest: 'b'.repeat(64) }, 'evidence_changed'],
    [{ notes: ' ' }, 'invalid_request'],
    [{ reviewer: 'human' }, 'invalid_request'],
    [{ outcome: 'verified' }, 'invalid_request'],
    [{ documentId: 'foreign' }, 'document_mismatch'],
  ] as const)
    expect(decode(await f.service({ ...request, ...patch }, new AbortController().signal))).toEqual(
      { error },
    )
  for (const key of Object.keys(request)) {
    const missing: Record<string, unknown> = { ...request }
    delete missing[key]
    expect(decode(await f.service(missing, new AbortController().signal))).toEqual({
      error: 'invalid_request',
    })
  }
  const changedAttachments = vi.fn(async (body: Record<string, unknown>) =>
    body.operation === 'attachment_metadata'
      ? { attachmentId: evidence.attachment.id, status: 'ready', kind: 'text' }
      : {
          attachmentId: evidence.attachment.id,
          sourceUri: evidence.source.uri,
          name: evidence.attachment.name,
          offset: evidence.attachment.offset,
          totalChars: evidence.attachment.totalChars,
          text: evidence.attachment.text.replace('after', 'other'),
        },
  )
  await expect(
    handlePresentationProduction(
      request,
      {
        store: new PresentationStore(f.userDataPath),
        compile: f.compile,
        attachments: changedAttachments,
      },
      new AbortController().signal,
    ),
  ).rejects.toThrow('evidence_changed')
  const controller = new AbortController()
  const attachments = vi.fn(async () => {
    controller.abort()
    return {}
  })
  await expect(
    handlePresentationProduction(
      request,
      { store: new PresentationStore(f.userDataPath), compile: f.compile, attachments },
      controller.signal,
    ),
  ).rejects.toThrow('aborted')
  expect(files(f.userDataPath)).toEqual(before)
})
