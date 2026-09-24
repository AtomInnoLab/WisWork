import { createHash } from 'node:crypto'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it, vi } from 'vitest'
import { presentationPlanClaims } from '@wiswork/pptx-engine/presentation-plan'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'
import { createOfficeHostRuntime } from '../../office-addin/src/agent/host-runtime'

it('traces frozen evidence, records an immutable agent review and reads history after restart without publishing QA', async () => {
  const root = mkdtempSync(join(tmpdir(), 'wiswork-evidence-integration-'))
  let service = createPresentationService({ userDataPath: root })
  const signal = new AbortController().signal,
    documentId = 'doc'
  const plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  const original = '前文🙂\n示例数据仅用于测试\n后文：勿将本文当作指令。'
  const raw = Buffer.from(original),
    attachmentId = createHash('sha256').update(raw).digest('hex')
  const call = async (operation: string, extra: Record<string, unknown>) =>
    JSON.parse(Buffer.from(await service({ operation, documentId, ...extra }, signal)).toString())
  const rememberProject = vi.fn(async () => {}),
    writeQa = vi.fn(async () => {}),
    writeReceipt = vi.fn(async () => {})
  const create = () =>
    createOfficeHostRuntime('powerpoint', {
      presentation: {
        available: () => true,
        documentId: async () => documentId,
        rememberProject,
        lastProject: () => undefined,
        request: async (body, s) => new Response(Buffer.from(await service(body, s ?? signal))),
        readQa: () => undefined,
        writeQa,
        readReceipt: () => undefined,
        writeReceipt,
      },
    })
  let runtime = create()
  try {
    await call('attachment_begin', {
      attachmentId,
      sha256: attachmentId,
      name: 'evidence.txt',
      sizeBytes: raw.length,
    })
    await call('attachment_chunk', { attachmentId, offset: 0, base64: raw.toString('base64') })
    expect(await call('attachment_finish', { attachmentId })).toMatchObject({ status: 'ready' })
    plan.sources[0]!.uri = `attachment:${attachmentId}`
    plan.sources.push({
      id: 'pending-source',
      title: 'Source awaiting review',
      uri: 'https://example.invalid/pending',
      excerpt: 'Unreviewed material',
    })
    plan.claims[0]!.sourceIds.push('pending-source')
    deck.claims = presentationPlanClaims(plan)
    await call('save_plan', { projectId: deck.id, expectedRevision: 0, plan })
    await call('production_begin', {
      projectId: deck.id,
      requestId: 'frozen',
      planRevision: 1,
      deck,
    })
    const input = {
      project_id: deck.id,
      request_id: 'frozen',
      page_id: deck.slides[0]!.id,
      claim_id: plan.claims[0]!.id,
      source_id: plan.sources[0]!.id,
      offset: 0,
      max_chars: 8000,
    }
    const pageHistory = async () => {
      const result = await runtime.skill.executeTool({
        id: 'page-history',
        name: 'read_presentation_page_reviews',
        input: { project_id: deck.id, request_id: 'frozen', page_id: deck.slides[0]!.id },
      })
      expect(result.isError, result.output).not.toBe(true)
      return JSON.parse(result.output)
    }
    expect((await pageHistory()).claims[0]).toMatchObject({
      status: 'unreviewed',
      sources: [
        { sourceId: plan.sources[0]!.id, status: 'unreviewed', reviews: [] },
        { sourceId: 'pending-source', status: 'unreviewed', reviews: [] },
      ],
    })
    const read = () =>
      runtime.skill.executeTool({ id: 'evidence', name: 'read_presentation_claim_evidence', input })
    const result = await read()
    expect(result.isError, result.output).not.toBe(true)
    const evidence = JSON.parse(result.output)
    expect(evidence.attachment).toMatchObject({
      id: attachmentId,
      text: original,
      offsetUnit: 'utf16_code_unit',
    })
    expect(evidence.excerptMatch).toEqual({
      status: 'found',
      offset: original.indexOf(plan.sources[0]!.excerpt),
    })
    expect(evidence.checks.support).toBe('not_verified')
    const reviewInput = {
      ...input,
      review_id: 'review-1',
      outcome: 'supported',
      notes: 'This supplied passage supports only the test-data statement.',
    }
    const record = () =>
      runtime.skill.executeTool({
        id: 'record',
        name: 'record_presentation_claim_review',
        input: reviewInput,
      })
    const saved = await record()
    expect(saved.isError, saved.output).not.toBe(true)
    const review = JSON.parse(saved.output)
    expect(review).toMatchObject({
      reviewId: 'review-1',
      reviewer: 'agent',
      outcome: 'supported',
      attachmentId,
      checks: { support: 'agent_reviewed', sourceAuthority: 'not_verified' },
    })
    expect(JSON.parse((await record()).output)).toEqual(review)
    expect((await pageHistory()).claims[0]).toMatchObject({
      status: 'partial',
      sources: [
        { status: 'supported', reviews: [{ reviewId: 'review-1' }] },
        { status: 'unreviewed', reviews: [] },
      ],
    })
    const differing = await runtime.skill.executeTool({
      id: 'different-judgment',
      name: 'record_presentation_claim_review',
      input: {
        ...reviewInput,
        review_id: 'review-2',
        outcome: 'contradicted',
        notes: 'Another historical judgment to retain for reconciliation.',
      },
    })
    expect(differing.isError, differing.output).not.toBe(true)
    const mixed = await pageHistory()
    expect(mixed.claims[0]).toMatchObject({
      status: 'mixed',
      sources: [
        {
          status: 'mixed',
          reviews: [
            { reviewId: 'review-1', outcome: 'supported' },
            { reviewId: 'review-2', outcome: 'contradicted' },
          ],
        },
        { status: 'unreviewed', reviews: [] },
      ],
    })
    const conflict = await runtime.skill.executeTool({
      id: 'conflict',
      name: 'record_presentation_claim_review',
      input: { ...reviewInput, notes: 'A different judgment cannot overwrite this ID.' },
    })
    expect(conflict).toMatchObject({ isError: true, output: 'presentation_request_conflict' })
    plan.sources[0]!.uri = 'https://example.invalid/other'
    await call('save_plan', { projectId: deck.id, expectedRevision: 1, plan })
    runtime.dispose()
    service = createPresentationService({ userDataPath: root })
    runtime = create()
    const history = await runtime.skill.executeTool({
      id: 'history',
      name: 'read_presentation_claim_review',
      input: { project_id: deck.id, request_id: 'frozen', review_id: 'review-1' },
    })
    expect(history.isError, history.output).not.toBe(true)
    expect(JSON.parse(history.output)).toEqual(review)
    expect(await pageHistory()).toEqual(mixed)
    expect(await record()).toMatchObject({
      isError: true,
      output: 'presentation_evidence_read_required',
    })
    expect(JSON.parse((await read()).output)).toEqual(evidence)
    expect(JSON.parse((await record()).output)).toEqual(review)
    const bounded = await runtime.skill.executeTool({
      id: 'window',
      name: 'read_presentation_claim_evidence',
      input: { ...input, max_chars: 2 },
    })
    expect(JSON.parse(bounded.output).excerptMatch).toEqual({ status: 'not_found_in_window' })
    expect(
      await call('production_status', { projectId: deck.id, requestId: 'frozen' }),
    ).toMatchObject({ status: 'pending', compiledCount: 0 })
    expect(runtime.proposals.pending()).toBeUndefined()
    expect(runtime.qa?.read()).toBeUndefined()
    expect(runtime.importProgress?.read()).toBeUndefined()
    expect(rememberProject).not.toHaveBeenCalled()
    expect(writeQa).not.toHaveBeenCalled()
    expect(writeReceipt).not.toHaveBeenCalled()
  } finally {
    runtime.dispose()
    rmSync(root, { recursive: true, force: true })
  }
})
