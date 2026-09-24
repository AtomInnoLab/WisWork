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

it('traces a frozen claim to uploaded original text through runtime and restarts without publishing QA', async () => {
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
    plan.sources[0]!.uri = 'https://example.invalid/other'
    await call('save_plan', { projectId: deck.id, expectedRevision: 1, plan })
    runtime.dispose()
    service = createPresentationService({ userDataPath: root })
    runtime = create()
    expect(JSON.parse((await read()).output)).toEqual(evidence)
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
