import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it, vi } from 'vitest'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'
import { createOfficeHostRuntime } from '../../office-addin/src/agent/host-runtime'

it('returns a scoped frozen precheck through the real runtime and PC without Office writes or QA publication', async () => {
  const root = mkdtempSync(join(tmpdir(), 'wiswork-content-integration-'))
  const service = createPresentationService({ userDataPath: root })
  const signal = new AbortController().signal
  const plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck(),
    documentId = 'doc'
  const call = async (operation: string, extra: Record<string, unknown>) =>
    JSON.parse(
      Buffer.from(
        await service({ operation, documentId, projectId: deck.id, ...extra }, signal),
      ).toString(),
    )
  const rememberProject = vi.fn(async () => {}),
    writeQa = vi.fn(async () => {}),
    writeReceipt = vi.fn(async () => {})
  const runtime = createOfficeHostRuntime('powerpoint', {
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
  try {
    expect(await call('save_plan', { expectedRevision: 0, plan })).toMatchObject({ revision: 1 })
    expect(
      await call('production_begin', { requestId: 'frozen', planRevision: 1, deck }),
    ).toMatchObject({ status: 'pending' })
    plan.sources[0]!.excerpt = ''
    await call('save_plan', { expectedRevision: 1, plan })
    const result = await runtime.skill.executeTool({
      id: 'content',
      name: 'check_presentation_page_content',
      input: { project_id: deck.id, request_id: 'frozen', page_id: deck.slides[0]!.id },
    })
    expect(result.isError, result.output).not.toBe(true)
    const output = JSON.parse(result.output)
    expect(output).toMatchObject({
      requestId: 'frozen',
      planRevision: 1,
      report: {
        pageId: deck.slides[0]!.id,
        checks: { content: 'needs_review', sources: 'not_verified', host: 'not_checked' },
      },
    })
    expect(output.report.findings).not.toContainEqual(
      expect.objectContaining({ code: 'source_excerpt_missing' }),
    )
    expect(await call('production_status', { requestId: 'frozen' })).toMatchObject({
      status: 'pending',
      compiledCount: 0,
    })
    expect(runtime.proposals.pending()).toBeUndefined()
    expect(runtime.importProgress?.read()).toBeUndefined()
    expect(runtime.qa?.read()).toBeUndefined()
    expect(rememberProject).not.toHaveBeenCalled()
    expect(writeQa).not.toHaveBeenCalled()
    expect(writeReceipt).not.toHaveBeenCalled()
  } finally {
    runtime.dispose()
    rmSync(root, { recursive: true, force: true })
  }
})
