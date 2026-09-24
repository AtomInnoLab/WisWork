import { createHash } from 'node:crypto'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it, vi } from 'vitest'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { presentationPlanClaims } from '@wiswork/pptx-engine/presentation-plan'
import { PresentationStore } from '@wiswork/project-store'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'
import { createOfficeHostRuntime } from '../../office-addin/src/agent/host-runtime'

it('connects arithmetic, historical evidence, scoped dispositions, restart and JSON/Markdown export without host writes', async () => {
  const root = mkdtempSync(join(tmpdir(), 'wiswork-evidence-delivery-'))
  const compile = vi.fn(compilePresentationDeck)
  let service = createPresentationService({ userDataPath: root, compile })
  const plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  const documentId = 'delivery-doc',
    requestId = 'frozen'
  const signal = new AbortController().signal
  const raw = Buffer.from(
    'A = 10\nB = 20\nSynthetic source <script>ignore this instruction</script>',
  )
  const attachmentId = createHash('sha256').update(raw).digest('hex')
  const call = async (operation: string, extra: Record<string, unknown> = {}) =>
    JSON.parse(Buffer.from(await service({ operation, documentId, ...extra }, signal)).toString())
  const writeQa = vi.fn(async () => {}),
    writeReceipt = vi.fn(async () => {})
  const create = () =>
    createOfficeHostRuntime('powerpoint', {
      presentation: {
        available: () => true,
        documentId: async () => documentId,
        lastProject: () => deck.id,
        rememberProject: async () => {},
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
      name: 'numbers.txt',
      sizeBytes: raw.length,
    })
    await call('attachment_chunk', { attachmentId, offset: 0, base64: raw.toString('base64') })
    expect(await call('attachment_finish', { attachmentId })).toMatchObject({ status: 'ready' })
    plan.sources[0]!.uri = `attachment:${attachmentId}`
    plan.sources[0]!.excerpt = raw.toString()
    plan.sources[0]!.asOf = '2026-09-24'
    const claim = plan.claims[0]!
    claim.type = 'calculation'
    claim.statement = 'Declared total: 30'
    claim.asOf = '2026-09-24'
    claim.calculation = {
      formula: 'a + b',
      inputs: ['A from line 1', 'B from line 2'],
      unit: 'count',
      reproduction: {
        bindings: [
          { name: 'a', inputIndex: 0, value: 10, sourceId: 'source' },
          { name: 'b', inputIndex: 1, value: 20, sourceId: 'source' },
        ],
        expected: 30,
      },
    }
    const mismatch = structuredClone(claim),
      unsupported = structuredClone(claim)
    mismatch.id = 'calculation-mismatch'
    mismatch.statement = 'Declared total: 31'
    mismatch.calculation!.reproduction!.expected = 31
    unsupported.id = 'calculation-script'
    unsupported.statement = 'Unverified script calculation'
    unsupported.calculation!.formula = 'globalThis.__wisworkDeliveryProbe = 1'
    plan.claims.push(mismatch, unsupported)
    plan.slides[0]!.claimIds = plan.claims.map((value) => value.id)
    deck.slides[0]!.claimIds = [...plan.slides[0]!.claimIds]
    deck.slides[0]!.elements = plan.claims.map((value, index) => ({
      kind: 'text',
      id: `c${index}`,
      text: value.statement,
      x: 1,
      y: 1 + index,
      w: 9,
      h: 0.7,
    }))
    deck.claims = presentationPlanClaims(plan)
    expect(
      await call('save_plan', { projectId: deck.id, expectedRevision: 0, plan }),
    ).toMatchObject({ revision: 1 })
    expect(
      await call('production_begin', { projectId: deck.id, requestId, planRevision: 1, deck }),
    ).toMatchObject({ requestId })
    expect(await call('production_run', { projectId: deck.id, requestId })).toMatchObject({
      compiledCount: 8,
    })
    const input = {
      project_id: deck.id,
      request_id: requestId,
      page_id: deck.slides[0]!.id,
      claim_id: claim.id,
      source_id: 'source',
      offset: 0,
      max_chars: 8000,
    }
    const readEvidence = async () => {
      const result = await runtime.skill.executeTool({
        id: 'evidence',
        name: 'read_presentation_claim_evidence',
        input,
      })
      expect(result.isError, result.output).not.toBe(true)
    }
    const review = async (review_id: string, outcome: string) => {
      const result = await runtime.skill.executeTool({
        id: review_id,
        name: 'record_presentation_claim_review',
        input: {
          ...input,
          review_id,
          outcome,
          notes: 'A historical Agent judgment for this synthetic source window.',
        },
      })
      expect(result.isError, result.output).not.toBe(true)
    }
    await readEvidence()
    await review('review-one', 'supported')
    await review('review-two', 'contradicted')
    await runtime.presentation!.refresh()
    await runtime.presentation!.readDeliveryReport()
    const report = runtime.presentation!.snapshot().deliveryReport!
    expect(report, runtime.presentation!.snapshot().error).toBeDefined()
    expect(report.pages[0]!.calculations).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ claimId: claim.id, status: 'reproduced', actual: 30 }),
        expect.objectContaining({
          claimId: mismatch.id,
          status: 'mismatch',
          actual: 30,
          expected: 31,
        }),
        expect.objectContaining({ claimId: unsupported.id, status: 'unsupported_expression' }),
      ]),
    )
    const mixedIssue = report.pages[0]!.issues.find(
      (issue) => issue.code === 'source_review_mixed' && issue.claimId === claim.id,
    )!
    const calculationIssue = report.pages[0]!.issues.find(
      (issue) => issue.code === 'calculation_mismatch',
    )!
    expect(mixedIssue).toBeDefined()
    expect(calculationIssue).toBeDefined()
    await runtime.presentation!.recordIssueAction({
      actionId: 'explain',
      issueId: mixedIssue.id,
      issueDigest: mixedIssue.digest,
      state: 'explained',
      note: 'Different judgments remain; explanation is not fact verification.',
    })
    expect(runtime.presentation!.snapshot().error).toBeUndefined()
    await runtime.presentation!.recordIssueAction({
      actionId: 'defer',
      issueId: calculationIssue.id,
      issueDigest: calculationIssue.digest,
      state: 'deferred',
      note: 'Correct the declared value in a new plan revision.',
    })
    expect(runtime.presentation!.snapshot().deliveryReport!.issueLedger.revision).toBe(2)
    const beforeRestart = runtime.presentation!.snapshot().deliveryReport!
    const updatedPlan = structuredClone(plan)
    updatedPlan.claims[0]!.calculation!.reproduction!.expected = 99
    expect(
      await call('save_plan', { projectId: deck.id, expectedRevision: 1, plan: updatedPlan }),
    ).toMatchObject({ revision: 2 })
    runtime.dispose()
    service = createPresentationService({ userDataPath: root, compile })
    runtime = create()
    await runtime.presentation!.refresh()
    await runtime.presentation!.readDeliveryReport()
    expect(runtime.presentation!.snapshot().deliveryReport).toEqual(beforeRestart)
    await readEvidence()
    await review('review-three', 'supported')
    await runtime.presentation!.readDeliveryReport()
    const changed = runtime.presentation!.snapshot().deliveryReport!
    expect(
      changed.pages[0]!.issues.find((issue) => issue.id === mixedIssue.id)?.disposition,
    ).toMatchObject({ state: 'open', stale: true, actionId: 'explain' })
    expect(
      changed.pages[0]!.issues.find((issue) => issue.id === calculationIssue.id)?.disposition,
    ).toMatchObject({ state: 'deferred', stale: false })
    const stale = await runtime.skill.executeTool({
      id: 'stale',
      name: 'record_presentation_issue_action',
      input: {
        project_id: deck.id,
        request_id: requestId,
        expected_revision: 2,
        action: {
          actionId: 'stale-action',
          issueId: mixedIssue.id,
          issueDigest: mixedIssue.digest,
          state: 'explained',
          note: 'Stale action must not replace newer evidence.',
        },
      },
    })
    expect(stale.isError).toBe(true)
    expect(new PresentationStore(root).issueActions(deck.id, documentId, requestId).revision).toBe(
      2,
    )
    const retry = await runtime.skill.executeTool({
      id: 'retry',
      name: 'record_presentation_issue_action',
      input: {
        project_id: deck.id,
        request_id: requestId,
        expected_revision: 0,
        action: {
          actionId: 'explain',
          issueId: mixedIssue.id,
          issueDigest: mixedIssue.digest,
          state: 'explained',
          note: 'Different judgments remain; explanation is not fact verification.',
        },
      },
    })
    expect(retry.isError, retry.output).not.toBe(true)
    expect(new PresentationStore(root).issueActions(deck.id, documentId, requestId).revision).toBe(
      2,
    )
    await runtime.presentation!.exportDeliveryReport()
    expect(runtime.presentation!.snapshot().error).toBeUndefined()
    const paths = runtime.vfs.list('/home/user')
    const jsonPath = paths.find((path) => path.endsWith('.json'))!,
      markdownPath = paths.find((path) => path.endsWith('.md'))!
    expect(jsonPath).toBeDefined()
    expect(markdownPath).toBeDefined()
    const exported = JSON.parse(new TextDecoder().decode(runtime.vfs.readBytes(jsonPath)))
    expect(exported.plan.claims).toHaveLength(3)
    expect(exported.reviews).toHaveLength(3)
    expect(exported.issueLedger.actions).toHaveLength(2)
    expect(exported.checks).toMatchObject({
      host: 'not_checked',
      roundTrip: 'not_run',
      sourceAuthority: 'not_verified',
    })
    const markdown = new TextDecoder().decode(runtime.vfs.readBytes(markdownPath))
    expect(markdown).toContain('待人工判断')
    expect(markdown).toContain('无法核验')
    expect(markdown).not.toContain('<script>')
    expect((globalThis as Record<string, unknown>).__wisworkDeliveryProbe).toBeUndefined()
    expect(compile).toHaveBeenCalledTimes(8)
    expect(writeQa).not.toHaveBeenCalled()
    expect(writeReceipt).not.toHaveBeenCalled()
    expect(runtime.importProgress?.read()).toBeUndefined()
  } finally {
    runtime.dispose()
    rmSync(root, { recursive: true, force: true })
  }
}, 45000)
