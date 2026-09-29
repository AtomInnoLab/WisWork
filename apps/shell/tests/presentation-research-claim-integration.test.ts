import { createHash } from 'node:crypto'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, expect, it, vi } from 'vitest'
import { presentationPlanClaims } from '@wiswork/pptx-engine/presentation-plan'
import { presentationClaimEvidenceContent } from '@wiswork/pptx-engine/presentation-claim-review'
import type { PresentationResearchDraft } from '@wiswork/project-store/presentation-research'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'
import { createOfficeHostRuntime } from '../../office-addin/src/agent/host-runtime'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const hash = (value: string | Uint8Array) => createHash('sha256').update(value).digest('hex')
async function setup(large = false) {
  const root = mkdtempSync(join(tmpdir(), 'research-claim-cross-'))
  roots.push(root)
  const compile = vi.fn()
  let service = createPresentationService({ userDataPath: root, compile })
  let corruptEvidence = false
  const documentId = 'doc',
    plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  const request = async (body: unknown, signal?: AbortSignal) => {
    const raw = await service(body, signal ?? new AbortController().signal)
    const value = JSON.parse(Buffer.from(raw).toString())
    if (
      corruptEvidence &&
      (body as { operation: string }).operation === 'production_claim_evidence' &&
      value.research
    )
      value.research.record.draft.scope = '伪造研究范围'
    return Response.json(value)
  }
  const runtime = () =>
    createOfficeHostRuntime('powerpoint', {
      presentation: {
        available: () => true,
        documentId: async () => documentId,
        lastProject: () => plan.projectId,
        rememberProject: async () => {},
        request,
      },
    })
  let host = runtime()
  const tool = (name: string, input: Record<string, unknown>) =>
    host.skill.executeTool({ id: name, name, input })
  const raw = async (operation: string, input: Record<string, unknown> = {}) =>
    JSON.parse(
      await (
        await request({
          operation,
          documentId,
          ...(operation.startsWith('attachment_') ? {} : { projectId: plan.projectId }),
          ...input,
        })
      ).text(),
    )
  const original = Buffer.from('原文对测试样本有效，不能推广到总体。'),
    attachmentId = hash(original)
  await raw('attachment_begin', {
    attachmentId,
    sha256: attachmentId,
    name: 'original.txt',
    sizeBytes: original.length,
  })
  await raw('attachment_chunk', { attachmentId, offset: 0, base64: original.toString('base64') })
  expect(await raw('attachment_finish', { attachmentId })).toMatchObject({ status: 'ready' })
  const draft: PresentationResearchDraft = {
    scope: '仅测试样本，非总体预测',
    sources: [
      {
        id: 'original-source',
        title: '原始资料',
        uri: `attachment:${attachmentId}`,
        excerpt: '原文对测试样本有效',
        locator: '第一段',
        asOf: '2026-09-01',
      },
      {
        id: 'unselected-source',
        title: '尚未取得原文',
        uri: 'https://example.invalid/secondary',
        excerpt: '另一口径',
        asOf: '2026-08-01',
      },
    ],
    facts: [
      {
        claimId: 'original-fact',
        statement: '原文对测试样本有效',
        type: 'fact',
        sourceRefs: ['original-source', 'unselected-source'],
        sourceTier: 'primary',
        slideRefs: [],
        confidence: 'medium',
        reviewStatus: 'needs_review',
        conflictsWith: ['opposing-judgment'],
        asOf: '2026-09-01',
        jurisdiction: '中国大陆；仅测试样本',
      },
      {
        claimId: 'opposing-judgment',
        statement: '不能据此推广到总体',
        type: 'judgment',
        sourceRefs: ['original-source'],
        sourceTier: 'unverified',
        slideRefs: [],
        confidence: 'low',
        reviewStatus: 'needs_review',
        conflictsWith: ['original-fact'],
        asOf: '2026-09-01',
        jurisdiction: '中国大陆；样本限制',
      },
    ],
  }
  if (large) {
    for (let index = 0; index < 22; index++)
      draft.sources.push({
        id: `additional-${index}`,
        title: '额外原研究上下文',
        uri: `https://example.invalid/original-${index}`,
        excerpt: '',
      })
    const target = 256 * 1024 - 1000
    for (const source of draft.sources.slice(2)) {
      const available = target - Buffer.byteLength(JSON.stringify(draft))
      source.excerpt = 'x'.repeat(Math.min(12000, Math.max(0, available)))
    }
    expect(Buffer.byteLength(JSON.stringify(draft))).toBe(target)
  }
  const built = await tool('build_research_ledger', {
    project_id: plan.projectId,
    ledger_id: 'research-a',
    expected_revision: 0,
    draft,
  })
  expect(built.isError, built.output).toBeFalsy()
  const recordA = JSON.parse(built.output).record
  plan.sources = [{ ...draft.sources[0]!, id: 'source' }]
  plan.claims = [
    {
      id: 'source-1',
      statement: draft.facts[0]!.statement,
      type: 'fact',
      sourceIds: ['source'],
      confidence: 'medium',
      reviewStatus: 'needs_review',
      asOf: draft.facts[0]!.asOf,
      jurisdiction: draft.facts[0]!.jurisdiction,
    },
  ]
  plan.research = {
    ledgerId: recordA.id,
    sequence: recordA.sequence,
    draftDigest: recordA.draftDigest,
    sources: [{ sourceId: 'source', researchSourceId: 'original-source' }],
    claims: [{ claimId: 'source-1', researchClaimId: 'original-fact' }],
  }
  deck.claims = presentationPlanClaims(plan)
  expect(await raw('save_plan', { expectedRevision: 0, plan })).not.toHaveProperty('error')
  expect(
    await raw('production_begin', { requestId: 'frozen', planRevision: 1, deck }),
  ).not.toHaveProperty('error')
  const input = {
    project_id: plan.projectId,
    request_id: 'frozen',
    page_id: deck.slides[0]!.id,
    claim_id: 'source-1',
    source_id: 'source',
    offset: 0,
    max_chars: 8000,
  }
  return {
    root,
    compile,
    plan,
    deck,
    draft,
    recordA,
    input,
    raw,
    tool,
    corrupt: () => {
      corruptEvidence = true
    },
    close: () => host.dispose(),
    restart: () => {
      host.dispose()
      service = createPresentationService({ userDataPath: root, compile })
      host = runtime()
    },
  }
}
it('reviews the actual frozen research context and persists page-scoped unresolved conflicts through restart', async () => {
  const f = await setup()
  try {
    const newer = structuredClone(f.draft)
    newer.facts[0]!.statement = '更新研究的不同陈述'
    const b = await f.tool('build_research_ledger', {
      project_id: f.plan.projectId,
      ledger_id: 'research-b',
      expected_revision: 2,
      draft: newer,
    })
    expect(b.isError, b.output).toBeFalsy()
    const recordB = JSON.parse(b.output).record
    const planB = structuredClone(f.plan)
    planB.claims[0]!.statement = newer.facts[0]!.statement
    planB.research = {
      ...planB.research!,
      ledgerId: recordB.id,
      sequence: recordB.sequence,
      draftDigest: recordB.draftDigest,
    }
    expect(await f.raw('save_plan', { expectedRevision: 1, plan: planB })).not.toHaveProperty(
      'error',
    )
    const read = await f.tool('read_presentation_claim_evidence', f.input)
    expect(read.isError, read.output).toBeFalsy()
    const evidence = JSON.parse(read.output)
    expect(evidence.claim).toEqual(f.plan.claims[0])
    expect(evidence.source.asOf).toBe('2026-09-01')
    expect(evidence.research.record).toEqual(f.recordA)
    expect(evidence.research.binding).toEqual(f.plan.research)
    expect(evidence.attachment.text).toContain(f.draft.sources[0]!.excerpt)
    const saved = await f.tool('record_presentation_claim_review', {
      ...f.input,
      review_id: 'supported-a',
      outcome: 'supported',
      notes: '原文支持本样本观察；不能推广到总体，原冲突及未选资料仍待复核。',
    })
    expect(saved.isError, saved.output).toBeFalsy()
    expect(JSON.parse(saved.output).evidenceDigest).toBe(
      hash(presentationClaimEvidenceContent(evidence)),
    )
    const reportResult = await f.tool('read_presentation_delivery_report', {
      project_id: f.plan.projectId,
      request_id: 'frozen',
    })
    expect(reportResult.isError, reportResult.output).toBeFalsy()
    let report = JSON.parse(reportResult.output)
    const conflict = report.pages[0].issues.find(
      (issue: { code: string }) => issue.code === 'research_claim_conflict',
    )
    expect(conflict).toMatchObject({
      category: 'needs_human',
      research: { ledgerId: 'research-a', relatedClaimIds: ['opposing-judgment'] },
      disposition: { state: 'open' },
    })
    expect(report.pages[0].issues.map((issue: { code: string }) => issue.code)).toEqual(
      expect.arrayContaining([
        'research_conflict_partner_omitted',
        'research_source_reference_unselected',
        'research_source_unavailable',
      ]),
    )
    const otherPage = report.pages
      .slice(1)
      .find((page: { issues: { code: string }[] }) =>
        page.issues.some((issue) => issue.code === 'research_claim_conflict'),
      )
    expect(otherPage).toBeDefined()
    const otherConflict = otherPage.issues.find(
      (issue: { code: string }) => issue.code === 'research_claim_conflict',
    )
    expect(otherConflict.id).not.toBe(conflict.id)
    for (const [index, state] of ['explained', 'deferred', 'open'].entries()) {
      const action = await f.tool('record_presentation_issue_action', {
        project_id: f.plan.projectId,
        request_id: 'frozen',
        expected_revision: index,
        action: {
          actionId: `action-${index}`,
          issueId: conflict.id,
          issueDigest: conflict.digest,
          state,
          note: '保留原冲突，尚未核验来源及总体推断。',
        },
      })
      expect(action.isError, action.output).toBeFalsy()
      report = JSON.parse(action.output)
      expect(
        report.pages[0].issues.find((issue: { id: string }) => issue.id === conflict.id).disposition
          .state,
      ).toBe(state)
      expect(
        report.pages
          .flatMap((page: { issues: unknown[] }) => page.issues)
          .find((issue: { id: string }) => issue.id === otherConflict.id).disposition.state,
      ).toBe('open')
    }
    f.restart()
    const restored = await f.tool('read_presentation_delivery_report', {
      project_id: f.plan.projectId,
      request_id: 'frozen',
    })
    expect(restored.isError, restored.output).toBeFalsy()
    expect(JSON.parse(restored.output).issueLedger.actions).toHaveLength(3)
    expect(JSON.parse(restored.output).checks.sourceAuthority).toBe('not_verified')
    expect(f.compile).not.toHaveBeenCalled()
  } finally {
    f.close()
  }
})
it('refuses a forged research draft returned with the original digest before enabling a new claim review', async () => {
  const f = await setup()
  try {
    f.corrupt()
    expect(await f.tool('read_presentation_claim_evidence', f.input)).toMatchObject({
      isError: true,
      output: 'presentation_response_invalid',
    })
    expect(
      (
        await f.tool('record_presentation_claim_review', {
          ...f.input,
          review_id: 'forged',
          outcome: 'supported',
          notes: '不能用伪造研究进行复核',
        })
      ).output,
    ).toBe('presentation_evidence_read_required')
    expect(
      await f.raw('production_read_claim_review', { requestId: 'frozen', reviewId: 'forged' }),
    ).toHaveProperty('error', 'not_found')
  } finally {
    f.close()
  }
})

it('reads the complete bounded research evidence above the old 256KiB response limit without truncating it', async () => {
  const f = await setup(true)
  try {
    const read = await f.tool('read_presentation_claim_evidence', f.input)
    expect(read.isError, read.output).toBeFalsy()
    expect(Buffer.byteLength(read.output)).toBeGreaterThan(256 * 1024)
    expect(Buffer.byteLength(read.output)).toBeLessThanOrEqual(512 * 1024)
    expect(JSON.parse(read.output).research.record).toEqual(f.recordA)
    expect(JSON.parse(read.output).research.record.draft.sources).toHaveLength(24)
    expect(f.compile).not.toHaveBeenCalled()
  } finally {
    f.close()
  }
})
