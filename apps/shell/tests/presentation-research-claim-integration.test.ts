import type { PresentationProfessionalContext } from '@wiswork/project-store/presentation-professional-context'
import JSZip from 'jszip'
import { createPresentationHostBundleSkill } from '../../office-addin/src/skills/powerpoint/presentation-host-bundle'
import { InMemoryVfs } from '../../office-addin/src/skills/shared/vfs'
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
async function setup(
  large = false,
  professionalContext?: PresentationProfessionalContext,
  boundResearch = true,
) {
  const root = mkdtempSync(join(tmpdir(), 'research-claim-cross-'))
  roots.push(root)
  const compile = vi.fn()
  let service = createPresentationService({ userDataPath: root, compile })
  let corruptEvidence = false
  let hideAssessmentReply = false
  let omitProfessionalContext = false
  let stripProfessionalContext = false
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
    if (
      hideAssessmentReply &&
      (body as { operation: string }).operation === 'production_record_claim_review'
    ) {
      hideAssessmentReply = false
      delete value.sourceAssessment
    }
    if (
      omitProfessionalContext &&
      (body as { operation: string }).operation === 'production_claim_evidence' &&
      value.claim
    )
      delete value.claim.professionalContext
    if (
      stripProfessionalContext &&
      (body as { operation: string }).operation === 'production_claim_evidence'
    ) {
      delete value.claim
      delete value.documentId
      delete value.source.asOf
    }
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
  if (professionalContext)
    draft.facts[0]!.professionalContext = structuredClone(professionalContext)
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
      ...(professionalContext ? { professionalContext: structuredClone(professionalContext) } : {}),
    },
  ]
  plan.research = {
    ledgerId: recordA.id,
    sequence: recordA.sequence,
    draftDigest: recordA.draftDigest,
    sources: [{ sourceId: 'source', researchSourceId: 'original-source' }],
    claims: [{ claimId: 'source-1', researchClaimId: 'original-fact' }],
  }
  if (!boundResearch) delete plan.research
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
    request,
    hideAssessmentReply: () => {
      hideAssessmentReply = true
    },
    compile,
    plan,
    deck,
    draft,
    recordA,
    input,
    raw,
    tool,
    stripProfessionalContext: () => {
      stripProfessionalContext = true
    },
    omitProfessionalContext: () => {
      omitProfessionalContext = true
    },
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

function sourceAssessmentFixture() {
  return {
    scope: '仅对原文窗口与测试样本的来源适用性判断',
    authority: {
      outcome: 'appropriate_for_claim',
      sourceTier: 'primary',
      reason: '在本合成测试范围内判断；尚未认证真实出版方。',
    },
    timeliness: {
      outcome: 'current_for_claim',
      referenceDate: '2026-09-29',
      claimAsOf: '2026-09-01',
      sourceAsOf: '2026-09-01',
      reason: '原标签一致，仅判断本范围，不证明法规现行或来源更新。',
    },
    jurisdiction: {
      claimJurisdiction: '中国大陆；仅测试样本',
      outcome: 'applicable',
      reason: '仅限声明的样本范围，不能推广。',
    },
    basis: [{ offset: 0, text: '原文对测试样本有效' }],
  }
}
it('preserves every source authority/time/scope assessment and exposes mixed historical judgments without closing research conflicts', async () => {
  const f = await setup()
  try {
    const read = await f.tool('read_presentation_claim_evidence', f.input)
    expect(read.isError, read.output).toBeFalsy()
    const assessment = sourceAssessmentFixture()
    const input = {
      ...f.input,
      review_id: 'source-positive',
      outcome: 'supported',
      notes: '本窗口支持样本陈述；来源判断尚未认证。',
      source_assessment: assessment,
    }
    const saved = await f.tool('record_presentation_claim_review', input)
    expect(saved.isError, saved.output).toBeFalsy()
    expect(JSON.parse(saved.output).sourceAssessment).toEqual(assessment)
    const other = structuredClone(assessment)
    other.authority = {
      outcome: 'uncertain',
      sourceTier: 'secondary',
      reason: '不能从上传资料证明实际出版方。',
    }
    other.timeliness = {
      ...other.timeliness,
      outcome: 'historical_only',
      referenceDate: '2026-09-28',
      reason: '不同参照日的历史判断仍保留。',
    }
    other.jurisdiction = {
      ...other.jurisdiction,
      outcome: 'mismatch',
      reason: '原样本结论不能直接应用总体。',
    }
    const second = await f.tool('record_presentation_claim_review', {
      ...input,
      review_id: 'source-uncertain',
      source_assessment: other,
    })
    expect(second.isError, second.output).toBeFalsy()
    const result = await f.tool('read_presentation_delivery_report', {
      project_id: f.plan.projectId,
      request_id: 'frozen',
    })
    expect(result.isError, result.output).toBeFalsy()
    const report = JSON.parse(result.output)
    expect(
      report.reviews.map((review: { sourceAssessment: unknown }) => review.sourceAssessment),
    ).toEqual([assessment, other])
    expect(report.pages[0].issues.map((issue: { code: string }) => issue.code)).toEqual(
      expect.arrayContaining([
        'source_authority_review_mixed',
        'source_timeliness_review_mixed',
        'source_jurisdiction_review_mixed',
        'research_claim_conflict',
      ]),
    )
    const issue = report.pages[0].issues.find(
      (issue: { code: string }) => issue.code === 'source_authority_review_mixed',
    )
    const action = await f.tool('record_presentation_issue_action', {
      project_id: f.plan.projectId,
      request_id: 'frozen',
      expected_revision: 0,
      action: {
        actionId: 'explain-source',
        issueId: issue.id,
        issueDigest: issue.digest,
        state: 'explained',
        note: '保留两个不同历史判断，不能视为认证。',
      },
    })
    expect(action.isError, action.output).toBeFalsy()
    expect(JSON.parse(action.output).checks.sourceAuthority).toBe('not_verified')
    f.restart()
    const restored = await f.tool('read_presentation_claim_review', {
      project_id: f.plan.projectId,
      request_id: 'frozen',
      review_id: 'source-positive',
    })
    expect(restored.isError, restored.output).toBeFalsy()
    expect(JSON.parse(restored.output).sourceAssessment).toEqual(assessment)
    expect(f.compile).not.toHaveBeenCalled()
  } finally {
    f.close()
  }
})
it('rejects forged literal source-assessment basis before the PC persists a new review', async () => {
  const f = await setup()
  try {
    expect((await f.tool('read_presentation_claim_evidence', f.input)).isError).toBeFalsy()
    for (const [index, basis] of [
      [{ offset: 1, text: '原文对测试样本有效' }],
      [{ offset: 0, text: '伪造原文' }],
      [{ offset: 8001, text: '原文' }],
    ].entries()) {
      const source_assessment = { ...sourceAssessmentFixture(), basis }
      const result = await f.tool('record_presentation_claim_review', {
        ...f.input,
        review_id: `invalid-basis-${index}`,
        outcome: 'supported',
        notes: '不能用错误原文依据写判断',
        source_assessment,
      })
      expect(result.isError).toBe(true)
      expect(
        await f.raw('production_read_claim_review', {
          requestId: 'frozen',
          reviewId: `invalid-basis-${index}`,
        }),
      ).toHaveProperty('error', 'not_found')
    }
  } finally {
    f.close()
  }
})
it('does not claim an omitted assessment acknowledgment succeeded and restores the immutable real PC result by read', async () => {
  const f = await setup()
  try {
    expect((await f.tool('read_presentation_claim_evidence', f.input)).isError).toBeFalsy()
    f.hideAssessmentReply()
    const source_assessment = sourceAssessmentFixture()
    const result = await f.tool('record_presentation_claim_review', {
      ...f.input,
      review_id: 'lost-assessment-ack',
      outcome: 'supported',
      notes: '实际写入，返回字段丢失后只读恢复',
      source_assessment,
    })
    expect(result).toMatchObject({ isError: true, output: 'presentation_response_invalid' })
    f.restart()
    const read = await f.tool('read_presentation_claim_review', {
      project_id: f.plan.projectId,
      request_id: 'frozen',
      review_id: 'lost-assessment-ack',
    })
    expect(read.isError, read.output).toBeFalsy()
    expect(JSON.parse(read.output).sourceAssessment).toEqual(source_assessment)
  } finally {
    f.close()
  }
})
it('includes the actual immutable source assessment and unresolved source issues in the current host delivery archive', async () => {
  const f = await setup()
  try {
    expect((await f.tool('read_presentation_claim_evidence', f.input)).isError).toBeFalsy()
    const source_assessment = sourceAssessmentFixture()
    const saved = await f.tool('record_presentation_claim_review', {
      ...f.input,
      review_id: 'package-source',
      outcome: 'supported',
      notes: '只针对原窗口；不得宣称整套完成',
      source_assessment,
    })
    expect(saved.isError, saved.output).toBeFalsy()
    const native = await new JSZip()
      .file('ppt/slides/slide1.xml', '<title>current host</title>')
      .generateAsync({ type: 'uint8array' })
    const vfs = new InMemoryVfs()
    const skill = createPresentationHostBundleSkill({
      available: () => true,
      nativeAvailable: () => true,
      exportDocument: async () => native,
      documentId: async () => 'doc',
      request: f.request,
      vfs,
    })
    const exported = await skill.executeTool({
      id: 'assessed-package',
      name: 'export_current_presentation_bundle',
      input: { project_id: f.plan.projectId, request_id: 'frozen' },
    })
    expect(exported.isError, exported.output).toBeFalsy()
    const zip = await JSZip.loadAsync(vfs.readBytes(JSON.parse(exported.output).paths[0]))
    const report = JSON.parse(await zip.file('evidence.json')!.async('string'))
    expect(report.reviews[0].sourceAssessment).toEqual(source_assessment)
    expect(report.pages[0].issues.map((issue: { code: string }) => issue.code)).toContain(
      'research_claim_conflict',
    )
    expect(report.checks.sourceAuthority).toBe('not_verified')
    expect(await zip.file('evidence.md')!.async('string')).toContain('2026&#45;09&#45;29')
    expect(f.compile).not.toHaveBeenCalled()
  } finally {
    f.close()
  }
})

const professionalContexts: PresentationProfessionalContext[] = [
  {
    domain: 'science',
    materialKind: 'paper',
    publicationId: 'synthetic-paper-id',
    version: '1',
    sample: '仅测试样本',
    method: '合成材料测试',
    statisticalBasis: '样本描述，无总体推断',
    limitations: '不构成真实科研结论',
  },
  {
    domain: 'law',
    materialKind: 'case',
    jurisdiction: '合成法域',
    effectLevel: '仅测试',
    effectiveFrom: '2026-01-01',
    effectiveUntil: '2026-06-30',
    applicabilityDate: '2026-09-01',
    caseNumber: 'synthetic-case',
    originalLocation: '第一段',
    limitations: '不构成法律意见',
  },
  {
    domain: 'finance',
    materialKind: 'financial_statement',
    reportingPeriod: '2025年度',
    asOf: '2025-12-31',
    currency: 'CNY',
    unit: '万元',
    accountingBasis: '合成口径',
    limitations: '不构成投资建议',
  },
]
it.each(professionalContexts)(
  'preserves actual $domain research context through frozen evidence, historical reviews and delivery',
  async (context) => {
    const f = await setup(false, context)
    try {
      const evidence = await f.tool('read_presentation_claim_evidence', f.input)
      expect(evidence.isError, evidence.output).toBeFalsy()
      const read = JSON.parse(evidence.output)
      expect(read.claim.professionalContext).toEqual(context)
      expect(read.research.record.draft.facts[0].professionalContext).toEqual(context)
      const saved = await f.tool('record_presentation_claim_review', {
        ...f.input,
        review_id: 'professional-review',
        outcome: 'supported',
        notes: '仅原窗口支持，不认证专业事实',
      })
      expect(saved.isError, saved.output).toBeFalsy()
      const report = await f.raw('production_delivery_report', { requestId: 'frozen' })
      expect(report.plan.claims[0].professionalContext).toEqual(context)
      const codes = report.pages[0].issues.map((issue: { code: string }) => issue.code)
      if (context.domain === 'law')
        expect(codes).toEqual(
          expect.arrayContaining([
            'professional_legal_rule_inactive',
            'professional_jurisdiction_mismatch',
          ]),
        )
      if (context.domain === 'finance') expect(codes).toContain('professional_financial_time_mixed')
      if (context.domain === 'science')
        expect(codes).not.toContain('professional_context_incomplete')
      expect(report.checks.sourceAuthority).toBe('not_verified')
      const altered = structuredClone(f.plan)
      altered.claims[0]!.professionalContext = { ...context, limitations: '删改原专业限定' }
      expect(await f.raw('save_plan', { expectedRevision: 1, plan: altered })).toMatchObject({
        error: 'research_binding_invalid',
      })
      f.restart()
      const history = await f.tool('read_presentation_claim_review', {
        project_id: f.plan.projectId,
        request_id: 'frozen',
        review_id: 'professional-review',
      })
      expect(history.isError, history.output).toBeFalsy()
      expect(JSON.parse(history.output).reviewId).toBe('professional-review')
      const reread = await f.tool('read_presentation_claim_evidence', f.input)
      expect(JSON.parse(reread.output).claim.professionalContext).toEqual(context)
      f.omitProfessionalContext()
      const forged = await f.tool('read_presentation_claim_evidence', f.input)
      expect(forged).toMatchObject({ isError: true, output: 'presentation_response_invalid' })
      expect(f.compile).not.toHaveBeenCalled()
    } finally {
      f.close()
    }
  },
)
it('carries professional context without a research binding and prevents omitted context from authorizing a review', async () => {
  const context = professionalContexts[0]!
  const f = await setup(false, context, false)
  try {
    const evidence = await f.tool('read_presentation_claim_evidence', f.input)
    expect(evidence.isError, evidence.output).toBeFalsy()
    const read = JSON.parse(evidence.output)
    expect(read.research).toBeUndefined()
    expect(read.documentId).toBe('doc')
    expect(read.claim.professionalContext).toEqual(context)
    const saved = await f.tool('record_presentation_claim_review', {
      ...f.input,
      review_id: 'unbound-professional',
      outcome: 'supported',
      notes: '仅测试原文',
    })
    expect(saved.isError, saved.output).toBeFalsy()
    f.omitProfessionalContext()
    expect(await f.tool('read_presentation_claim_evidence', f.input)).toMatchObject({
      isError: true,
      output: 'presentation_response_invalid',
    })
    expect(f.compile).not.toHaveBeenCalled()
  } finally {
    f.close()
  }
})

it('refuses to write when a professional response is stripped to a legacy shape', async () => {
  const f = await setup(false, professionalContexts[0]!, false)
  try {
    f.stripProfessionalContext()
    const read = await f.tool('read_presentation_claim_evidence', f.input)
    expect(read.isError, read.output).toBeFalsy()
    expect(JSON.parse(read.output).claim).toBeUndefined()
    const saved = await f.tool('record_presentation_claim_review', {
      ...f.input,
      review_id: 'stripped-review',
      outcome: 'supported',
      notes: '不可保存丢失的专业上下文',
    })
    expect(saved).toMatchObject({ isError: true, output: 'presentation_evidence_changed' })
    expect(
      await f.raw('production_read_claim_review', {
        requestId: 'frozen',
        reviewId: 'stripped-review',
      }),
    ).toMatchObject({ error: 'not_found' })
  } finally {
    f.close()
  }
})
it('exports professional context and source warnings from the actual immutable production into the delivery ZIP', async () => {
  const context = professionalContexts[1]!
  const f = await setup(false, context)
  try {
    const native = await new JSZip()
      .file('ppt/slides/slide1.xml', '<test-native-current/>')
      .generateAsync({ type: 'uint8array' })
    const vfs = new InMemoryVfs()
    const skill = createPresentationHostBundleSkill({
      available: () => true,
      documentId: async () => 'doc',
      nativeAvailable: () => true,
      exportDocument: async () => native,
      request: f.request,
      vfs,
    })
    const exported = await skill.executeTool({
      id: 'professional-package',
      name: 'export_current_presentation_bundle',
      input: { project_id: f.plan.projectId, request_id: 'frozen' },
    })
    expect(exported.isError, exported.output).toBeFalsy()
    const zip = await JSZip.loadAsync(vfs.readBytes(JSON.parse(exported.output).paths[0]))
    const report = JSON.parse(await zip.file('evidence.json')!.async('string'))
    expect(report.plan.claims[0].professionalContext).toEqual(context)
    expect(report.research.record.draft.facts[0].professionalContext).toEqual(context)
    expect(report.pages[0].issues.map((i: { code: string }) => i.code)).toContain(
      'professional_legal_rule_inactive',
    )
    const research = JSON.parse(await zip.file('research.json')!.async('string'))
    expect(research.draft.facts[0].professionalContext).toEqual(context)
    const markdown = await zip.file('evidence.md')!.async('string')
    expect(markdown).toContain('professionalContext')
    expect(markdown).toContain('synthetic&#45;case')
    expect(report.checks.timeliness).toBe('not_verified')
    expect(f.compile).not.toHaveBeenCalled()
  } finally {
    f.close()
  }
})
