// @vitest-environment jsdom
import React, { act } from 'react'
import { createRoot } from 'react-dom/client'
import { expect, it, vi } from 'vitest'
import { PresentationDeliveryReportCard } from '../src/agent/presentation-delivery-report-card.js'
import type { PresentationDeliveryReport } from '@wiswork/pptx-engine/presentation-delivery-report'
import type { PresentationProjectController } from '../src/skills/powerpoint/presentation-project.js'

it('lets the reviewer reach and record issues after the first 20', async () => {
  const issues = Array.from({ length: 25 }, (_, index) => ({
    id: `issue-${index}`,
    code: index === 0 ? 'source_original_not_frozen' : 'source_review_missing',
    claimId: 'claim',
    sourceId: 'source',
    digest: 'a'.repeat(64),
    category: 'unverifiable' as const,
    disposition: { state: 'open' as const, stale: false },
  }))
  const report = {
    requestId: 'request',
    planRevision: 1,
    plan: {
      claims: [{ id: 'claim', statement: 'Claim' }],
      sources: [
        {
          id: 'source',
          title: '审计报告',
          uri: 'attachment:abc',
          locator: '第 12 页',
          asOf: '2025-12-31',
          excerpt: '<script>不可执行的原文</script>',
        },
      ],
    },
    issueLedger: { revision: 0, actions: [] },
    sourceAudit: [{ sourceId: 'source', attachmentId: 'a'.repeat(64), status: 'not_found' }],
    pages: [
      { pageId: 'page', title: 'Page', productionState: 'compiled', calculations: [], issues },
    ],
  } as unknown as PresentationDeliveryReport
  const recordIssueAction = vi.fn()
  const controller = { recordIssueAction } as unknown as PresentationProjectController
  const container = document.createElement('div')
  const root = createRoot(container)
  try {
    await act(async () =>
      root.render(
        React.createElement(PresentationDeliveryReportCard, {
          report,
          controller,
          disabled: false,
        }),
      ),
    )
    expect(container.querySelector('[aria-label="处置状态 issue-24"]')).toBeNull()
    expect(container.textContent).toContain('显示更多问题 · 剩余 5 项')
    expect(container.querySelector('a[href="#evidence-source-request-source"]')).not.toBeNull()
    expect(container.querySelector('#evidence-source-request-source')?.textContent).toContain(
      '第 12 页',
    )
    expect(container.textContent).toContain('完整原文中未找到片段')
    expect(container.textContent).toContain('项目未保存该来源原文')
    expect(container.textContent).toContain('不核验事实支持')
    expect(container.querySelector('#evidence-source-request-source script')).toBeNull()
    await act(async () =>
      Array.from(container.querySelectorAll('button'))
        .find((button) => button.textContent?.includes('显示更多问题'))!
        .click(),
    )
    expect(container.querySelector('[aria-label="处置状态 issue-24"]')).not.toBeNull()
    expect(container.textContent).not.toContain('显示更多问题')
    const note = container.querySelector('[aria-label="处置理由 issue-24"]') as HTMLTextAreaElement
    await act(async () => {
      const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
      setter.call(note, '人工复核后补充说明')
      note.dispatchEvent(new Event('input', { bubbles: true }))
    })
    await act(async () =>
      note.closest('form')!.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true })),
    )
    expect(recordIssueAction).toHaveBeenCalledWith(expect.objectContaining({ issueId: 'issue-24' }))
  } finally {
    await act(async () => root.unmount())
  }
})

it('shows exact frozen research identities, both conflict statements and every binding finding without certifying them', async () => {
  const { researchRecord } = await import('./presentation-research-fixture.js')
  const record = researchRecord()
  const report = {
    requestId: 'request-old',
    planRevision: 1,
    plan: { sources: [], claims: [] },
    pages: [],
    sourceAudit: [],
    issueLedger: { revision: 0, actions: [] },
    research: {
      record,
      findings: [
        {
          code: 'omitted_conflict_partner',
          claimId: 'renamed-claim',
          researchClaimId: 'claim1',
          relatedResearchClaimId: 'claim2',
        },
        {
          code: 'unselected_source_ref',
          claimId: 'renamed-claim',
          researchClaimId: 'claim1',
          sourceId: 'source1',
        },
        {
          code: 'source_unavailable',
          claimId: 'renamed-claim',
          researchClaimId: 'claim1',
          sourceId: 'source1',
        },
        { code: 'unmapped_claim', claimId: 'standalone-claim' },
      ],
    },
  } as unknown as PresentationDeliveryReport
  const controller = { recordIssueAction: vi.fn() } as unknown as PresentationProjectController
  const container = document.createElement('div')
  const root = createRoot(container)
  try {
    await act(async () =>
      root.render(
        React.createElement(PresentationDeliveryReportCard, {
          report,
          controller,
          disabled: false,
        }),
      ),
    )
    expect(container.textContent).toContain('冻结计划绑定研究 #1')
    expect(container.textContent).toContain('ledger1')
    expect(container.textContent).toContain('销售增长')
    expect(container.textContent).toContain('销售下降')
    expect(container.textContent).toContain('未选用冲突另一方')
    expect(container.textContent).toContain('研究引用未选入计划')
    expect(container.textContent).toContain('原文尚不可用')
    expect(container.textContent).toContain('计划主张未映射研究')
    expect(container.textContent).toContain('renamed-claim')
    expect(container.textContent).toContain('claim2')
    expect(container.textContent).toContain('source1')
    expect(container.textContent).toContain('不代表事实支持或 QA 通过')
    await act(async () =>
      root.render(
        React.createElement(PresentationDeliveryReportCard, {
          report: { ...report, research: undefined },
          controller,
          disabled: false,
        }),
      ),
    )
    expect(container.textContent).not.toContain('冻结计划绑定研究')
  } finally {
    await act(async () => root.unmount())
  }
})

it('shows page research context and keeps explained, deferred and reopened actions isolated by issue and task', async () => {
  const { researchRecord } = await import('./presentation-research-fixture.js')
  const record = researchRecord()
  record.draft.facts[0]!.asOf = '2025-12-31'
  record.draft.facts[0]!.jurisdiction = '中国大陆'
  record.draft.facts[1]!.jurisdiction = '欧洲市场'
  const issues = ['page-a', 'page-b'].map((pageId) => ({
    id: `${pageId}-conflict`,
    code: 'research_claim_conflict',
    claimId: 'claim',
    digest: (pageId === 'page-a' ? 'a' : 'b').repeat(64),
    category: 'needs_human',
    disposition: { state: 'open', stale: false },
    research: {
      ledgerId: record.id,
      sequence: record.sequence,
      draftDigest: record.draftDigest,
      researchClaimId: 'claim1',
      relatedClaimIds: ['claim2'],
      sourceIds: ['source1'],
    },
  }))
  const report = {
    requestId: 'request-a',
    planRevision: 1,
    plan: { sources: [], claims: [{ id: 'claim', statement: '计划销售趋势' }] },
    sourceAudit: [],
    issueLedger: { revision: 0, actions: [] },
    research: { record, findings: [] },
    pages: issues.map((issue, index) => ({
      pageId: `page-${index}`,
      title: `页面${index}`,
      productionState: 'compiled',
      calculations: [],
      issues: [issue],
    })),
  } as unknown as PresentationDeliveryReport
  const recordIssueAction = vi.fn()
  const controller = { recordIssueAction } as unknown as PresentationProjectController
  const container = document.createElement('div')
  document.body.append(container)
  const root = createRoot(container)
  const render = async (value = report, disabled = false) =>
    act(async () =>
      root.render(
        React.createElement(PresentationDeliveryReportCard, {
          report: value,
          controller,
          disabled,
        }),
      ),
    )
  const edit = async (id: string, state: string, note: string) =>
    act(async () => {
      const select = container.querySelector(`[aria-label="处置状态 ${id}"]`) as HTMLSelectElement
      select.value = state
      select.dispatchEvent(new Event('change', { bubbles: true }))
      const textarea = container.querySelector(
        `[aria-label="处置理由 ${id}"]`,
      ) as HTMLTextAreaElement
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!.call(
        textarea,
        note,
      )
      textarea.dispatchEvent(new Event('input', { bubbles: true }))
    })
  const submit = async (id: string) =>
    act(async () => {
      const form = container.querySelector(`[aria-label="处置理由 ${id}"]`)!.closest('form')!
      ;(form.querySelector('button') as HTMLButtonElement).click()
    })
  try {
    await render()
    const page = container.querySelector('[aria-label="证据页面 页面0"]')!
    expect(page.textContent).toContain('研究结论存在冲突')
    const context = page.querySelector(
      'details[aria-label="逐页研究问题上下文"]',
    ) as HTMLDetailsElement
    expect(context.open).toBe(false)
    expect(context.textContent).toContain('ledger1')
    expect(context.textContent).toContain('销售增长')
    expect(context.textContent).toContain('销售下降')
    expect(context.textContent).toContain('中国大陆')
    expect(context.textContent).toContain('欧洲市场')
    expect(context.textContent).toContain('2025-12-31')
    expect(context.querySelector('a')?.getAttribute('href')).toBe('https://example.com/report')
    expect(context.textContent).toContain('原文尚不可用')
    expect(context.textContent).toContain('已说明或暂缓仍保留研究缺口')
    await edit('page-a-conflict', 'explained', '保留相反结论供审查')
    await submit('page-a-conflict')
    expect(recordIssueAction).toHaveBeenLastCalledWith(
      expect.objectContaining({
        issueId: 'page-a-conflict',
        issueDigest: 'a'.repeat(64),
        state: 'explained',
        note: '保留相反结论供审查',
      }),
    )
    await edit('page-b-conflict', 'deferred', '等待原始资料')
    await submit('page-b-conflict')
    expect(recordIssueAction).toHaveBeenLastCalledWith(
      expect.objectContaining({ issueId: 'page-b-conflict', state: 'deferred' }),
    )
    await edit('page-a-conflict', 'open', '重新打开审查')
    await submit('page-a-conflict')
    expect(recordIssueAction).toHaveBeenLastCalledWith(
      expect.objectContaining({ issueId: 'page-a-conflict', state: 'open' }),
    )
    await render({ ...report, requestId: 'request-b' })
    expect(
      (container.querySelector('[aria-label="处置理由 page-a-conflict"]') as HTMLTextAreaElement)
        .value,
    ).toBe('')
    await render({ ...report, requestId: 'request-b' }, true)
    expect(
      (container.querySelector('[aria-label="处置状态 page-a-conflict"]') as HTMLSelectElement)
        .disabled,
    ).toBe(true)
    for (const [code, reason] of [
      ['research_unmapped_claim', '计划主张未映射研究'],
      ['research_conflict_partner_omitted', '未选用冲突另一方'],
      ['research_source_reference_unselected', '研究引用未选入计划'],
      ['research_source_unavailable', '原文尚不可用'],
    ]) {
      const changed = {
        ...report,
        pages: [
          {
            ...report.pages[0]!,
            issues: [
              {
                ...report.pages[0]!.issues[0]!,
                code: code!,
                research: {
                  ...report.pages[0]!.issues[0]!.research!,
                  sourceIds: [],
                },
              },
            ],
          },
        ],
      }
      await render(changed)
      expect(container.querySelector('[aria-label="证据页面 页面0"]')!.textContent).toContain(
        reason,
      )
      expect(container.querySelector('[aria-label="逐页研究问题上下文"] a')).not.toBeNull()
    }
    await render({
      ...report,
      research: { ...report.research!, record: { ...record, id: 'other-ledger' } },
    })
    expect(container.querySelector('[aria-label="逐页研究问题上下文"]')).toBeNull()
    await render({ ...report, research: undefined })
    expect(container.querySelector('[aria-label="逐页研究问题上下文"]')).toBeNull()
  } finally {
    await act(async () => root.unmount())
    container.remove()
  }
})

it('shows every scoped historical source assessment and Chinese source issue reasons without certification', async () => {
  const reasons = {
    source_authority_review_missing: '缺少来源权威性判断',
    source_authority_review_uncertain: '来源权威性尚不确定',
    source_authority_review_insufficient: '来源权威性不足以支持该主张',
    source_authority_review_mixed: '来源权威性判断或声明级别不同',
    source_timeliness_review_missing: '缺少来源时效判断',
    source_timeliness_review_uncertain: '来源时效尚不确定',
    source_timeliness_review_historical_only: '来源仅适用于历史时点',
    source_timeliness_review_superseded: '来源已被后续资料取代',
    source_timeliness_review_mixed: '来源时效判断或比较框架不同',
    source_jurisdiction_review_missing: '缺少来源适用范围判断',
    source_jurisdiction_review_uncertain: '来源适用范围尚不确定',
    source_jurisdiction_review_mismatch: '来源与主张适用范围不匹配',
    source_jurisdiction_review_mixed: '来源适用范围判断不同',
  }
  const assessment = {
    scope: '该销售主张的来源',
    authority: {
      outcome: 'appropriate_for_claim',
      sourceTier: 'primary',
      reason: '发布方提供原始数据',
    },
    timeliness: {
      outcome: 'current_for_claim',
      referenceDate: '2026-09-29',
      claimAsOf: '2025年度',
      sourceAsOf: '2025-12-31',
      reason: '适用于冻结主张时点',
    },
    jurisdiction: {
      claimJurisdiction: '中国大陆',
      outcome: 'applicable',
      reason: '资料覆盖中国大陆',
    },
    basis: [{ offset: 40, text: '<script>销售原文</script>' }],
  }
  const reviews = ['older', 'newer', 'other-page'].map((reviewId, index) => ({
    requestId: 'request',
    pageId: index === 2 ? 'page-b' : 'page-a',
    claimId: 'claim',
    sourceId: 'source',
    reviewId,
    offset: 40 + index,
    maxChars: 80,
    createdAt: `2026-09-29T00:0${index}:00.000Z`,
    sourceAssessment:
      index === 1
        ? {
            ...assessment,
            authority: {
              outcome: 'uncertain',
              sourceTier: 'secondary',
              reason: '较晚判断仍无独立权威核验',
            },
            timeliness: {
              ...assessment.timeliness,
              referenceDate: '2026-09-28',
              reason: '采用另一比较时点',
            },
          }
        : assessment,
  }))
  const issues = Object.keys(reasons).map((code, index) => ({
    id: `assessment-${index}`,
    code,
    claimId: 'claim',
    sourceId: 'source',
    digest: 'a'.repeat(64),
    category: 'needs_human',
    disposition: { state: 'open', stale: false },
  }))
  const report = {
    requestId: 'request',
    planRevision: 1,
    plan: {
      claims: [{ id: 'claim', statement: '销售主张' }],
      sources: [{ id: 'source', title: '原资料', uri: 'attachment:source', excerpt: '' }],
    },
    pages: [
      { pageId: 'page-a', title: 'A', productionState: 'compiled', calculations: [], issues },
      {
        pageId: 'page-b',
        title: 'B',
        productionState: 'compiled',
        calculations: [],
        issues: [{ ...issues[0], id: 'other-issue', digest: 'b'.repeat(64) }],
      },
    ],
    reviews,
    sourceAudit: [],
    issueLedger: { revision: 0, actions: [] },
  } as unknown as PresentationDeliveryReport
  const recordIssueAction = vi.fn()
  const controller = { recordIssueAction } as unknown as PresentationProjectController
  const container = document.createElement('div')
  document.body.append(container)
  const root = createRoot(container)
  const render = async (value = report) =>
    act(async () =>
      root.render(
        React.createElement(PresentationDeliveryReportCard, {
          report: value,
          controller,
          disabled: false,
        }),
      ),
    )
  const action = async (id: string, state: string) => {
    await act(async () => {
      const select = container.querySelector(`[aria-label="处置状态 ${id}"]`) as HTMLSelectElement
      select.value = state
      select.dispatchEvent(new Event('change', { bubbles: true }))
      const note = container.querySelector(`[aria-label="处置理由 ${id}"]`) as HTMLTextAreaElement
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!.call(
        note,
        '保留来源判断限制',
      )
      note.dispatchEvent(new Event('input', { bubbles: true }))
    })
    await act(async () =>
      (
        container
          .querySelector(`[aria-label="处置理由 ${id}"]`)!
          .closest('form')!
          .querySelector('button') as HTMLButtonElement
      ).click(),
    )
  }
  try {
    await render()
    for (const reason of Object.values(reasons)) expect(container.textContent).toContain(reason)
    const pageA = container.querySelector('[aria-label="证据页面 A"]')!
    const detail = pageA.querySelector(
      'details[aria-label="来源评估历史 claim source"]',
    ) as HTMLDetailsElement
    expect(detail.open).toBe(false)
    for (const text of [
      'older',
      'newer',
      '窗口 UTF-16 40 · 最多 80 字符',
      '2026-09-29T00:00:00.000Z',
      '一手来源',
      '二手来源',
      '发布方提供原始数据',
      '较晚判断仍无独立权威核验',
      '2026-09-28',
      '2026-09-29',
      '2025年度',
      '2025-12-31',
      '中国大陆',
      '<script>销售原文</script>',
      '历史 Agent 判断',
      '不代表事实矛盾',
    ])
      expect(detail.textContent).toContain(text)
    expect(detail.textContent).not.toContain('other-page')
    expect(detail.querySelector('script')).toBeNull()
    const pageB = container.querySelector('[aria-label="证据页面 B"]')!
    expect(
      pageB.querySelector('details[aria-label="来源评估历史 claim source"]')?.textContent,
    ).toContain('other-page')
    expect(pageB.textContent).not.toContain('较晚判断仍无独立权威核验')
    expect(container.textContent).toContain('来源真实性、时效未核验')
    await action('assessment-0', 'explained')
    expect(recordIssueAction).toHaveBeenLastCalledWith(
      expect.objectContaining({
        issueId: 'assessment-0',
        issueDigest: 'a'.repeat(64),
        state: 'explained',
      }),
    )
    await action('other-issue', 'deferred')
    expect(recordIssueAction).toHaveBeenLastCalledWith(
      expect.objectContaining({
        issueId: 'other-issue',
        issueDigest: 'b'.repeat(64),
        state: 'deferred',
      }),
    )
    await action('assessment-0', 'open')
    expect(recordIssueAction).toHaveBeenLastCalledWith(
      expect.objectContaining({ issueId: 'assessment-0', state: 'open' }),
    )
    await render({ ...report, pages: [{ ...report.pages[1]!, issues: [] }] })
    expect(
      container.querySelector('[aria-label="来源评估历史 claim source"]')?.textContent,
    ).toContain('适合该主张')
    await render({ ...report, requestId: 'other-task' })
    expect(container.querySelector('[aria-label="来源评估历史 claim source"]')).toBeNull()
    expect(
      (container.querySelector('[aria-label="处置理由 assessment-0"]') as HTMLTextAreaElement)
        .value,
    ).toBe('')
    await render({
      ...report,
      reviews: reviews.map(({ sourceAssessment: _assessment, ...review }) => review),
    } as unknown as PresentationDeliveryReport)
    expect(container.querySelector('[aria-label="来源评估历史 claim source"]')).toBeNull()
  } finally {
    await act(async () => root.unmount())
    container.remove()
  }
})

it('shows frozen page professional context, missing fields and original research context with isolated actions', async () => {
  const reasons = {
    professional_context_incomplete: '专业上下文尚有缺失',
    professional_source_secondary: '专业结论来源为二手或未核验资料',
    professional_legal_rule_inactive: '法律材料不在明确适用日期范围内',
    professional_jurisdiction_mismatch: '通用与专业适用范围不同',
    professional_financial_time_mixed: '通用与专业数据时点不同',
    professional_financial_unit_mismatch: '计算与专业上下文单位不同',
    professional_financial_currency_mismatch: '计算与专业上下文币种不同',
  }
  const contexts = [
    {
      domain: 'science',
      materialKind: 'paper',
      publicationId: 'DOI:original',
      version: 'v1',
      sample: '样本100',
      method: '随机分组',
      statisticalBasis: '95%区间',
      limitations: '仅覆盖样本',
    },
    {
      domain: 'law',
      materialKind: 'case',
      jurisdiction: '中国大陆',
      effectLevel: '参考裁判',
      effectiveFrom: '2020-01-01',
      effectiveUntil: '2025-12-31',
      applicabilityDate: '2026-09-29',
      caseNumber: '原案号',
      originalLocation: '原裁判第3段',
      limitations: '仅限原案情',
    },
    {
      domain: 'finance',
      materialKind: 'financial_statement',
      reportingPeriod: '2025年度',
      asOf: '2025-12-31',
      currency: 'CNY',
      unit: '万元',
      accountingBasis: '原会计准则',
      formula: 'a+b',
      limitations: '未审计',
    },
    { domain: 'science' },
  ]
  const { researchRecord } = await import('./presentation-research-fixture.js')
  const record = researchRecord()
  const claims = contexts.map((professionalContext, index) => ({
    id: `claim-${index}`,
    type: index === 2 ? 'calculation' : 'fact',
    statement: `原主张${index}`,
    professionalContext,
  }))
  const issues = Object.keys(reasons).map((code, index) => ({
    id: `professional-${index}`,
    code,
    claimId: 'claim-0',
    digest: 'c'.repeat(64),
    category: 'needs_human',
    disposition: { state: 'open', stale: false },
  }))
  const report = {
    requestId: 'professional-task',
    planRevision: 1,
    plan: {
      sources: [],
      claims,
      slides: [
        { id: 'page-a', claimIds: claims.map((c) => c.id) },
        { id: 'page-b', claimIds: ['claim-1'] },
      ],
    },
    pages: [
      { pageId: 'page-a', title: '专业页', productionState: 'compiled', calculations: [], issues },
      {
        pageId: 'page-b',
        title: '法律页',
        productionState: 'compiled',
        calculations: [],
        issues: [{ ...issues[0], id: 'other-professional' }],
      },
    ],
    reviews: [],
    sourceAudit: [],
    issueLedger: { revision: 0, actions: [] },
    research: {
      record: {
        ...record,
        draft: {
          ...record.draft,
          facts: record.draft.facts.map((fact) => ({ ...fact, professionalContext: contexts[0] })),
        },
      },
      findings: [],
    },
  } as unknown as PresentationDeliveryReport
  const recordIssueAction = vi.fn()
  const controller = { recordIssueAction } as unknown as PresentationProjectController
  const container = document.createElement('div')
  document.body.append(container)
  const root = createRoot(container)
  const render = async (value = report) =>
    act(async () =>
      root.render(
        React.createElement(PresentationDeliveryReportCard, {
          report: value,
          controller,
          disabled: false,
        }),
      ),
    )
  try {
    await render()
    for (const reason of Object.values(reasons)) expect(container.textContent).toContain(reason)
    const page = container.querySelector('[aria-label="证据页面 专业页"]')!
    const details = page.querySelectorAll('details[aria-label^="专业上下文"]')
    expect(details).toHaveLength(4)
    for (const detail of Array.from(details))
      expect((detail as HTMLDetailsElement).open).toBe(false)
    for (const text of [
      '科研',
      '法律',
      '金融',
      'DOI:original',
      '样本100',
      '随机分组',
      '95%区间',
      '原案号',
      '原裁判第3段',
      '2020-01-01',
      '2025-12-31',
      '2026-09-29',
      '原会计准则',
      'CNY',
      '万元',
      'a+b',
      '专业字段缺失：材料类型、出版或发布标识、版本、样本、方法、统计依据、局限',
      '完整字段仍不代表事实支持',
    ])
      expect(page.textContent).toContain(text)
    expect(container.querySelector('[aria-label="证据页面 法律页"]')!.textContent).not.toContain(
      'DOI:original',
    )
    expect(
      container.querySelector('[aria-label="冻结计划绑定研究"] [aria-label="专业上下文 claim1"]')
        ?.textContent,
    ).toContain('DOI:original')
    const note = container.querySelector(
      '[aria-label="处置理由 professional-0"]',
    ) as HTMLTextAreaElement
    await act(async () => {
      const select = container.querySelector(
        '[aria-label="处置状态 professional-0"]',
      ) as HTMLSelectElement
      select.value = 'explained'
      select.dispatchEvent(new Event('change', { bubbles: true }))
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!.call(
        note,
        '保留专业限制',
      )
      note.dispatchEvent(new Event('input', { bubbles: true }))
    })
    await act(async () =>
      (note.closest('form')!.querySelector('button') as HTMLButtonElement).click(),
    )
    expect(recordIssueAction).toHaveBeenCalledWith(
      expect.objectContaining({
        issueId: 'professional-0',
        issueDigest: 'c'.repeat(64),
        state: 'explained',
        note: '保留专业限制',
      }),
    )
    expect(
      (container.querySelector('[aria-label="处置理由 other-professional"]') as HTMLTextAreaElement)
        .value,
    ).toBe('')
    await render({
      ...report,
      requestId: 'next-task',
      research: undefined,
      plan: {
        ...report.plan,
        claims: report.plan.claims.map(({ professionalContext: _context, ...claim }) => claim),
      },
    })
    expect(container.querySelector('[aria-label^="专业上下文"]')).toBeNull()
    expect(
      (container.querySelector('[aria-label="处置理由 professional-0"]') as HTMLTextAreaElement)
        .value,
    ).toBe('')
  } finally {
    await act(async () => root.unmount())
    container.remove()
  }
})

it('binds complete professional workflow guidance to the current report domain and isolates missing-context actions', async () => {
  const { benchmarkPlan } =
    await import('../../../packages/pptx-engine/tests/fixtures/presentation-plan.js')
  const { benchmarkDeck } =
    await import('../../../packages/pptx-engine/tests/fixtures/presentation-benchmark.js')
  const { buildPresentationDeliveryReport } =
    await import('@wiswork/pptx-engine/presentation-delivery-report')
  const { PRESENTATION_DOMAIN_PROFILES, presentationPlanClaims } =
    await import('@wiswork/pptx-engine/presentation-plan')
  const makeReport = async (domain?: 'science' | 'law' | 'finance', complete = false) => {
    const plan = benchmarkPlan(),
      deck = benchmarkDeck()
    if (domain) {
      plan.domain = domain
      plan.slides.forEach((slide, index) => {
        slide.domainSection = PRESENTATION_DOMAIN_PROFILES[domain].sections[index % 5]
      })
    }
    if (complete)
      plan.claims[0]!.professionalContext = {
        domain: 'law',
        materialKind: 'contract',
        jurisdiction: '原法域',
        effectLevel: '合同约定',
        applicabilityDate: '2026-09-29',
        originalLocation: '第1条',
        limitations: '仅合同范围',
      }
    deck.claims = presentationPlanClaims(plan)
    const metadata = {
      projectId: plan.projectId,
      documentId: 'doc',
      requestId: `${domain ?? 'old'}-task`,
      planRevision: 1,
      inputDigest: 'a'.repeat(64),
      planDigest: 'b'.repeat(64),
    }
    return buildPresentationDeliveryReport({
      plan,
      deck,
      metadata,
      reviews: [],
      pageStates: plan.slides.map((slide) => ({ pageId: slide.id, state: 'compiled' })),
      issueLedger: {
        version: 1,
        projectId: metadata.projectId,
        documentId: 'doc',
        requestId: metadata.requestId,
        inputDigest: metadata.inputDigest,
        planDigest: metadata.planDigest,
        revision: 0,
        actions: [],
      },
    })
  }
  const report = await makeReport('science')
  const issue = report.pages[0]!.issues.find(
    (item) => item.code === 'professional_context_missing',
  )!
  const otherIssue = report.pages[1]!.issues.find(
    (item) => item.code === 'professional_context_missing',
  )!
  const recordIssueAction = vi.fn()
  const controller = { recordIssueAction } as unknown as PresentationProjectController
  const container = document.createElement('div')
  document.body.append(container)
  const root = createRoot(container)
  const render = async (value = report) =>
    act(async () =>
      root.render(
        React.createElement(PresentationDeliveryReportCard, {
          report: value,
          controller,
          disabled: false,
        }),
      ),
    )
  try {
    await render()
    expect(container.textContent).toContain('专业领域主张缺少专业上下文')
    const detail = container.querySelector('[aria-label="专业制作工作流"]') as HTMLDetailsElement
    expect(detail.open).toBe(false)
    expect(detail.textContent).toContain('科研')
    const workflow = report.professionalWorkflow!
    for (const text of [
      ...workflow.sourcePriority,
      ...['材料类型', '出版或发布标识', '版本', '样本', '方法', '统计依据', '局限'],
      ...workflow.manualChecks,
      workflow.disclosure,
      ...workflow.reviewSteps.flatMap((step) => [step.title, step.instruction]),
    ])
      expect(detail.textContent).toContain(text)
    for (const label of ['来源优先级', '专业上下文字段', '复核步骤', '人工检查', '范围说明'])
      expect(detail.textContent).toContain(label)
    expect(
      container.querySelector(`[aria-label="证据页面 ${report.pages[0]!.title}"]`)?.textContent,
    ).toContain('下一步：读取该主张来源原文，再补充专业限定；无法确认的字段保持未知。')
    const note = container.querySelector(
      `[aria-label="处置理由 ${issue.id}"]`,
    ) as HTMLTextAreaElement
    await act(async () => {
      const select = container.querySelector(
        `[aria-label="处置状态 ${issue.id}"]`,
      ) as HTMLSelectElement
      select.value = 'deferred'
      select.dispatchEvent(new Event('change', { bubbles: true }))
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!.call(
        note,
        '等待原文专业审查',
      )
      note.dispatchEvent(new Event('input', { bubbles: true }))
    })
    await act(async () =>
      (note.closest('form')!.querySelector('button') as HTMLButtonElement).click(),
    )
    expect(recordIssueAction).toHaveBeenCalledWith(
      expect.objectContaining({ issueId: issue.id, issueDigest: issue.digest, state: 'deferred' }),
    )
    expect(
      (container.querySelector(`[aria-label="处置理由 ${otherIssue.id}"]`) as HTMLTextAreaElement)
        .value,
    ).toBe('')
    for (const [domain, title] of [
      ['law', '法律'],
      ['finance', '金融'],
    ] as const) {
      const next = await makeReport(domain)
      await render(next)
      const nextDetail = container.querySelector('[aria-label="专业制作工作流"]')!
      expect(nextDetail.querySelector('summary')?.textContent).toContain(title)
      expect(nextDetail.querySelector('summary')?.textContent).not.toContain('科研')
      for (const instruction of next.professionalWorkflow!.manualChecks)
        expect(nextDetail.textContent).toContain(instruction)
      expect(
        (container.querySelector(`[aria-label="处置理由 ${issue.id}"]`) as HTMLTextAreaElement)
          .value,
      ).toBe('')
    }
    const positive = await makeReport('science', true)
    expect(positive.pages[0]!.issues.some((item) => item.code.startsWith('professional_'))).toBe(
      false,
    )
    await render(positive)
    expect(container.querySelector('[aria-label="专业制作工作流"]')).not.toBeNull()
    expect(container.textContent).not.toContain('下一步：读取该主张来源原文')
    expect(container.querySelector('[aria-label="专业上下文 source-1"]')).not.toBeNull()
    await render(await makeReport())
    expect(container.querySelector('[aria-label="专业制作工作流"]')).toBeNull()
    expect(container.textContent).not.toContain('下一步：读取该主张来源原文')
  } finally {
    await act(async () => root.unmount())
    container.remove()
  }
})
