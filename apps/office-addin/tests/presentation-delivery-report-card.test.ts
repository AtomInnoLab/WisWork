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
