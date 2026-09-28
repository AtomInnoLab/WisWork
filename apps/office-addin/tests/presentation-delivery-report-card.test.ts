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
