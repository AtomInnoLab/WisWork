// @vitest-environment jsdom
import React, { act } from 'react'
import { createRoot } from 'react-dom/client'
import { expect, it, vi } from 'vitest'
import { PresentationDeliveryReportCard } from '../src/agent/presentation-delivery-report-card.js'
import { buildPresentationDeliveryReport } from '@wiswork/pptx-engine/presentation-delivery-report'
import {
  professionalAssessmentInput,
  professionalAssessmentReview,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-professional-assessment.js'
import type { PresentationProjectController } from '../src/skills/powerpoint/presentation-project.js'
it('shows folded full professional history and uses existing exact issue action identity without closing conflicts', async () => {
  const input = professionalAssessmentInput()
  input.reviews = [
    professionalAssessmentReview(input),
    professionalAssessmentReview(input, 'later', ['consistent', 'consistent']),
  ]
  const report = await buildPresentationDeliveryReport(input),
    node = document.createElement('div'),
    root = createRoot(node)
  const recordIssueAction = vi.fn(async () => {})
  const controller = { recordIssueAction } as unknown as PresentationProjectController
  document.body.appendChild(node)
  const render = async () =>
    act(async () =>
      root.render(
        React.createElement(PresentationDeliveryReportCard, {
          report,
          controller,
          disabled: false,
        }),
      ),
    )
  try {
    await render()
    const history = node.querySelector(
      'details[aria-label="专业评估历史 source-1 source"]',
    ) as HTMLDetailsElement
    expect(history).not.toBeNull()
    expect(history.open).toBe(false)
    expect(history.textContent).toContain('professional-review')
    expect(history.textContent).toContain('later')
    expect(history.textContent).toContain('结论适用范围')
    expect(history.textContent).toContain('专业限定')
    expect(history.textContent).toContain('样本：样本仅100人')
    expect(history.textContent).toContain('局限：仅适用于声明样本')
    expect(history.textContent).toContain('冲突')
    expect(history.textContent).toContain('一致')
    expect(history.textContent).toContain('不确定')
    expect(history.textContent).toContain('原文逐字依据')
    expect(history.textContent).toContain('UTF-16 0')
    expect(history.textContent).toContain('不代表专业结论成立或来源已认证')
    const issue = report.pages[0]!.issues.find(
      (item) => item.code === 'professional_review_conclusion_scope_mixed',
    )!
    expect(node.textContent).toContain('结论适用范围的历史专业判断不同')
    const select = node.querySelector(
      `select[aria-label="处置状态 ${issue.id}"]`,
    ) as HTMLSelectElement
    const form = select.closest('form')!
    const textarea = form.querySelector('textarea')!
    await act(async () => {
      select.value = 'explained'
      select.dispatchEvent(new Event('change', { bubbles: true }))
    })
    await act(async () => {
      const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
      setter.call(textarea, '保留范围问题，不作认证')
      textarea.dispatchEvent(new Event('input', { bubbles: true }))
    })
    const button = form.querySelector('button')!
    expect(button.disabled).toBe(false)
    await act(async () => button.click())
    expect(recordIssueAction).toHaveBeenCalledWith(
      expect.objectContaining({
        issueId: issue.id,
        issueDigest: issue.digest,
        state: 'explained',
        note: '保留范围问题，不作认证',
      }),
    )
    expect(history.textContent).toContain('冲突')
  } finally {
    await act(async () => root.unmount())
    node.remove()
  }
})
it('renders finance dimensions and hides old or foreign-task professional opinions on report change', async () => {
  const input = professionalAssessmentInput({
    domain: 'finance',
    currency: 'CNY',
    unit: '万元',
    accountingBasis: '审计口径',
  })
  input.reviews = [professionalAssessmentReview(input, 'finance', ['uncertain', 'not_applicable'])]
  let report = await buildPresentationDeliveryReport(input)
  const node = document.createElement('div'),
    root = createRoot(node)
  const controller = { recordIssueAction: vi.fn() } as unknown as PresentationProjectController
  document.body.appendChild(node)
  const render = async () =>
    act(async () =>
      root.render(
        React.createElement(PresentationDeliveryReportCard, {
          report,
          controller,
          disabled: false,
        }),
      ),
    )
  try {
    await render()
    expect(node.textContent).toContain('可比口径')
    expect(node.textContent).toContain('预测与前瞻')
    expect(node.textContent).toContain('不适用（仅预测维度）')
    expect(node.textContent).toContain('币种：CNY')
    report = { ...report, requestId: 'another-task' }
    await render()
    expect(node.querySelector('[aria-label="专业评估历史 source-1 source"]')).toBeNull()
    delete input.reviews[0]!.sourceAssessment!.professional
    report = await buildPresentationDeliveryReport(input)
    await render()
    expect(node.querySelector('[aria-label="专业评估历史 source-1 source"]')).toBeNull()
  } finally {
    await act(async () => root.unmount())
    node.remove()
  }
})
