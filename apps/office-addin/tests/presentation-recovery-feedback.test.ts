import { describe, expect, it } from 'vitest'
import { presentationRecoveryReceiptFeedback } from '../src/agent/presentation-recovery-feedback.js'

describe('presentation recovery receipt feedback', () => {
  it('shows semantic scope, saved import counts and the host check without internal names', () => {
    const text = presentationRecoveryReceiptFeedback({
      phase: 'tool_pending',
      toolName: 'import_presentation_production',
      importReceipt: { state: 'uncertain', completed: 2, total: 4 },
    })
    expect(text).toContain('页面导入')
    expect(text).toContain('2/4 页已记录导入')
    expect(text).toContain('下一页写入结果不确定')
    expect(text).toContain('核对宿主页面')
    expect(text).not.toMatch(/import_presentation_production|tool_pending|reconcile_/)
  })
  it('keeps applied, undone and discarded records under settlement wording', () => {
    const text = presentationRecoveryReceiptFeedback({
      phase: 'tool_completed',
      toolName: 'edit_existing_presentation_text',
      changeReceipt: { total: 3, unresolved: 1 },
    })
    expect(text).toContain('2 项已结算')
    expect(text).toContain('撤销或丢弃')
    expect(text).toContain('1 项未结算')
    expect(text).not.toContain('3 项修改已完成')
  })
  it('shows a valid legacy whole-deck import with the journal page limit', () => {
    expect(
      presentationRecoveryReceiptFeedback({
        phase: 'tool_completed',
        importReceipt: { state: 'complete', completed: 100 },
      }),
    ).toContain('100 页已记录导入')
  })
  it.each([
    { importReceipt: { state: 'complete', completed: 2, total: 3 } },
    { importReceipt: { state: 'partial', completed: -1, total: 3 } },
    { importReceipt: { state: 'complete', completed: 1000 } },
    { changeReceipt: { total: 2, unresolved: 3 } },
    { changeReceipt: { total: 65, unresolved: 0 } },
  ])('does not turn damaged receipt counts into progress', (receipt) => {
    const text = presentationRecoveryReceiptFeedback({
      phase: 'tool_pending',
      ...receipt,
    } as Parameters<typeof presentationRecoveryReceiptFeedback>[0])
    expect(text).not.toMatch(/页已记录导入|项已结算|项未结算/)
  })
  it('keeps unknown tool identifiers out of product feedback', () => {
    expect(
      presentationRecoveryReceiptFeedback({
        phase: 'tool_pending',
        toolName: 'private_customer_operation',
      }),
    ).not.toContain('private_customer_operation')
  })
})
