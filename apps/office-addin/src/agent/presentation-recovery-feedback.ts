import { acpPresentationStage, acpPresentationStageLabel } from '@wiswork/agent-harness'
import type { PresentationAgentRunRecovery } from '../skills/powerpoint/presentation-document.js'

type Recovery = Pick<
  PresentationAgentRunRecovery,
  'phase' | 'toolName' | 'importReceipt' | 'changeReceipt'
>

/** A display of recorded outcomes, not a claim that the project or host is verified. */
export function presentationRecoveryReceiptFeedback(record: Recovery | undefined): string {
  if (!record) return ''
  const parts: string[] = []
  const stage = record.toolName ? acpPresentationStage(record.toolName) : undefined
  if (stage) parts.push(`最近工作阶段：${acpPresentationStageLabel(stage)}。`)
  if (record.phase === 'tool_pending') parts.push('该操作结果仍需核对。')
  const count = (value: unknown, maximum: number): value is number =>
    typeof value === 'number' && Number.isSafeInteger(value) && value >= 0 && value <= maximum
  const imported = record.importReceipt
  if (
    imported &&
    ['complete', 'partial', 'uncertain'].includes(imported.state) &&
    count(imported.completed, 100) &&
    (imported.total === undefined ||
      (count(imported.total, 32) && imported.total >= imported.completed)) &&
    (imported.state !== 'complete' ||
      imported.total === undefined ||
      imported.total === imported.completed)
  ) {
    parts.push(
      `对应导入回执：${imported.completed}${imported.total === undefined ? '' : `/${imported.total}`} 页已记录导入；视觉审查与交付仍需核验。`,
    )
    if (imported.state === 'uncertain')
      parts.push('下一页写入结果不确定；请在项目工作台核对宿主页面，核对后再继续，勿直接重试导入。')
    else if (imported.state === 'partial') parts.push('请核对持久导入记录，再从剩余页面继续。')
    else parts.push('请核对宿主页面并继续审查，勿重复导入。')
  }
  const changed = record.changeReceipt
  if (
    changed &&
    count(changed.total, 64) &&
    changed.total > 0 &&
    count(changed.unresolved, changed.total)
  ) {
    parts.push(
      `对应修改历史 ${changed.total} 项，其中 ${changed.unresolved} 项未结算；${changed.total - changed.unresolved} 项已结算（可含撤销或丢弃）。`,
    )
    parts.push(
      changed.unresolved
        ? '请在变更历史核对保存点与宿主对象，再选择继续或撤销；未自动重放修改。'
        : '请在变更历史核对当前应用结果；未自动重放修改。',
    )
  }
  return parts.join('')
}
