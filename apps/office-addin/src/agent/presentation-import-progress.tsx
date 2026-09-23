import { useSyncExternalStore } from 'react'
import type { PresentationImportProgress } from '../skills/powerpoint/presentation-page-delivery.js'

export interface PresentationImportProgressController {
  read(): PresentationImportProgress | undefined
  revision(): number
  subscribe(listener: () => void): () => void
}

export function PresentationImportProgressCard({
  controller,
}: {
  controller: PresentationImportProgressController
}) {
  useSyncExternalStore(controller.subscribe, controller.revision, controller.revision)
  let progress: PresentationImportProgress | undefined
  try {
    progress = controller.read()
  } catch {
    return (
      <section className="presentation-project" aria-label="页面导入记录">
        <p role="alert">导入记录无法读取，请检查文档后再继续，避免重复插页。</p>
      </section>
    )
  }
  if (!progress) return null
  return (
    <section className="presentation-project" aria-label="页面导入记录">
      <strong>{progress.source === 'production' ? '页生产任务导入记录' : '页面导入记录'}</strong>
      {progress.source === 'production' && (
        <p>本记录仅确认页面编号与顺序。该任务的逐页质量检查和编辑绑定尚未接入。</p>
      )}
      <p role="status">
        已记录完成 {progress.completed} / {progress.total} 页
      </p>
      <p>
        {progress.status === 'uncertain'
          ? '有页面的写入结果不确定，已停止自动重试，请先检查文档。'
          : progress.status === 'complete'
            ? '全部页面已有导入记录；尚未完成视觉与保存重开验证。'
            : progress.status === 'partial'
              ? '已完成页会保留。可让 Agent 继续导入剩余页，执行前会再次确认。'
              : '页面已准备好，确认后按页导入。'}
      </p>
      <details>
        <summary>查看各页记录</summary>
        <ol>
          {progress.pages.map((page) => (
            <li key={page.id}>
              {page.title} ·{' '}
              {{ pending: '待导入', complete: '已记录完成', uncertain: '结果不确定' }[page.state]}
            </li>
          ))}
        </ol>
      </details>
    </section>
  )
}
