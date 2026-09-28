import { useSyncExternalStore } from 'react'
import type { PresentationQaRecord } from '../skills/powerpoint/presentation-qa.js'
export interface PresentationQaController {
  read(): PresentationQaRecord | undefined
  revision(): number
  subscribe(listener: () => void): () => void
}
export function PresentationQaCard({
  controller,
  onRecheck,
  disabled = false,
}: {
  controller: PresentationQaController
  onRecheck?: (record: PresentationQaRecord, pageIds: string[]) => void
  disabled?: boolean
}) {
  useSyncExternalStore(controller.subscribe, controller.revision, controller.revision)
  let record: PresentationQaRecord | undefined
  try {
    record = controller.read()
  } catch {
    return (
      <section className="presentation-project" aria-label="页面 QA 记录">
        <p role="alert">检查记录无法读取，请重新采集目标页面。</p>
      </section>
    )
  }
  if (!record) return null
  const affected = record.pages.filter((page) => page.recheckRequired).map((page) => page.pageId)
  return (
    <section className="presentation-project" aria-label="页面 QA 记录">
      <strong>
        {record.source === 'production' ? '页生产任务 QA 记录' : '页面 QA 记录'} ·{' '}
        {record.pages.length} 页
      </strong>
      <p>以下为历史检查记录，需重新采集才能确认当前状态；不代表来源核验或保存重开验收。</p>
      {affected.length > 0 && (
        <>
          <p role="status">
            {affected.length} 页的页面或共享样式可能已变化，需重新采集，历史结论不代表当前状态。
            受影响页面 ID：{affected.join('、')}
          </p>
          {onRecheck && (
            <button type="button" disabled={disabled} onClick={() => onRecheck(record, affected)}>
              准备重审受影响页
            </button>
          )}
        </>
      )}
      <details>
        <summary>查看结构与视觉复核</summary>
        <ol>
          {record.pages.map((page) => (
            <li key={page.pageId}>
              <strong>{page.title}</strong>
              <small>页面 ID：{page.pageId}</small>
              {onRecheck && (
                <button
                  type="button"
                  aria-label={`准备重审 ${page.pageId}`}
                  disabled={disabled}
                  onClick={() => onRecheck(record, [page.pageId])}
                >
                  准备重审此页
                </button>
              )}
              {page.recheckRequired && (
                <p>
                  <strong>已发起修改，需重新采集</strong>
                </p>
              )}
              <p>
                {page.recheckRequired ? '历史结构：' : '结构：'}
                {
                  { passed: '未发现几何告警', warning: '有几何告警', incomplete: '检查不完整' }[
                    page.structure.status
                  ]
                }{' '}
                · 越界 {page.structure.overflowCount} · 重叠 {page.structure.overlapCount}
              </p>
              <p>
                {page.recheckRequired ? '历史视觉记录：' : '视觉：'}
                {
                  {
                    needs_review: '截图已采集，待复核',
                    pass: 'Agent 判断通过',
                    needs_changes: 'Agent 建议修改',
                  }[page.visual.status]
                }
              </p>
              {page.screenshotRenderer === 'libreoffice' && (
                <p>图片来自本机 LibreOffice 备用预览；PowerPoint 宿主外观仍待核验。</p>
              )}
              {page.visual.notes && <p>{page.visual.notes}</p>}
              <small>
                采集于 <time dateTime={page.capturedAt}>{page.capturedAt}</time>
              </small>
              {page.visual.reviewedAt && (
                <small>
                  {' '}
                  · 复核记录于{' '}
                  <time dateTime={page.visual.reviewedAt}>{page.visual.reviewedAt}</time>
                </small>
              )}
            </li>
          ))}
        </ol>
      </details>
    </section>
  )
}
