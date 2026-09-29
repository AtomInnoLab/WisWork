import { useRef, useState, useSyncExternalStore } from 'react'
import type { PresentationQaRecord } from '../skills/powerpoint/presentation-qa.js'
import {
  validatePresentationQaAttempt,
  type PresentationQaAttempt,
} from '../skills/powerpoint/presentation-qa-attempts.js'
export interface PresentationQaController {
  read(): PresentationQaRecord | undefined
  attempts?(): PresentationQaAttempt[]
  closeAttempt?(expected: PresentationQaAttempt): Promise<PresentationQaAttempt>
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
  const [pendingClose, setPendingClose] = useState<PresentationQaAttempt>()
  const [closing, setClosing] = useState(false)
  const [closeError, setCloseError] = useState(false)
  const busy = useRef(false)
  let record: PresentationQaRecord | undefined
  let readError = false
  try {
    record = controller.read()
  } catch {
    readError = true
  }
  let attempts: PresentationQaAttempt[] = [],
    unavailable = false
  try {
    const value = controller.attempts?.() ?? []
    if (
      !Array.isArray(value) ||
      value.length > 64 ||
      new Set(value.map((attempt) => attempt?.id)).size !== value.length ||
      !value.every(validatePresentationQaAttempt)
    )
      throw Error('invalid_state')
    attempts = value
  } catch {
    unavailable = true
  }
  const attemptHistory = unavailable ? (
    <p role="alert">截图尝试历史暂不可读取；不沿用未知尝试结果，请刷新核对。</p>
  ) : attempts.length ? (
    <details>
      <summary>截图尝试记录 · {attempts.length} 条</summary>
      <p>只记录采集尝试和已保存截图；不代表图片已收到、视觉复核、专业事实或真实宿主验收通过。</p>
      <ol>
        {attempts.map((attempt) => (
          <li key={attempt.id}>
            <strong>
              页面 {attempt.pageId} ·{' '}
              {
                {
                  started: '开始记录尚未闭合，不能证明仍在执行，不自动重放',
                  recorded: '已保存截图，仍待视觉复核',
                  waiting: '等待截图能力；可在能力恢复后按明确请求重试',
                  failed: '采集未完成，请核对状态后再决定是否重试',
                  cancelled: '采集已取消，未认证截图或页面状态',
                  closed: '未决记录已结束；不代表宿主操作已取消或 QA 通过',
                }[attempt.status]
              }
            </strong>
            <p>
              <time dateTime={attempt.startedAt}>{attempt.startedAt}</time> · 原开始记录
            </p>
            {attempt.finishedAt && (
              <p>
                <time dateTime={attempt.finishedAt}>{attempt.finishedAt}</time> · 结束回执
              </p>
            )}
            {attempt.status === 'started' &&
              controller.closeAttempt &&
              (pendingClose?.id === attempt.id ? (
                <div>
                  <p>只结束这条未决记录，不会取消实际宿主操作，也不代表截图或 QA 通过。</p>
                  <button
                    type="button"
                    disabled={disabled || closing}
                    onClick={() => {
                      if (busy.current || disabled || !pendingClose) return
                      const expected = structuredClone(pendingClose)
                      if (JSON.stringify(expected) !== JSON.stringify(attempt)) {
                        setCloseError(true)
                        return
                      }
                      busy.current = true
                      setClosing(true)
                      setCloseError(false)
                      void controller.closeAttempt!(expected)
                        .then(
                          () => setPendingClose(undefined),
                          () => setCloseError(true),
                        )
                        .finally(() => {
                          busy.current = false
                          setClosing(false)
                        })
                    }}
                  >
                    确认结束记录
                  </button>
                  <button
                    type="button"
                    disabled={disabled || closing}
                    onClick={() => {
                      setPendingClose(undefined)
                      setCloseError(false)
                    }}
                  >
                    取消
                  </button>
                  {closeError && <p role="alert">记录未能确认结束；请刷新核对后再决定是否重试。</p>}
                </div>
              ) : (
                <button
                  type="button"
                  disabled={disabled || closing}
                  onClick={() => {
                    setPendingClose(structuredClone(attempt))
                    setCloseError(false)
                  }}
                >
                  结束未决记录
                </button>
              ))}
          </li>
        ))}
      </ol>
    </details>
  ) : null
  if (readError)
    return (
      <section className="presentation-project" aria-label="页面 QA 记录">
        <p role="alert">检查记录无法读取，请重新采集目标页面。</p>
        {attemptHistory}
      </section>
    )
  if (!record)
    return attemptHistory ? (
      <section className="presentation-project" aria-label="页面 QA 记录">
        {attemptHistory}
      </section>
    ) : null
  const affected = record.pages.filter((page) => page.recheckRequired).map((page) => page.pageId)
  return (
    <section className="presentation-project" aria-label="页面 QA 记录">
      <strong>
        {record.source === 'production' ? '页生产任务 QA 记录' : '页面 QA 记录'} ·{' '}
        {record.pages.length} 页
      </strong>
      <p>以下为历史检查记录，需重新采集才能确认当前状态；不代表来源核验或保存重开验收。</p>
      {attemptHistory}
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
