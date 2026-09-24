import { useEffect, useState } from 'react'
import type {
  PresentationChangeAction,
  PresentationChangesController,
} from './presentation-changes.js'
const labels: Record<PresentationChangeAction, string> = {
  inspect: '检查',
  undo: '撤销',
  resume: '继续',
  commit: '提交替换',
  discard: '丢弃替换',
}
export function PresentationChangesCard({
  controller,
  disabled = false,
}: {
  controller: PresentationChangesController
  disabled?: boolean
}) {
  const [snapshot, setSnapshot] = useState(() => controller.snapshot())
  useEffect(() => {
    const update = () => setSnapshot(controller.snapshot())
    const unsubscribe = controller.subscribe(update)
    update()
    return unsubscribe
  }, [controller])
  const busy = disabled || snapshot.phase !== 'idle'
  return (
    <section className="presentation-project" aria-label="修改保存点工作台">
      <strong>修改差异与撤销</strong>
      <p>
        最近保存点 / 不完整历史：文字、几何和整页各保留最近一条，图片最多 32
        条；不支持跨类型多级撤销，图片暂无撤销。
      </p>
      {snapshot.projectId && (
        <p>
          当前项目：{snapshot.projectId} · 任务：{snapshot.requestId}
        </p>
      )}
      <button type="button" disabled={busy} onClick={() => void controller.refresh()}>
        刷新保存点
      </button>
      {snapshot.phase === 'loading' && <p role="status">正在读取保存点…</p>}
      {snapshot.phase === 'acting' && <p role="status">正在检查并准备操作…</p>}
      {snapshot.notice && <p role="status">{snapshot.notice}</p>}
      {snapshot.error && <p role="alert">{snapshot.error}</p>}
      {snapshot.phase === 'idle' && !snapshot.entries.length && (
        <p>{snapshot.projectId ? '当前任务暂无保存点。' : '恢复当前任务后可查看保存点。'}</p>
      )}
      <ol>
        {snapshot.entries.map((entry) => (
          <li key={entry.id}>
            <strong>
              {
                {
                  text: '文字差异',
                  geometry: '几何差异（pt）',
                  image: '图片身份与摘要差异（非视觉 diff）',
                  page: '整页身份与摘要差异（非视觉 diff）',
                }[entry.kind]
              }
            </strong>
            <p>
              页面：{entry.pageId} · 状态：{entry.state}
            </p>
            <details>
              <summary>展开修改前后内容（长内容已折叠）</summary>
              <p>修改前</p>
              <pre style={{ whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' }}>{entry.before}</pre>
              <p>修改后</p>
              <pre style={{ whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' }}>{entry.after}</pre>
            </details>
            {entry.actions.map((action) => (
              <button
                key={action}
                type="button"
                disabled={busy}
                aria-label={`${labels[action]} ${entry.pageId}`}
                onClick={() => void controller.run(entry.id, action)}
              >
                {labels[action]}
              </button>
            ))}
          </li>
        ))}
      </ol>
      <p>写入操作仍需提案确认；变更后需重新采集页面 QA。</p>
    </section>
  )
}
