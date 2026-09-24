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
        变更历史最多 64 条，达到容量上限后停止新增；旧版本仅有最近保存点，不完整历史的先后顺序未知。
        每次撤销仍检查当前对象；对象 ID 或内容已变化时可能无法继续撤销。图片恢复后对象 ID 会改变。
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
        <p>当前文档或任务暂无可用保存点。</p>
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
              {entry.source === 'existing_batch'
                ? '现稿批量 · '
                : entry.source === 'existing'
                  ? '现稿 · '
                  : ''}
              页面：{entry.pageId} · 状态：
              {entry.state}
            </p>
            {entry.sequence !== undefined && (
              <p>{entry.legacy ? '旧保存点 · 顺序未知' : `记录 #${entry.sequence}`}</p>
            )}
            {entry.changeSet && (
              <div>
                <p>
                  操作：{entry.changeSet.intent} · 风险：
                  {entry.changeSet.risk === 'high' ? '高' : '中'}
                </p>
                <p>
                  范围：{entry.changeSet.scope.slideIds.length} 页
                  {entry.changeSet.scope.shapeIds
                    ? ` · ${entry.changeSet.scope.shapeIds.length} 个对象`
                    : ''}
                </p>
                <p>验收要求（不代表已通过）：{entry.changeSet.validation.join('、')}</p>
              </div>
            )}
            {entry.review && (
              <p>
                历史截图复核：{entry.review.status === 'pass' ? '通过' : '未通过'} ·{' '}
                {entry.review.reviewedAt}。此结果不代表当前页面 QA 通过，需重新采集截图确认。
              </p>
            )}
            {entry.source === 'existing_batch' && (
              <div>
                <p>
                  历史截图复核：{entry.reviews?.length ?? 0}/{entry.affectedPageCount ?? 0}{' '}
                  个受影响页面。此结果不代表当前页面 QA 通过。
                </p>
                {entry.reviews?.map((review) => (
                  <p key={review.hostSlideId}>
                    {review.hostSlideId}：{review.status === 'pass' ? '通过' : '未通过'} ·{' '}
                    {review.reviewedAt}
                  </p>
                ))}
              </div>
            )}
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
