import { useEffect, useState } from 'react'
import type { PresentationAcquisitionHistory } from '@wiswork/project-store/presentation-acquisition'

const failureReasons = {
  remote_image_unavailable: '图片暂时无法获取',
  remote_webpage_unavailable: '网页暂时无法获取',
  quota_exceeded: '文档资料容量已满',
  parse_failed: '资料无法解析',
  digest_mismatch: '资料完整性校验未通过',
  animated_image_unsupported: '动画需要明确选择静态首帧',
  remote_image_source_conflict: '图片与已有资料来源冲突',
  remote_webpage_source_conflict: '网页与已有资料来源冲突',
  aborted: '获取已中断',
  invalid_state: '获取记录状态异常',
  acquisition_failed: '资料获取失败',
} as const

export function PresentationAcquisitionHistoryCard({
  read,
  refreshKey,
}: {
  read?: () => Promise<PresentationAcquisitionHistory | undefined>
  refreshKey?: unknown
}) {
  const [history, setHistory] = useState<PresentationAcquisitionHistory>()
  const [error, setError] = useState(false)
  const [refresh, setRefresh] = useState(0)
  useEffect(() => {
    let current = true
    setHistory(undefined)
    setError(false)
    if (read)
      void read()
        .then((value) => {
          if (current) setHistory(value)
        })
        .catch(() => {
          if (current) setError(true)
        })
    return () => {
      current = false
    }
  }, [read, refreshKey, refresh])
  if (!history && !error) return null
  const groups = new Map<string, NonNullable<typeof history>['records']>()
  for (const record of history?.records ?? []) {
    const key = `${record.kind}:${record.sourceUrlHash}`
    const group = groups.get(key) ?? []
    group.push(record)
    groups.set(key, group)
  }
  const status = (state: string) =>
    state === 'ready'
      ? '已获取本机产物'
      : state === 'rejected'
        ? '获取失败'
        : '缺少完成回执（不表示后台仍在运行）'
  return (
    <details>
      <summary>网络资料获取历史</summary>
      <button type="button" onClick={() => setRefresh((value) => value + 1)}>
        刷新获取历史
      </button>
      {error && <p role="status">获取历史暂时无法读取，请稍后刷新。</p>}
      {history && (
        <>
          <p>已获取不代表事实、许可或 QA 验收通过。</p>
          <p>记录时间是获取操作时间；缓存复用也会记录，不代表来源被重新抓取或更新。</p>
          <p>
            累计 {history.totalAttempts} 次尝试；显示最近 {history.records.length} 次
            {history.totalAttempts > history.records.length ? '，更早记录已超出保留窗口' : ''}。
          </p>
          {history.records.length === 0 && <p>当前文档暂无网络资料获取记录。</p>}
          {[...groups].map(([key, records]) => {
            const latest = records.at(-1)!
            return (
              <section key={key} data-acquisition-source={key}>
                <p>
                  {latest.kind === 'image' ? '图片' : '网页'}：{latest.source}
                </p>
                <p>
                  最新：尝试 #{latest.attempt} · {status(latest.state)}
                </p>
                <details>
                  <summary>每次尝试与时间</summary>
                  <ol>
                    {[...records].reverse().map((record) => (
                      <li key={record.id}>
                        尝试 #{record.attempt} · {status(record.state)}
                        <p>
                          开始：<time dateTime={record.startedAt}>{record.startedAt}</time>
                        </p>
                        {record.state !== 'fetching' && (
                          <p>
                            结束：<time dateTime={record.finishedAt}>{record.finishedAt}</time>
                          </p>
                        )}
                        {record.state !== 'fetching' && record.attachmentId && (
                          <details>
                            <summary>产物技术详情</summary>
                            <p>附件 ID：{record.attachmentId}</p>
                          </details>
                        )}
                        {record.state === 'rejected' && (
                          <p>{failureReasons[record.error]}；可检查来源后显式重试。</p>
                        )}
                      </li>
                    ))}
                  </ol>
                </details>
              </section>
            )
          })}
        </>
      )}
    </details>
  )
}
