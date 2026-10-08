import type { ReactNode } from 'react'
import { acpPresentationStageLabel } from '@wiswork/agent-harness'
import type { PresentationStageGroup } from './presentation-stage-timeline.js'

export function PresentationStageCard({
  group,
  runActive = true,
  children,
}: {
  group: PresentationStageGroup
  runActive?: boolean
  children?: ReactNode
}) {
  const label = acpPresentationStageLabel(group.stage)
  return (
    <article
      className={`tool-event stage-event${group.failed ? ' tool-error' : ''}`}
      aria-label={`${label}阶段`}
      aria-busy={runActive && group.running > 0}
    >
      <strong>{label}</strong>
      <p role="status">
        {runActive ? '进行中' : '待核对'} {group.running} · 已结束 {group.ended} · 未完成{' '}
        {group.failed} 项操作
      </p>
      {group.failed > 0 && (
        <p role="alert">操作记录中有失败。是否仍需处理以当前项目与恢复记录为准。</p>
      )}
      {group.running > 0 ? (
        runActive ? (
          <p>当前操作仍在进行，等待结果；可结束前台并保留已保存成果。</p>
        ) : (
          <p role="alert">尚未收到最终回执。请核对项目与恢复记录；前台停止不表示宿主写入已撤回。</p>
        )
      ) : group.failed === 0 ? (
        <p>阶段操作已结束，页面与整套交付以项目记录为准。</p>
      ) : null}
      <details>
        <summary>内部操作 · {group.events.length} 项</summary>
        {children ?? (
          <ul>
            {group.events.map((event) => (
              <li key={event.id}>
                <code>{event.name}</code> ·{' '}
                {event.state === 'running'
                  ? runActive
                    ? '进行中'
                    : '待核对'
                  : event.state === 'error'
                    ? '未完成'
                    : '已结束'}
              </li>
            ))}
          </ul>
        )}
      </details>
    </article>
  )
}
