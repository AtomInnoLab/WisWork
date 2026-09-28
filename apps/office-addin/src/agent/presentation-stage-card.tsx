import { acpPresentationStageLabel } from '@wiswork/agent-harness'
import type { PresentationStageGroup } from './presentation-stage-timeline.js'

export function PresentationStageCard({ group }: { group: PresentationStageGroup }) {
  const label = acpPresentationStageLabel(group.stage)
  return (
    <article
      className={`tool-event stage-event${group.failed ? ' tool-error' : ''}`}
      aria-label={`${label}阶段`}
      aria-busy={group.running > 0}
    >
      <strong>{label}</strong>
      <p role="status">
        进行中 {group.running} · 已结束 {group.ended} · 未完成 {group.failed} 项操作
      </p>
      {group.failed > 0 && (
        <p role="alert">操作记录中有失败。是否仍需处理以当前项目与恢复记录为准。</p>
      )}
      {group.running > 0 ? (
        <p>当前操作仍在进行，等待结果；可结束前台并保留已保存成果。</p>
      ) : group.failed === 0 ? (
        <p>阶段操作已结束，页面与整套交付以项目记录为准。</p>
      ) : null}
      <details>
        <summary>内部操作 · {group.events.length} 项</summary>
        <ul>
          {group.events.map((event) => (
            <li key={event.id}>
              <code>{event.name}</code> ·{' '}
              {event.state === 'running' ? '进行中' : event.state === 'error' ? '未完成' : '已结束'}
            </li>
          ))}
        </ul>
      </details>
    </article>
  )
}
