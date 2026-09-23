import { useSyncExternalStore } from 'react'
import type { PresentationProjectController } from '../skills/powerpoint/presentation-project.js'

export function PresentationProjectCard(props: {
  controller: PresentationProjectController
  disabled: boolean
}) {
  const { controller } = props
  const { phase, project, error } = useSyncExternalStore(
    (listener) => controller.subscribe(listener),
    () => controller.snapshot(),
    () => controller.snapshot(),
  )
  const active = phase !== 'idle'
  const disabled = props.disabled || active
  return (
    <section className="presentation-project" aria-label="演示文稿项目" aria-busy={active}>
      <strong>{project?.title ?? '演示文稿项目'}</strong>
      <p role="status">
        {active
          ? {
              loading: '正在读取项目',
              restoring: '正在恢复最近完成版本',
              resuming: '正在编译已保存版本',
            }[phase]
          : project
            ? `${project.slideCount} 页 · ${project.status === 'pending' ? '已保存，待编译' : '已编译，尚未完成视觉验证'}`
            : error
              ? ''
              : '当前文档暂无已保存项目'}
      </p>
      {error && (
        <p className="error-text" role="alert">
          {error}
        </p>
      )}
      <div className="presentation-project-actions">
        <button type="button" disabled={disabled} onClick={() => void controller.refresh()}>
          刷新
        </button>
        {project?.latestCompiledRequestId && (
          <button type="button" disabled={disabled} onClick={() => void controller.restore()}>
            恢复最近完成版本
          </button>
        )}
        {project?.status === 'pending' && (
          <button
            type="button"
            disabled={disabled}
            onClick={() => void controller.resume(project.latestRequestId)}
          >
            继续编译
          </button>
        )}
        {active && (
          <button type="button" onClick={() => controller.cancel()}>
            取消
          </button>
        )}
      </div>
      {project && (
        <details>
          <summary>页面、版本与检查</summary>
          <p>页面清单（最近保存版本）</p>
          <ol>
            {project.slides.map((slide) => (
              <li key={slide.id}>{slide.title}</li>
            ))}
          </ol>
          <p>保存版本（最近 20 次）</p>
          <ul>
            {project.history.map((entry) => (
              <li key={entry.requestId}>
                版本 {entry.sequence} · {entry.slideCount} 页 ·{' '}
                {entry.status === 'pending' ? '待编译' : '已编译'}
              </li>
            ))}
          </ul>
          {project.checks ? (
            <ul>
              <li>结构检查：通过</li>
              <li>几何检查：{project.checks.geometry === 'passed' ? '通过' : '有警告'}</li>
              <li>渲染检查：未执行</li>
              <li>来源：未核验</li>
              <li>Office 往返检查：未执行</li>
            </ul>
          ) : (
            <p>最近保存版本尚无编译检查结果。</p>
          )}
          <p>恢复将下载文件放回会话；导入当前文档仍需单独确认。</p>
        </details>
      )}
    </section>
  )
}
