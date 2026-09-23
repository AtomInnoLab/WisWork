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
              producing: '正在编译剩余页面',
            }[phase]
          : project
            ? `${project.slideCount} 页 · ${project.status === 'planned' ? '计划已保存，尚未编译' : project.status === 'pending' ? '已保存，待编译' : '已编译，尚未完成视觉验证'}`
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
        {project?.status === 'pending' && project.latestRequestId && (
          <button
            type="button"
            disabled={disabled}
            onClick={() => void controller.resume(project.latestRequestId!)}
          >
            继续编译
          </button>
        )}
        {project?.production && project.production.status !== 'compiled' && (
          <button
            type="button"
            disabled={disabled}
            onClick={() => void controller.runProduction(project.production!.requestId)}
          >
            继续页任务
          </button>
        )}
        {active && (
          <button type="button" onClick={() => controller.cancel()}>
            取消
          </button>
        )}
      </div>
      {project?.production && (
        <section aria-label="逐页生产进度">
          <p>
            已编译 {project.production.compiledCount} / {project.production.total} 页 · 计划第{' '}
            {project.production.planRevision} 版
          </p>
          <p>单页编译成果尚未导入或验收，不代表整套交付。</p>
          {project.production.revision && (
            <p>
              单页修订 · 父请求：{project.production.revision.parentRequestId} · 目标页：
              {project.production.revision.pageId}
              。尚未替换当前页；可下载单页检查，宿主替换待接入，禁止整批追加导入。
            </p>
          )}
          {project.plan && project.production.planRevision !== project.plan.revision && (
            <p>页任务使用旧计划，继续任务按原快照，不代表当前计划。</p>
          )}
          <ol>
            {project.production.pages.map((page) => (
              <li key={page.id}>
                {page.title} ·{' '}
                {
                  {
                    pending: '待制作',
                    building: '制作中',
                    failed: '失败待重试',
                    compiled: '已编译（未导入验收）',
                  }[page.state]
                }{' '}
                · 尝试 {page.attempt} 次
                {page.error
                  ? ` · ${{ compile_failed: '编译失败', invalid_deck: '页面内容无效', aborted: '已停止', output_too_large: '成果过大', asset_unavailable: '素材不可用' }[page.error]}`
                  : ''}
              </li>
            ))}
          </ol>
        </section>
      )}
      {project?.plan && (
        <details>
          <summary>制作计划 · 第 {project.plan.revision} 版</summary>
          <p>{project.plan.value.title}</p>
          <p>目标：{project.plan.value.brief.objective}</p>
          <p>
            受众：{project.plan.value.brief.audience} · {project.plan.value.brief.minutes} 分钟 ·{' '}
            {project.plan.value.brief.language}
          </p>
          <p>
            字体：{project.plan.value.style.fontFace} · 主色 #{project.plan.value.style.accentColor}
          </p>
          <p>
            {project.plan.value.sources.length} 个来源 · {project.plan.value.claims.length}{' '}
            条主张（未核验）
          </p>
          {project.status !== 'planned' && (
            <p>
              {project.requestPlanRevision === project.plan.revision
                ? '最近编译请求已关联当前计划；内容与视觉仍需验收。'
                : '计划已有更新或尚未关联：现有编译成果不代表当前计划。'}
            </p>
          )}
          <ol>
            {project.plan.value.slides.map((slide) => (
              <li key={slide.id}>
                {slide.title}：{slide.purpose}
              </li>
            ))}
          </ol>
          <details>
            <summary>来源与主张</summary>
            <ul>
              {project.plan.value.sources.map((source) => (
                <li key={source.id}>
                  {source.title} · {source.uri}
                  {source.locator ? ` · ${source.locator}` : ''}
                  <p>{source.excerpt}</p>
                </li>
              ))}
            </ul>
            <ul>
              {project.plan.value.claims.map((claim) => (
                <li key={claim.id}>
                  {claim.statement} · {claim.type} · 待核验
                </li>
              ))}
            </ul>
          </details>
        </details>
      )}
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
