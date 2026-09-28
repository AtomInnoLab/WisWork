import { PresentationDeliveryReportCard } from './presentation-delivery-report-card.js'
import { useSyncExternalStore } from 'react'
import type { PresentationProjectController } from '../skills/powerpoint/presentation-project.js'

export function PresentationProjectCard(props: {
  controller: PresentationProjectController
  disabled: boolean
}) {
  const { controller } = props
  const { phase, project, error, deliveryReport, deliveryNotice, sourceAudit } =
    useSyncExternalStore(
      (listener) => controller.subscribe(listener),
      () => controller.snapshot(),
      () => controller.snapshot(),
    )
  const job = project?.productionJob
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
              producing: '正在处理页任务',
              auditing: '正在核对计划引文与原文',
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
      {project?.commentsUnavailable && <p>本机审阅评论暂不可读取；项目与页面状态不受影响。</p>}
      {project?.reviewComments && (
        <details aria-label="本机审阅评论">
          <summary>
            本机审阅评论 · 待处理 {project.reviewComments.openCount} · 已解决{' '}
            {project.reviewComments.resolvedCount}
          </summary>
          <p>作者名称是未验证的显示标签；评论不代表来源、内容或视觉 QA 通过。</p>
          {project.reviewComments.recent.length <
            project.reviewComments.openCount + project.reviewComments.resolvedCount && (
            <p>只显示最近 8 条；可请 Agent 读取完整列表。</p>
          )}
          <ol>
            {project.reviewComments.recent.map((comment) => (
              <li key={comment.id}>
                {comment.state === 'open' ? '待处理' : '已解决'} ·{' '}
                {comment.targetKind === 'slide'
                  ? '页面'
                  : comment.targetKind === 'claim'
                    ? '主张'
                    : '来源'}{' '}
                {comment.targetId}
                {project.plan && comment.planRevision !== project.plan.revision
                  ? ` · 旧计划第 ${comment.planRevision} 版`
                  : ''}
                {' · '}
                {comment.authorLabel}：{comment.text}
              </li>
            ))}
          </ol>
        </details>
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
        {project?.production && project.production.status !== 'compiled' && !job && (
          <button
            type="button"
            disabled={disabled}
            onClick={() => void controller.runProduction(project.production!.requestId)}
          >
            继续页任务
          </button>
        )}
        {active && <p>取消仅停止当前等待，不会取消 PC 后台任务。</p>}
        {active && (
          <button type="button" onClick={() => controller.cancel()}>
            取消
          </button>
        )}
      </div>
      {project?.production &&
        !project.jobsUnavailable &&
        !job &&
        project.production.status !== 'compiled' && (
          <button
            type="button"
            disabled={disabled}
            onClick={() => void controller.startProductionJob(project.production!.requestId)}
          >
            后台制作剩余页面
          </button>
        )}
      {project?.productionTasks && project.productionTasks.length > 0 && (
        <label>
          选择页生产任务
          <select
            aria-label="选择页生产任务"
            disabled={disabled || project.jobsUnavailable}
            value={project.production?.requestId ?? ''}
            onChange={(event) => void controller.selectProduction(event.target.value)}
          >
            {project.productionTasks.map((task) => (
              <option key={task.requestId} value={task.requestId}>
                任务 {task.sequence} · {task.requestId} · 计划第 {task.planRevision} 版 ·{' '}
                {task.compiledCount}/{task.total} 页 ·{' '}
                {task.jobState
                  ? {
                      running: '制作中',
                      pausing: '等待暂停',
                      paused: '已暂停',
                      cancelling: '等待取消',
                      cancelled: '已取消',
                      interrupted: '已中断',
                      completed: '编译完成',
                      failed: '失败待继续',
                    }[task.jobState]
                  : {
                      pending: '待制作',
                      building: '制作中',
                      partial: '部分完成',
                      compiled: '编译完成',
                    }[task.status]}
              </option>
            ))}
          </select>
        </label>
      )}
      {project?.jobsUnavailable && <p>当前 PC 不支持后台任务，请升级；仍可使用继续页任务。</p>}
      {job && project?.production && (
        <section aria-label="后台生产任务">
          <p role="status">
            后台页编译 ·{' '}
            {
              {
                running: '制作中',
                pausing: '正在编译的页面完成后暂停',
                paused: '已暂停',
                cancelling: '当前页完成后取消',
                cancelled: '已取消，成果保留',
                interrupted: 'PC 运行已中断',
                completed: '编译完成',
                failed: '失败待继续',
              }[job.state]
            }
          </p>
          <p>
            下一步：
            {['paused', 'interrupted', 'failed'].includes(job.state)
              ? '继续后台制作，已完成页会保留。'
              : job.state === 'completed'
                ? project.production.revision
                  ? '保存修订单页检查，再单独确认宿主页替换；不能整批追加导入。'
                  : '保存成果并准备导入，视觉与来源仍需验收。'
                : job.state === 'cancelled'
                  ? '保存已完成单页；此任务不会重新启动。'
                  : '可离开当前面板；暂停和取消在正在编译的页面完成后生效。'}
          </p>
          <div className="presentation-project-actions">
            {job.state === 'running' && (
              <button
                type="button"
                disabled={disabled}
                onClick={() => void controller.pauseProductionJob(job.requestId)}
              >
                暂停后台任务
              </button>
            )}
            {['paused', 'interrupted', 'failed'].includes(job.state) && (
              <button
                type="button"
                disabled={disabled}
                onClick={() => void controller.resumeProductionJob(job.requestId)}
              >
                继续后台任务
              </button>
            )}
            {!['completed', 'cancelled', 'cancelling'].includes(job.state) && (
              <button
                type="button"
                disabled={disabled}
                onClick={() => void controller.cancelProductionJob(job.requestId)}
              >
                取消后台任务
              </button>
            )}
          </div>
          <details className="presentation-job-events">
            <summary>生产事件 · 最近 {job.events.length} 条</summary>
            {job.revision > job.events.length && <p>更早历史已截断，此处不是完整审计记录。</p>}
            <p>仅记录冻结页面编译，不包含研究、宿主写入或 QA。</p>
            <ol>
              {job.events.map((event) => (
                <li key={event.sequence}>
                  <time dateTime={event.createdAt}>{event.createdAt}</time> ·{' '}
                  {
                    {
                      'run.started': '开始后台制作',
                      'run.pause_requested': '已请求暂停',
                      'run.paused': '已暂停',
                      'run.cancel_requested': '已请求取消',
                      'run.cancelled': '已取消',
                      'run.interrupted': '运行已中断',
                      'run.completed': '编译完成',
                      'run.failed': '运行失败',
                      'page.started': '页面开始编译',
                      'page.compiled': '页面编译完成',
                      'page.failed': '页面编译失败',
                    }[event.type]
                  }
                  {'pageId' in event
                    ? ` · 页面 ${project.production?.pages.find((page) => page.id === event.pageId)?.title ?? '未知页面'}`
                    : ''}
                  {'attempt' in event ? ` · 尝试 ${event.attempt}` : ''}
                  {'error' in event && event.error
                    ? ` · ${{ compile_failed: '编译失败', invalid_deck: '页面内容无效', aborted: '已停止', output_too_large: '成果过大', asset_unavailable: '素材不可用', source_unavailable: '附件不可读或来源摘录未匹配', font_unavailable: '指定字体及回退字体均不可用', invalid_state: '任务状态异常' }[event.error]}`
                    : ''}
                </li>
              ))}
            </ol>
          </details>
        </section>
      )}
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
              。尚未替换当前页；可下载单页检查，宿主页替换需单独确认，禁止整批追加导入。
            </p>
          )}
          {project.plan && project.production.planRevision !== project.plan.revision && (
            <p>页任务使用旧计划，继续任务按原快照，不代表当前计划。</p>
          )}
          {project.production.status === 'compiled' && !project.production.revision && (
            <button
              type="button"
              disabled={disabled}
              onClick={() => void controller.prepareProduction()}
            >
              准备完整成果导入
            </button>
          )}
          {project.production.status === 'compiled' &&
            controller.pdfAvailable?.() &&
            controller.exportProductionPdf && (
              <button
                type="button"
                disabled={disabled}
                onClick={() => void controller.exportProductionPdf?.()}
              >
                导出 PDF 预览
              </button>
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
                  ? ` · ${{ compile_failed: '编译失败', invalid_deck: '页面内容无效', aborted: '已停止', output_too_large: '成果过大', asset_unavailable: '素材不可用', source_unavailable: '附件不可读或来源摘录未匹配', font_unavailable: '指定字体及回退字体均不可用' }[page.error]}`
                  : ''}
                {page.state === 'compiled' && (
                  <button
                    type="button"
                    disabled={disabled}
                    onClick={() => void controller.downloadProductionPage(page.id)}
                  >
                    保存单页到附件：{page.title}
                  </button>
                )}
              </li>
            ))}
          </ol>
        </section>
      )}
      {project?.production && (
        <div className="presentation-project-actions">
          <button
            type="button"
            disabled={disabled}
            onClick={() => void controller.readDeliveryReport()}
          >
            读取内容证据报告
          </button>
          <button
            type="button"
            disabled={disabled}
            onClick={() => void controller.exportDeliveryReport()}
          >
            导出证据 JSON + Markdown 到附件
          </button>
        </div>
      )}
      {deliveryNotice && <p role="status">{deliveryNotice}</p>}
      {deliveryReport && (
        <PresentationDeliveryReportCard
          key={`${deliveryReport.documentId}-${deliveryReport.projectId}-${deliveryReport.requestId}`}
          report={deliveryReport}
          controller={controller}
          disabled={disabled}
        />
      )}
      {project?.sourcePreparationUnavailable && (
        <p role="alert">附件来源状态暂不可读取；请刷新项目，再核对资料是否已上传并解析。</p>
      )}
      {project?.sourcePreparation && project.sourcePreparation.length > 0 && (
        <section aria-label="附件来源准备状态">
          <p role="status">
            附件来源{' '}
            {
              project.sourcePreparation.filter((source) => source.status === 'excerpt_matched')
                .length
            }{' '}
            / {project.sourcePreparation.length} 已就绪
          </p>
          {project.sourcePreparation.some((source) => source.status !== 'excerpt_matched') && (
            <p>请处理未确认的来源：补齐资料、修订计划摘录，或升级旧版 PC 后核对。</p>
          )}
          {project.production?.pages.some((page) => page.error === 'source_unavailable') && (
            <p>附件已就绪但摘录未匹配时，请核对附件原文，修订计划摘录后启动新任务。</p>
          )}
          <ul>
            {project.sourcePreparation.map((source) => (
              <li key={source.sourceId}>
                {project.plan?.value.sources.find((item) => item.id === source.sourceId)?.title ??
                  source.sourceId}
                ：
                {
                  {
                    ready: '旧版 PC 仅确认已解析，摘录未核对',
                    excerpt_matched: '摘录已匹配，编译前仍会复核',
                    uploading: '上传中',
                    failed: '解析失败',
                    missing: '附件缺失',
                    unsupported: '不是可读文本',
                    excerpt_mismatch: '摘录不在附件原文中',
                    excerpt_missing: '计划摘录为空',
                    source_mismatch: '网页快照与计划网址不匹配',
                  }[source.status]
                }
              </li>
            ))}
          </ul>
          <p>
            就绪表示当前附件的计划摘录逐字匹配原文；编译前会复核。它不代表引文准确、事实可靠或信息仍然有效。
          </p>
        </section>
      )}
      {project?.plan && (
        <details>
          <summary>制作计划 · 第 {project.plan.revision} 版</summary>
          {project.plan.value.sources.some((source) =>
            /^attachment:[a-f0-9]{64}$/.test(source.uri),
          ) && (
            <section aria-label="计划来源原文核对">
              <button
                type="button"
                disabled={disabled}
                onClick={() => void controller.auditSources()}
              >
                核对计划引文与附件原文
              </button>
              <p>只检查引文是否逐字出现在当前文档附件中，不核验事实、适用范围或时效。</p>
              {sourceAudit?.planRevision === project.plan.revision && (
                <div role="status">
                  <p>
                    第 {sourceAudit.planRevision} 版来源原文核对：
                    {sourceAudit.sources.filter((source) => source.status === 'found').length}/
                    {sourceAudit.sources.length} 份找到字面匹配。
                  </p>
                  <ul>
                    {sourceAudit.sources.map((source) => (
                      <li key={source.sourceId}>
                        {project.plan!.value.sources.find((item) => item.id === source.sourceId)
                          ?.title ?? source.sourceId}
                        ：
                        {
                          {
                            found: '原文中找到引文',
                            not_found: '未在完整原文中找到引文',
                            empty_excerpt: '计划未填写引文',
                            not_ready: '附件尚未解析就绪',
                            unsupported: '附件不是可读文本',
                            missing: '当前文档缺少附件',
                          }[source.status]
                        }
                        {source.status === 'found' ? ` · UTF-16 位置 ${source.offset}` : ''}
                      </li>
                    ))}
                  </ul>
                </div>
              )}
            </section>
          )}
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
