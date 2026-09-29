import { presentationProductionEventRows } from './presentation-workflow.js'
import { PRESENTATION_WIDTH, PRESENTATION_HEIGHT } from '@wiswork/pptx-engine/presentation'
import { PresentationDeliveryReportCard } from './presentation-delivery-report-card.js'
import { useSyncExternalStore } from 'react'
import type { PresentationProjectController } from '../skills/powerpoint/presentation-project.js'

export function PresentationProjectCard(props: {
  controller: PresentationProjectController
  disabled: boolean
  onEndFrontend?: () => void
}) {
  const { controller } = props
  const { phase, project, error, deliveryReport, deliveryNotice, sourceAudit } =
    useSyncExternalStore(
      (listener) => controller.subscribe(listener),
      () => controller.snapshot(),
      () => controller.snapshot(),
    )
  const productionEvents = presentationProductionEventRows(project)
  const job = project?.productionJob
  const active = phase !== 'idle'
  const disabled = props.disabled || active
  const production = project?.production
  const knownJob =
    !project?.jobsUnavailable &&
    !!job &&
    !!production &&
    [
      'running',
      'pausing',
      'paused',
      'cancelling',
      'cancelled',
      'interrupted',
      'completed',
      'failed',
    ].includes(job.state) &&
    job.requestId === production.requestId &&
    job.projectId === project?.projectId &&
    production.projectId === project.projectId
  const running = knownJob && job.state === 'running'
  const resumable = knownJob && ['paused', 'interrupted', 'failed'].includes(job.state)
  const startable =
    !project?.jobsUnavailable &&
    job === null &&
    !!production &&
    production.projectId === project?.projectId &&
    ['pending', 'partial'].includes(production.status)
  const cancellable =
    knownJob && ['running', 'pausing', 'paused', 'interrupted', 'failed'].includes(job.state)
  const endFrontend = () => {
    if (!props.onEndFrontend) return
    if (active) controller.cancel()
    props.onEndFrontend()
  }
  return (
    <section className="presentation-project" aria-label="演示文稿项目" aria-busy={active}>
      <strong>{project?.title ?? '演示文稿项目'}</strong>
      <details aria-label="完成操作">
        <summary>完成</summary>
        <p>结束前台执行与等待；已保存项目和页面保留。正在提交的修改请核对恢复记录。</p>
        <div className="presentation-project-actions">
          <button type="button" disabled={!props.onEndFrontend} onClick={endFrontend}>
            保留成果并结束前台
          </button>
          {(running || resumable || startable) && (
            <button
              type="button"
              disabled={!props.onEndFrontend || (!running && disabled)}
              onClick={() => {
                endFrontend()
                if (running || disabled) return
                if (resumable) void controller.resumeProductionJob(job!.requestId)
                else if (startable) void controller.startProductionJob(production!.requestId)
              }}
            >
              继续后台制作并结束前台
            </button>
          )}
          {cancellable && (
            <button
              type="button"
              disabled={disabled}
              onClick={() => void controller.cancelProductionJob(job!.requestId)}
            >
              取消剩余后台页面
            </button>
          )}
        </div>
        {running && <p>当前任务已在后台运行；结束前台不会重新启动任务。</p>}
        {cancellable && <p>取消剩余后台页面在当前编译页结束后生效，已完成成果保留。</p>}
        {!running && !resumable && !startable && (
          <p>此处仅结束前台；当前没有可确认继续的后台任务。</p>
        )}
      </details>
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
      {startable && (
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
      {knownJob && job && project?.production && (
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
            <summary>
              生产事件 · 最近 {productionEvents?.retainedEventCount ?? 0} 条（按页合并）
            </summary>
            {productionEvents?.truncated && (
              <p>更早历史已截断，此处不是完整审计记录或失败次数统计。</p>
            )}
            <p>仅记录冻结页面编译，不包含研究、宿主写入或 QA。页内详情只含当前保留的尝试事件。</p>
            <ol>
              {productionEvents?.rows.map((row) => (
                <li key={row.id} data-page-id={row.pageId}>
                  {row.attempts ? (
                    <details>
                      <summary>
                        <time dateTime={row.at}>{row.at}</time> · {row.text}
                      </summary>
                      <ol>
                        {row.attempts.map((attempt) => (
                          <li key={attempt.id}>
                            <time dateTime={attempt.at}>{attempt.at}</time> · {attempt.text}
                          </li>
                        ))}
                      </ol>
                    </details>
                  ) : (
                    <>
                      <time dateTime={row.at}>{row.at}</time> · {row.text}
                    </>
                  )}
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
                    locator_mismatch: '计划定位与原文实际位置不匹配',
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
          <details aria-label="样式规范">
            <summary>样式规范</summary>
            <p>
              背景 #{project.plan.value.style.background} · 正文 #
              {project.plan.value.style.textColor}
            </p>
            <p>备用字体：{project.plan.value.style.fontFallbacks?.join('、') || '未指定'}</p>
            {project.plan.value.brandKit && (
              <>
                <p>
                  品牌：{project.plan.value.brandKit.name} · 第{' '}
                  {project.plan.value.brandKit.revision} 版
                </p>
                <p>
                  品牌色板：
                  {project.plan.value.brandKit.allowedColors.map((color) => `#${color}`).join('、')}
                </p>
                <p>
                  标志：
                  {project.plan.value.brandKit.logo
                    ? project.plan.value.brandKit.logo.placement === 'all'
                      ? '每页使用'
                      : '封面使用'
                    : '未指定'}
                </p>
                <ul>
                  {project.plan.value.brandKit.layoutComponents?.map((component) => (
                    <li key={component.id}>
                      {component.name} · {component.slots.length} 个布局区域
                    </li>
                  ))}
                </ul>
                {project.plan.value.brandKit.layoutComponents?.slice(0, 3).map((component) => (
                  <figure key={component.id}>
                    <svg
                      role="img"
                      aria-label={`${component.name}布局示意`}
                      viewBox={`0 0 ${PRESENTATION_WIDTH} ${PRESENTATION_HEIGHT}`}
                      width="100%"
                      style={{ maxWidth: 240 }}
                    >
                      <rect
                        width={PRESENTATION_WIDTH}
                        height={PRESENTATION_HEIGHT}
                        fill="none"
                        stroke="currentColor"
                        strokeWidth={0.04}
                      />
                      {component.slots.map((slot) => (
                        <g key={slot.id}>
                          <rect
                            x={slot.x}
                            y={slot.y}
                            width={slot.w}
                            height={slot.h}
                            fill="none"
                            stroke="currentColor"
                            strokeWidth={0.04}
                          />
                          <text
                            x={slot.x + 0.06}
                            y={slot.y + 0.25}
                            fill="currentColor"
                            fontSize={0.2}
                          >
                            {
                              {
                                text: '文本',
                                shape: '形状',
                                image: '图片',
                                table: '表格',
                                chart: '图表',
                              }[slot.kind]
                            }
                          </text>
                        </g>
                      ))}
                    </svg>
                    <figcaption>{component.name} · 布局示意，非成品预览</figcaption>
                  </figure>
                ))}
              </>
            )}
            <p>此处展示已保存的样式选择；实际字体、素材和布局效果仍需视觉检查。</p>
          </details>
          {project.plan.value.brief.requiredContent.length > 0 && (
            <p>必含内容：{project.plan.value.brief.requiredContent.join('；')}</p>
          )}
          {project.plan.value.brief.constraints.length > 0 && (
            <p>约束：{project.plan.value.brief.constraints.join('；')}</p>
          )}
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
          <ol aria-label="逐页施工图">
            {project.plan.value.slides.map((slide) => (
              <li key={slide.id}>
                <strong>{slide.title}</strong>
                <p>用途：{slide.purpose}</p>
                <p>
                  页型：
                  {
                    {
                      cover: '封面',
                      content: '内容',
                      comparison: '对比',
                      process: '流程',
                      chart: '图表',
                      summary: '总结',
                    }[slide.layout]
                  }
                </p>
                <p>页面结论与主张（待核验）：</p>
                {slide.claimIds.length === 0 ? (
                  <p>尚未登记主张</p>
                ) : (
                  <ul>
                    {slide.claimIds.map((id) => {
                      const claim = project.plan!.value.claims.find((item) => item.id === id)
                      return (
                        <li key={id}>
                          {claim?.statement ?? `未找到主张 ${id}`}
                          {claim && (
                            <p>
                              依据：
                              {claim.sourceIds
                                .map(
                                  (sourceId) =>
                                    project.plan!.value.sources.find(
                                      (source) => source.id === sourceId,
                                    )?.title ?? sourceId,
                                )
                                .join('、') || '未登记来源'}{' '}
                              · 待核验
                            </p>
                          )}
                        </li>
                      )
                    })}
                  </ul>
                )}
                <p>所需素材：{slide.requiredAssets.join('、') || '未指定'}</p>
                {slide.dependsOn && slide.dependsOn.length > 0 && (
                  <p>
                    依赖页面：
                    {slide.dependsOn
                      .map(
                        (id) =>
                          project.plan!.value.slides.find((item) => item.id === id)?.title ?? id,
                      )
                      .join('、')}
                  </p>
                )}
                {slide.layoutComponentId && (
                  <p>
                    布局组件：
                    {project.plan!.value.brandKit?.layoutComponents?.find(
                      (item) => item.id === slide.layoutComponentId,
                    )?.name ?? slide.layoutComponentId}
                  </p>
                )}
                <details>
                  <summary>本页验收要求</summary>
                  <ul>
                    {slide.acceptanceCriteria.map((criterion, index) => (
                      <li key={index}>{criterion}</li>
                    ))}
                  </ul>
                </details>
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
