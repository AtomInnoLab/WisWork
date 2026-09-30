import type { PresentationFeedbackComparison } from '@wiswork/pptx-engine/presentation-feedback-comparison'
import { downloadLocalFile } from './session-download.js'
import type {
  PresentationProductionFeedbackLedger,
  PresentationProductionFeedbackPage,
} from '@wiswork/project-store/presentation-feedback'
import { presentationProductionEventRows } from './presentation-workflow.js'
import { PRESENTATION_WIDTH, PRESENTATION_HEIGHT } from '@wiswork/pptx-engine/presentation'
import { PresentationDeliveryReportCard } from './presentation-delivery-report-card.js'
import { useState, useSyncExternalStore } from 'react'
import type {
  PresentationProjectController,
  PresentationProjectStatus,
} from '../skills/powerpoint/presentation-project.js'

function ProductionFeedback({
  controller,
  feedback,
  unavailable,
  disabled,
  pages,
}: {
  controller: PresentationProjectController
  feedback?: PresentationProductionFeedbackLedger | null
  unavailable?: true
  disabled: boolean
  pages: { id: string; title: string }[]
}) {
  const latest =
    feedback?.snapshots.at(-1)?.pages ??
    pages.map((page) => ({ pageId: page.id, status: 'not_evaluated' as const }))
  const [draft, setDraft] = useState<PresentationProductionFeedbackPage[]>(() =>
    structuredClone(latest),
  )
  const [downloadError, setDownloadError] = useState(false)
  const assessed = latest.filter((page) => page.status !== 'not_evaluated').length,
    needs = latest.filter((page) => page.status === 'needs_correction').length
  return (
    <section aria-label="人工修正反馈">
      <h4>人工修正反馈</h4>
      <p>请按这次生成版本逐页评价是否需要人工修改。未评估页面保持未知，评价不代表验收通过。</p>
      <button
        type="button"
        disabled={disabled}
        onClick={() => void controller.readProductionFeedback?.()}
      >
        读取当前任务反馈
      </button>
      {unavailable && (
        <p role="status">人工修正反馈不可用，请重新读取；未评估页不能算作无需修正。</p>
      )}
      {feedback !== undefined && (
        <>
          <p>
            已评估 {assessed} / {pages.length} 页 · 需要人工修正 {needs} 页 · 未评估（未知）{' '}
            {pages.length - assessed} 页
          </p>
          <p>
            {assessed
              ? `已评估页需要修正比例：${needs} / ${assessed}`
              : '需要修正比例：未计算（尚未评估）'}
          </p>
          <ol>
            {pages.map((page) => {
              const entry = draft.find((item) => item.pageId === page.id)!
              const noteTooLong = new TextEncoder().encode(entry.note ?? '').byteLength > 2000
              return (
                <li key={page.id}>
                  <p>{page.title}</p>
                  <label>
                    人工修正评价
                    <select
                      aria-label={`人工修正评价 ${page.title}`}
                      disabled={disabled}
                      value={entry.status}
                      onChange={(event) =>
                        setDraft((items) =>
                          items.map((item) =>
                            item.pageId === page.id
                              ? {
                                  ...item,
                                  status: event.target
                                    .value as PresentationProductionFeedbackPage['status'],
                                }
                              : item,
                          ),
                        )
                      }
                    >
                      <option value="not_evaluated">未评估</option>
                      <option value="needs_correction">需要人工修正</option>
                      <option value="no_correction">无需修正</option>
                    </select>
                  </label>
                  <label>
                    可选说明
                    <textarea
                      aria-label={`人工修正说明 ${page.title}`}
                      maxLength={2000}
                      disabled={disabled}
                      value={entry.note ?? ''}
                      onChange={(event) =>
                        setDraft((items) =>
                          items.map((item) =>
                            item.pageId === page.id
                              ? {
                                  ...item,
                                  ...(event.target.value
                                    ? { note: event.target.value }
                                    : { note: undefined }),
                                }
                              : item,
                          ),
                        )
                      }
                    />
                  </label>
                  <p>用户自报评价</p>
                  {noteTooLong && <p role="alert">说明过长，请缩短后保存。</p>}
                </li>
              )
            })}
          </ol>
          <button
            type="button"
            disabled={
              disabled ||
              draft.some((entry) => new TextEncoder().encode(entry.note ?? '').byteLength > 2000)
            }
            onClick={() =>
              void controller.recordProductionFeedback?.(
                draft.map((entry) => ({
                  pageId: entry.pageId,
                  status: entry.status,
                  ...(entry.note ? { note: entry.note } : {}),
                })),
              )
            }
          >
            保存人工修正反馈
          </button>
          {feedback && (
            <>
              <p>
                已保存 {feedback.revision} 次用户评价 · 最近记录{' '}
                {feedback.snapshots.at(-1)?.recordedAt}
              </p>
              <button
                type="button"
                disabled={disabled}
                onClick={() => {
                  try {
                    downloadLocalFile(
                      new TextEncoder().encode(JSON.stringify(feedback, null, 2)),
                      `presentation-feedback-${feedback.requestId}.json`,
                      'application/json',
                    )
                    setDownloadError(false)
                  } catch {
                    setDownloadError(true)
                  }
                }}
              >
                下载人工修正反馈 JSON
              </button>
              {downloadError && <p role="status">本地下载未完成，请重试；已保存反馈保留。</p>}
            </>
          )}
        </>
      )}
    </section>
  )
}

const comparisonConditionLabels: Record<
  PresentationFeedbackComparison['conditions'][number]['key'],
  string
> = {
  brief: '制作要求',
  sources: '来源声明',
  claims: '主张声明',
  research: '研究绑定',
  style: '样式',
  brandKit: '品牌规范',
  parallelism: '有效并行方式',
}
const comparisonGapLabels: Record<PresentationFeedbackComparison['gaps'][number], string> = {
  baseline_not_generic: '基线不是未指定行业的通用计划',
  candidate_not_industry: '当前任务不是五类行业计划',
  input_conditions_differ: '两次制作的声明输入条件不同',
  baseline_feedback_missing: '基线没有保存人工评价',
  candidate_feedback_missing: '当前任务没有保存人工评价',
  baseline_not_fully_evaluated: '基线仍有未评估页面',
  candidate_not_fully_evaluated: '当前任务仍有未评估页面',
}
function FeedbackComparison({
  controller,
  project,
  baselineRequestId,
  comparison,
  unavailable,
  disabled,
}: {
  controller: PresentationProjectController
  project: PresentationProjectStatus
  baselineRequestId?: string
  comparison?: PresentationFeedbackComparison
  unavailable?: true
  disabled: boolean
}) {
  const [downloadError, setDownloadError] = useState(false)
  const tasks =
    project.productionTasks?.filter(
      (task) =>
        task.status === 'compiled' &&
        task.compiledCount === task.total &&
        task.requestId !== project.production?.requestId,
    ) ?? []
  const domains: Record<string, string> = {
    pitch: '路演',
    report: '汇报',
    training: '培训',
    research: '研究报告',
    sales: '销售方案',
    science: '科研',
    law: '法律',
    finance: '金融',
  }
  const signed = (value: number) => (value > 0 ? `+${value}` : String(value))
  return (
    <section aria-label="两次制作反馈对照">
      <h4>两次制作反馈对照</h4>
      <p>两次制作反馈对照，不代表技能效果或验收通过。用户自报评价；未评估页面保持未知。</p>
      <label>
        选择通用制作基线
        <select
          aria-label="选择反馈对照基线"
          disabled={disabled}
          value={baselineRequestId ?? ''}
          onChange={(event) =>
            controller.selectFeedbackComparisonBaseline?.(event.target.value || undefined)
          }
        >
          <option value="">请选择已编译的其它任务</option>
          {tasks.map((task) => (
            <option key={task.requestId} value={task.requestId}>
              制作 {task.sequence} · 计划第 {task.planRevision} 版 · {task.total} 页
            </option>
          ))}
        </select>
      </label>
      {!tasks.length && <p>尚无其它已编译任务可作基线。通用计划与行业计划的角色会在读取后核对。</p>}
      <button
        type="button"
        disabled={disabled || !baselineRequestId}
        onClick={() => void controller.readFeedbackComparison?.()}
      >
        读取两次制作反馈对照
      </button>
      {unavailable && (
        <p role="status">反馈对照不可用，请核对所选任务后重新读取；未显示差值不代表零修正。</p>
      )}
      {comparison && (
        <>
          <p>
            此报告保留比较时读取的反馈版本，后续评价不会改动本报告。计划条件相同不证明原资料、模型或环境相同，也不证明实际使用了行业技能。
          </p>
          {(['baseline', 'candidate'] as const).map((role) => {
            const side = comparison[role],
              counts = side.counts
            return (
              <section
                key={role}
                aria-label={role === 'baseline' ? '对照基线版本' : '对照行业版本'}
              >
                <h5>
                  {role === 'baseline' ? '通用基线' : '当前制作'} · {side.requestId}
                </h5>
                <p>
                  冻结计划第 {side.planRevision} 版 · 领域：
                  {side.plan.domain
                    ? (domains[side.plan.domain] ?? side.plan.domain)
                    : '未指定（通用）'}
                </p>
                <p>
                  反馈版本：{side.feedbackRevision ?? '未保存'} · 记录时间：
                  {side.feedbackRecordedAt ?? '未知'}
                </p>
                <p>
                  已评估 {counts.evaluatedPages} / {counts.totalPages} 页 · 需要人工修正{' '}
                  {counts.needsCorrectionPages} 页 · 无需修正 {counts.noCorrectionPages} 页 ·
                  未评估（未知） {counts.notEvaluatedPages} 页
                </p>
                <p>
                  评价覆盖：{counts.notEvaluatedPages === 0 ? '完整' : '尚不完整'} ·
                  已评估页需要修正比例：
                  {counts.needsCorrectionRate === null
                    ? '未知'
                    : `${counts.needsCorrectionPages} / ${counts.evaluatedPages}（${(counts.needsCorrectionRate * 100).toFixed(1)}%）`}
                </p>
              </section>
            )
          })}
          <h5>对照条件</h5>
          <ul>
            {comparison.conditions.map((condition) => (
              <li key={condition.key}>
                {comparisonConditionLabels[condition.key]}：{condition.match ? '相同' : '不同'}
              </li>
            ))}
          </ul>
          {comparison.gaps.length > 0 && (
            <ul aria-label="反馈对照缺口">
              {comparison.gaps.map((gap) => (
                <li key={gap}>{comparisonGapLabels[gap]}</li>
              ))}
            </ul>
          )}
          <p>
            描述性差值（当前制作减通用基线）：
            {comparison.delta
              ? `需要修正页数 ${signed(comparison.delta.needsCorrectionPages)} 页；已评估页修正比例 ${comparison.delta.needsCorrectionRate * 100 > 0 ? '+' : ''}${(comparison.delta.needsCorrectionRate * 100).toFixed(1)} 个百分点。`
              : '不可计算；缺失评价、未评估页面或比较条件不满足时保持未知。'}
            仅描述这一对观察，不作因果或显著性结论。
          </p>
          <button
            type="button"
            disabled={disabled}
            onClick={() => {
              try {
                downloadLocalFile(
                  new TextEncoder().encode(JSON.stringify(comparison, null, 2)),
                  `presentation-feedback-comparison-${comparison.candidate.requestId}.json`,
                  'application/json',
                )
                setDownloadError(false)
              } catch {
                setDownloadError(true)
              }
            }}
          >
            下载制作反馈对照 JSON
          </button>
          {downloadError && <p role="status">本地下载未完成，请重试；对照报告保留。</p>}
        </>
      )}
    </section>
  )
}

export function PresentationProjectCard(props: {
  controller: PresentationProjectController
  disabled: boolean
  onEndFrontend?: () => void
}) {
  const { controller } = props
  const {
    phase,
    project,
    error,
    feedbackComparisonBaselineRequestId,
    feedbackComparison,
    feedbackComparisonUnavailable,
    productionFeedback,
    productionFeedbackUnavailable,
    deliveryReport,
    deliveryNotice,
    sourceAudit,
    planNotice,
    deliveryBundles,
    bundleNotice,
  } = useSyncExternalStore(
    (listener) => controller.subscribe(listener),
    () => controller.snapshot(),
    () => controller.snapshot(),
  )
  const productionEvents = presentationProductionEventRows(project)
  const [revisionChoice, setRevisionChoice] = useState<{
    projectId: string
    current: number
    revision: number
  }>()
  const history =
    project?.plan?.revisions?.filter((event) => event.revision < project.plan!.revision) ?? []
  const selectedRevision =
    revisionChoice &&
    revisionChoice.projectId === project?.projectId &&
    revisionChoice.current === project?.plan?.revision &&
    history.some((event) => event.revision === revisionChoice.revision)
      ? revisionChoice.revision
      : history.at(-1)?.revision
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
              planning: '正在更新制作计划',
              accepting: '正在保存计划与样式接受决定',
              bundling: '正在处理当前文稿交付包',
              readingResearch: '正在读取计划绑定研究',
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
      {planNotice && <p role="status">{planNotice}</p>}
      {project?.plan?.value.research && (
        <details aria-label="当前计划绑定研究">
          <summary>当前计划绑定研究 #{project.plan.value.research.sequence}</summary>
          <p>
            当前计划明确引用此研究版本，不自动替换为最近研究；整理完成与映射匹配不代表事实、来源权威性、时效或
            QA 通过。
          </p>
          <p>
            映射来源 {project.plan.value.research.sources.length} 条 · 映射主张{' '}
            {project.plan.value.research.claims.length}{' '}
            条；未选引用与冲突另一方仍需在完整研究和冻结报告中核对。
          </p>
          {controller.readBoundResearch && (
            <button
              type="button"
              disabled={disabled}
              onClick={() => void controller.readBoundResearch?.()}
            >
              读取计划绑定研究
            </button>
          )}
          <details>
            <summary>研究身份与映射详情</summary>
            <p>研究记录 ID：{project.plan.value.research.ledgerId}</p>
            <p>研究摘要：{project.plan.value.research.draftDigest}</p>
            <ul>
              {project.plan.value.research.sources.map((item) => (
                <li key={item.sourceId}>
                  计划来源 {item.sourceId} → 原研究来源 {item.researchSourceId}
                </li>
              ))}
            </ul>
            <ul>
              {project.plan.value.research.claims.map((item) => (
                <li key={item.claimId}>
                  计划主张 {item.claimId} → 原研究主张 {item.researchClaimId}
                </li>
              ))}
            </ul>
          </details>
        </details>
      )}
      {project?.deliveryBundlesAvailable && (
        <details aria-label="当前 PowerPoint 文稿交付包">
          <summary>当前 PowerPoint 文稿交付包</summary>
          <p>
            导出当前打开的 PowerPoint 文稿，并附历史 QA
            与保存点；这些记录不代表当前宿主内容已通过验收，检查待完成。
          </p>
          {controller.currentBundleAvailable?.() && (
            <>
              <button
                type="button"
                disabled={disabled || !project.production}
                onClick={() => void controller.exportCurrentBundle?.(false)}
              >
                导出当前文稿交付包
              </button>
              <button
                type="button"
                disabled={disabled || !project.production}
                onClick={() => void controller.exportCurrentBundle?.(true)}
              >
                导出当前文稿交付包（含宿主 PDF）
              </button>
              <button
                type="button"
                disabled={disabled || !project.production}
                onClick={() => void controller.exportCurrentBundle?.(false, true)}
              >
                导出当前文稿交付包（含 8 页宿主截图）
              </button>
              <p>截图仅在当前文稿恰好 8 页且宿主支持截图时采集；逐页人工复核和保存重开仍需完成。</p>
              <p>宿主 PDF 不可用时仍保留 PPTX 包，并标明 PDF 未包含。</p>
            </>
          )}
          <button
            type="button"
            disabled={disabled || !project.production}
            onClick={() => void controller.readDeliveryBundles?.()}
          >
            刷新本机交付包
          </button>
          {bundleNotice && <p role="status">{bundleNotice}</p>}
          {deliveryBundles?.length === 0 && <p>所选页任务暂无本机交付包。</p>}
          <ol>
            {deliveryBundles?.map((bundle) => (
              <li key={bundle.bundleId}>
                {bundle.state === 'ready' ? '本机包已保存' : '上传未完成'} ·{' '}
                <time dateTime={bundle.createdAt}>{bundle.createdAt}</time> · 检查待完成
                {bundle.state === 'ready' && (
                  <button
                    type="button"
                    disabled={disabled}
                    onClick={() => void controller.restoreDeliveryBundle?.(bundle.bundleId)}
                  >
                    恢复 ZIP 到会话附件
                  </button>
                )}
                {controller.deleteDeliveryBundle && (
                  <button
                    type="button"
                    disabled={disabled}
                    onClick={() => {
                      if (
                        window.confirm(
                          '仅删除 PC 缓存的交付包，以释放本机容量。原始 PowerPoint 文稿不受影响；已下载的会话附件保留。确定删除吗？',
                        )
                      )
                        void controller.deleteDeliveryBundle?.(bundle.bundleId)
                    }}
                  >
                    删除本机包
                  </button>
                )}
                <details>
                  <summary>包与历史检查详情</summary>
                  <p>包 ID：{bundle.bundleId}</p>
                  <p>
                    历史 QA：
                    {bundle.manifest.checks.hostQa === 'historical_records_only'
                      ? '包含已有记录'
                      : '未检查'}
                    ；PDF：
                    {bundle.manifest.checks.pdf === 'included'
                      ? '已包含宿主 PDF'
                      : bundle.manifest.checks.pdf === 'unavailable'
                        ? '宿主不可用'
                        : '未请求'}
                    。
                  </p>
                </details>
              </li>
            ))}
          </ol>
        </details>
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
          {project.production.status === 'compiled' && controller.pdfAvailable?.() && (
            <p>PDF 预览来自已编译页任务，不是当前 PowerPoint 宿主文稿。</p>
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
                {page.reusedFromRequestId && (
                  <p>
                    保留既有编译成果 · 来源任务 {page.reusedFromRequestId}；当前内容与视觉仍需核验。
                  </p>
                )}
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
      {project?.production?.status === 'compiled' &&
        controller.readProductionFeedback &&
        controller.recordProductionFeedback && (
          <ProductionFeedback
            key={`${project.projectId}:${project.production.requestId}:${productionFeedback?.revision ?? 'unread'}:${productionFeedback === null}`}
            controller={controller}
            feedback={productionFeedback}
            unavailable={productionFeedbackUnavailable}
            disabled={disabled}
            pages={project.production.pages}
          />
        )}
      {project?.production?.status === 'compiled' &&
        controller.selectFeedbackComparisonBaseline &&
        controller.readFeedbackComparison && (
          <FeedbackComparison
            key={`${project.projectId}:${project.production.requestId}:${feedbackComparisonBaselineRequestId ?? ''}`}
            controller={controller}
            project={project}
            baselineRequestId={feedbackComparisonBaselineRequestId}
            comparison={feedbackComparison}
            unavailable={feedbackComparisonUnavailable}
            disabled={disabled}
          />
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
          {(project.planAcceptance || project.planAcceptanceUnavailable) && (
            <section aria-label="计划与样式接受决定">
              {project.planAcceptanceCurrent && !project.planAcceptanceUnavailable ? (
                <p>
                  用户已接受第 {project.planAcceptanceCurrent.planRevision} 版计划与样式 ·{' '}
                  <time dateTime={project.planAcceptanceCurrent.acceptedAt}>
                    {project.planAcceptanceCurrent.acceptedAt}
                  </time>
                </p>
              ) : (
                <p>当前计划与样式尚无匹配的接受决定。</p>
              )}
              <p>
                可选记录：只表示接受本版故事线与视觉方向；事实来源和页面质量仍需核验。普通制作可继续。
              </p>
              {project.planAcceptanceUnavailable && (
                <p role="alert">接受决定记录暂不可读取，请刷新后核对。</p>
              )}
              {controller.acceptPlan && (
                <button
                  type="button"
                  disabled={
                    disabled ||
                    active ||
                    project.planAcceptanceUnavailable ||
                    Boolean(project.planAcceptanceCurrent)
                  }
                  onClick={() => void controller.acceptPlan!(project.plan!.revision)}
                >
                  接受当前计划与样式
                </button>
              )}
            </section>
          )}
          {project.hostAssociationsUnavailable && (
            <p role="alert">宿主页关联暂不可读取，请刷新后核对导入记录和当前文档。</p>
          )}
          {project.hostAssociations && (
            <details aria-label="计划与宿主页关联">
              <summary>计划与宿主页关联</summary>
              <p>依据持久导入回执核对页面是否存在；不证明页面内容仍匹配计划或已通过验收。</p>
              <ol>
                {project.hostAssociations.pages.map((page) => (
                  <li key={page.pageId}>
                    {project.plan!.value.slides.find((slide) => slide.id === page.pageId)?.title ??
                      page.pageId}
                    {page.hostPages.length === 0 ? (
                      <p>尚无可定位的导入记录。</p>
                    ) : (
                      <ul>
                        {page.hostPages.map((host) => (
                          <li key={`${host.requestId}/${host.slideId}`}>
                            {host.presence === 'present'
                              ? host.position
                                ? `宿主第 ${host.position} 页`
                                : '宿主页存在'
                              : '宿主页已缺失'}{' '}
                            ·{' '}
                            {host.revisionRelation === 'current'
                              ? `当前计划第 ${host.planRevision} 版`
                              : host.revisionRelation === 'historical'
                                ? `历史计划第 ${host.planRevision} 版`
                                : '计划修订未知'}{' '}
                            ·{' '}
                            {host.sourceProof === 'digest'
                              ? '源产物摘要已匹配'
                              : host.sourceProof === 'identity'
                                ? '来源身份对应，旧回执无文件摘要'
                                : '源产物尚未核对'}
                          </li>
                        ))}
                      </ul>
                    )}
                  </li>
                ))}
              </ol>
              {!!project.hostAssociations.unverifiedSources && (
                <p>
                  {project.hostAssociations.unverifiedSources} 份导入源产物尚未核对，请恢复 PC
                  连接或核对原始任务后刷新。
                </p>
              )}
              {project.hostAssociations.legacyImports > 0 && (
                <p>
                  {project.hostAssociations.legacyImports}{' '}
                  份旧导入记录缺少页级身份，未按标题或页序猜测关联。
                </p>
              )}
              {project.hostAssociations.uncertainImports > 0 && (
                <p>
                  {project.hostAssociations.uncertainImports}{' '}
                  份导入记录含未确定的写入，请先核对导入恢复状态。
                </p>
              )}
            </details>
          )}
          {controller.editPlan && history.length > 0 && (
            <details aria-label="历史计划恢复">
              <summary>恢复历史计划</summary>
              <p>
                恢复将创建新修订；已有 PowerPoint
                页面保留。内容相同时无需新建修订，部分旧版本可能无法恢复。
              </p>
              <label>
                选择历史计划版本
                <select
                  aria-label="选择历史计划版本"
                  disabled={disabled}
                  value={selectedRevision ?? ''}
                  onChange={(event) =>
                    setRevisionChoice({
                      projectId: project.projectId,
                      current: project.plan!.revision,
                      revision: Number(event.target.value),
                    })
                  }
                >
                  {history.map((event) => (
                    <option key={event.revision} value={event.revision}>
                      第 {event.revision} 版
                      {event.snapshot ? ` · ${event.snapshot.slideCount} 页` : ''} · 登记于{' '}
                      {new Date(event.createdAt).toLocaleString()}
                    </option>
                  ))}
                </select>
              </label>
              <button
                type="button"
                disabled={disabled || !selectedRevision}
                onClick={() => {
                  if (selectedRevision)
                    void controller.editPlan?.(project.plan!.revision, {
                      kind: 'restore',
                      revision: selectedRevision,
                    })
                }}
              >
                恢复为新计划修订
              </button>
            </details>
          )}
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
                    {sourceAudit.finishedAt && (
                      <>
                        历史核对结果 ·{' '}
                        <time dateTime={sourceAudit.finishedAt}>{sourceAudit.finishedAt}</time>{' '}
                        ·{' '}
                      </>
                    )}
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
                            source_mismatch: '网页快照与计划网址不匹配',
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
          {controller.editPlan && (
            <p>
              只调整制作计划；已有 PowerPoint
              页面保留。旧后台任务按原快照继续，新制作使用更新后的计划。
            </p>
          )}
          <ol aria-label="逐页施工图">
            {project.plan.value.slides.map((slide, index) => (
              <li key={slide.id}>
                <strong>{slide.title}</strong>
                {controller.editPlan && (
                  <details>
                    <summary>调整本页计划</summary>
                    <p>
                      {slide.locked ? '计划页已锁定' : '计划页未锁定'}
                      ；锁定保护计划及后续生产输入，宿主页请另行核对。
                    </p>
                    <div className="presentation-project-actions">
                      <button
                        type="button"
                        disabled={disabled}
                        aria-label={`${slide.locked ? '解除计划页锁定' : '锁定计划页'} ${index + 1}：${slide.title}`}
                        onClick={() =>
                          void controller.editPlan?.(project.plan!.revision, {
                            kind: 'lock',
                            pageId: slide.id,
                            locked: !slide.locked,
                          })
                        }
                      >
                        {slide.locked ? '解除计划页锁定' : '锁定计划页'}
                      </button>
                      <button
                        type="button"
                        aria-label={`上移计划页 ${index + 1}：${slide.title}`}
                        disabled={disabled || slide.locked || index === 0}
                        onClick={() =>
                          void controller.editPlan?.(project.plan!.revision, {
                            kind: 'move',
                            pageId: slide.id,
                            direction: 'up',
                          })
                        }
                      >
                        上移
                      </button>
                      <button
                        type="button"
                        aria-label={`下移计划页 ${index + 1}：${slide.title}`}
                        disabled={
                          disabled ||
                          slide.locked ||
                          index === project.plan!.value.slides.length - 1
                        }
                        onClick={() =>
                          void controller.editPlan?.(project.plan!.revision, {
                            kind: 'move',
                            pageId: slide.id,
                            direction: 'down',
                          })
                        }
                      >
                        下移
                      </button>
                      <button
                        type="button"
                        aria-label={`删除计划页 ${index + 1}：${slide.title}`}
                        disabled={
                          disabled || slide.locked || project.plan!.value.slides.length === 1
                        }
                        onClick={() =>
                          void controller.editPlan?.(project.plan!.revision, {
                            kind: 'delete',
                            pageId: slide.id,
                          })
                        }
                      >
                        删除计划页
                      </button>
                    </div>
                  </details>
                )}
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
