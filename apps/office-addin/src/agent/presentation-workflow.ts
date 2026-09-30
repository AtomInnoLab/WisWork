import {
  parsePresentationDeliveryBundleReceipt,
  type PresentationDeliveryBundleReceipt,
} from '@wiswork/project-store/presentation-delivery-bundle'
import type { PresentationProjectStatus } from '../skills/powerpoint/presentation-project.js'
import type { PresentationImportProgress } from '../skills/powerpoint/presentation-page-delivery.js'
import {
  validatePresentationQaRecord,
  type PresentationQaRecord,
} from '../skills/powerpoint/presentation-qa.js'
import {
  validatePresentationQaAttempt,
  type PresentationQaAttempt,
} from '../skills/powerpoint/presentation-qa-attempts.js'
import type { PresentationDeliveryReport } from '@wiswork/pptx-engine/presentation-delivery-report'
import { PRESENTATION_DOMAIN_PROFILES } from '@wiswork/pptx-engine/presentation-plan'
import type { PresentationSourceAuditHistory } from '@wiswork/project-store/presentation-source-audit'

const jobEventLabels = {
  'run.started': '开始逐页制作',
  'run.pause_requested': '请求暂停制作',
  'run.paused': '已暂停制作',
  'run.cancel_requested': '请求取消制作',
  'run.cancelled': '已取消制作',
  'run.interrupted': '制作中断',
  'run.completed': '逐页编译完成',
  'run.failed': '页任务失败',
  'page.started': '开始编译',
  'page.compiled': '编译完成',
  'page.failed': '编译失败',
} as const

export interface PresentationProductionEventRow {
  id: string
  text: string
  at: string
  pageId?: string
  attempts?: { id: string; text: string; at: string }[]
}

/** Fold only retained display events; persisted events and page progress stay unchanged. */
export function presentationProductionEventRows(
  project: PresentationProjectStatus | undefined,
):
  | { rows: PresentationProductionEventRow[]; retainedEventCount: number; truncated: boolean }
  | undefined {
  const job = project?.productionJob
  const production = project?.production
  if (
    !project ||
    !job ||
    !production ||
    job.projectId !== project.projectId ||
    production.projectId !== project.projectId ||
    job.requestId !== production.requestId
  )
    return undefined
  const events = job.events.slice(-20)
  const errorText = (event: (typeof events)[number]) =>
    'error' in event && event.error
      ? ` · ${
          {
            compile_failed: '编译失败',
            invalid_deck: '页面内容无效',
            aborted: '已停止',
            output_too_large: '成果过大',
            asset_unavailable: '素材不可用',
            source_unavailable: '附件不可读或来源摘录未匹配',
            font_unavailable: '指定字体及回退字体均不可用',
            invalid_state: '任务状态异常',
          }[event.error]
        }`
      : ''
  type PageEvent = Extract<(typeof events)[number], { pageId: string }>
  const pages = new Map<string, PageEvent[]>()
  for (const event of events) {
    if (!('pageId' in event)) continue
    const attempts = pages.get(event.pageId) ?? []
    attempts.push(event)
    pages.set(event.pageId, attempts)
  }
  const rows: PresentationProductionEventRow[] = []
  for (const event of events) {
    if (!('pageId' in event)) {
      rows.push({
        id: `job-run:${JSON.stringify([project.projectId, job.requestId, event.sequence])}`,
        text: `${jobEventLabels[event.type]}${errorText(event)}`,
        at: event.createdAt,
      })
      continue
    }
    const attempts = pages.get(event.pageId)!
    if (attempts.at(-1) !== event) continue
    const page = production.pages.find((page) => page.id === event.pageId)
    const planned =
      project.plan?.revision === production.planRevision
        ? project.plan.value.slides.find((page) => page.id === event.pageId)
        : undefined
    const id = `job-page:${JSON.stringify([project.projectId, job.requestId, event.pageId])}`
    const eventText = (attempt: PageEvent) =>
      `第 ${attempt.attempt} 次 · ${jobEventLabels[attempt.type]}${attempt.type === 'page.compiled' ? '（未导入验收）' : ''}${errorText(attempt)}`
    rows.push({
      id,
      pageId: event.pageId,
      text: `页面 ${page?.title ?? planned?.title ?? event.pageId} · ${eventText(event)}`,
      at: event.createdAt,
      attempts: attempts.map((attempt) => ({
        id: `${id}:event-${attempt.sequence}`,
        text: eventText(attempt),
        at: attempt.createdAt,
      })),
    })
  }
  return {
    rows,
    retainedEventCount: events.length,
    truncated:
      job.events.length > events.length ||
      job.revision > job.events.length ||
      (events[0]?.sequence ?? 1) > 1,
  }
}

export interface PresentationWorkflowSummary {
  stages: {
    name: string
    detail: string
    status: 'pending' | 'working' | 'attention' | 'recorded'
  }[]
  timeline: {
    id: string
    text: string
    at?: string
    type?:
      | 'asset.fetching'
      | 'asset.ready'
      | 'asset.rejected'
      | 'research.started'
      | 'research.completed'
      | 'research.failed'
      | 'source_audit.started'
      | 'source_audit.completed'
      | 'source_audit.failed'
      | 'plan.proposed'
      | 'plan.revised'
      | 'style.proposed'
      | 'plan.approved'
      | 'style.approved'
      | 'delivery.bundle.started'
      | 'delivery.bundle.ready'
      | 'host.import.started'
      | 'host.import.recorded'
      | 'host.import.uncertain'
      | 'qa.capture.recorded'
      | 'qa.visual.recorded'
      | 'qa.evidence.invalidated'
      | 'qa.attempt.started'
      | 'qa.attempt.recorded'
      | 'qa.attempt.waiting'
      | 'qa.attempt.failed'
      | 'qa.attempt.cancelled'
      | 'qa.attempt.closed'
    scope?:
      | 'production_asset_resolution'
      | 'source_excerpt_audit'
      | 'research_ledger'
      | 'saved_plan'
      | 'saved_style'
      | 'explicit_user_decision'
      | 'current_document_delivery_bundle'
      | 'host_page_import'
      | 'saved_page_qa'
      | 'page_qa_attempt'
    recordsLabel?: string
    records?: {
      id: string
      text: string
      at: string
      type?:
        | 'host.import.started'
        | 'host.import.recorded'
        | 'qa.capture.recorded'
        | 'qa.visual.recorded'
        | 'qa.evidence.invalidated'
        | 'qa.attempt.started'
        | 'qa.attempt.recorded'
        | 'qa.attempt.waiting'
        | 'qa.attempt.failed'
        | 'qa.attempt.cancelled'
        | 'qa.attempt.closed'
    }[]
  }[]
  attention: { id: string; text: string }[]
  pages: {
    id: string
    title: string
    section?: string
    production: string
    imported: string
    qa: string
    evidence: string
    nextAction: string
  }[]
  nextAction: string
  nextTool?: 'start_job' | 'resume_job' | 'run_pages' | 'prepare_import' | 'read_report'
}

/** A view over durable records, never an additional source of completion truth. */
export function presentationWorkflowSummary(
  project: PresentationProjectStatus | undefined,
  imported: PresentationImportProgress | undefined,
  qa: PresentationQaRecord | undefined,
  report?: PresentationDeliveryReport,
  delivery?: { bundles?: PresentationDeliveryBundleReceipt[]; unavailable?: boolean },
  qaAttempts?: { attempts?: PresentationQaAttempt[]; unavailable?: boolean },
): PresentationWorkflowSummary | undefined {
  if (!project) return undefined
  const plan = project.plan?.value
  const preparedSources = project.sourcePreparation
  const sourceProblems =
    preparedSources?.filter((source) => source.status !== 'excerpt_matched') ?? []
  const sourceStatus = preparedSources
    ? `；当前文档引用附件 ${preparedSources.filter((source) => source.status === 'excerpt_matched').length}/${preparedSources.length} 份摘录已匹配原文${sourceProblems.length ? `，${sourceProblems.length} 份需处理` : ''}`
    : project.sourcePreparationUnavailable
      ? '；当前文档引用附件状态暂不可读取'
      : ''
  const domainProfile = plan?.domain ? PRESENTATION_DOMAIN_PROFILES[plan.domain] : undefined
  const sectionLabels = new Map<string, string>(
    domainProfile?.sections.map(
      (section, index) => [section, domainProfile.labels[index]!] as [string, string],
    ) ?? [],
  )
  const production = project.production
  const pageIds = production?.pages.map((page) => page.id)
  const importMatches = Boolean(
    production &&
    imported?.source === 'production' &&
    imported.projectId === project.projectId &&
    imported.requestId === production.requestId &&
    imported.total === production.total &&
    JSON.stringify(imported.pages.map((page) => page.id)) === JSON.stringify(pageIds),
  )
  const qaMatches = Boolean(
    production &&
    qa?.source === 'production' &&
    qa.projectId === project.projectId &&
    qa.requestId === production.requestId &&
    qa.pages.length === production.total &&
    qa.pages.every((page) => pageIds?.includes(page.pageId)),
  )
  const reviewed = qaMatches
    ? qa!.pages.filter(
        (page) =>
          !page.recheckRequired &&
          page.screenshotRenderer !== 'libreoffice' &&
          page.structure.status === 'passed' &&
          page.visual.status === 'pass',
      ).length
    : 0
  const fallbackReviewed = qaMatches
    ? qa!.pages.filter(
        (page) =>
          !page.recheckRequired &&
          page.screenshotRenderer === 'libreoffice' &&
          page.structure.status === 'passed' &&
          page.visual.status === 'pass',
      ).length
    : 0
  const reportMatches = Boolean(
    production &&
    report &&
    report.projectId === project.projectId &&
    report.requestId === production.requestId &&
    report.planRevision === production.planRevision,
  )
  const openIssues = reportMatches
    ? report!.pages.reduce(
        (sum, page) =>
          sum +
          page.issues.filter(
            (issue) => issue.disposition.state === 'open' || issue.disposition.stale,
          ).length,
        0,
      )
    : 0
  const failed = production?.pages.filter((page) => page.state === 'failed').length ?? 0
  const planChangedSinceProduction = Boolean(
    plan && production && production.planRevision !== project.plan!.revision,
  )
  const revisions = project.plan?.revisions
  const productionStyle = revisions?.find((entry) => entry.revision === production?.planRevision)
    ?.snapshot?.styleDigest
  const latestStyle = revisions?.find((entry) => entry.revision === project.plan?.revision)
    ?.snapshot?.styleDigest
  const styleChangedSinceProduction = Boolean(
    planChangedSinceProduction && productionStyle && latestStyle && productionStyle !== latestStyle,
  )
  const uncertain = importMatches
    ? imported!.pages.filter((page) => page.state === 'uncertain').length
    : 0
  const recheck = qaMatches ? qa!.pages.filter((page) => page.recheckRequired).length : 0
  const attention: PresentationWorkflowSummary['attention'] = []
  if (planChangedSinceProduction)
    attention.push({
      id: 'plan-revision',
      text: `当前选中页任务依据计划第 ${production!.planRevision} 版，现已保存第 ${project.plan!.revision} 版；请确认继续旧任务或选择新任务。`,
    })
  if (styleChangedSinceProduction)
    attention.push({
      id: 'style-revision',
      text: '品牌或样式规则已变化；当前页面的历史视觉审查只适用于旧计划。选择新任务后需重新审查受影响页面。',
    })
  if (failed)
    attention.push({
      id: 'failed-pages',
      text: `${failed} 页编译失败；已成功页面保留，请修复失败页后继续生产。`,
    })
  if (uncertain)
    attention.push({
      id: 'uncertain-import',
      text: `${uncertain} 页写入结果不确定；请先检查 PowerPoint 文档，再继续导入。`,
    })
  if (recheck)
    attention.push({
      id: 'qa-recheck',
      text: `${recheck} 页历史审查已失效；请重新采集并审查受影响页面。`,
    })
  if (openIssues)
    attention.push({
      id: 'content-issues',
      text: `${openIssues} 项内容证据问题待处理；请查看交付报告。`,
    })
  if (sourceProblems.length) {
    const labels = sourceProblems
      .slice(0, 3)
      .map(
        (source) =>
          `${plan?.sources.find((item) => item.id === source.sourceId)?.title ?? source.sourceId}（${
            {
              uploading: '上传中',
              failed: '解析失败',
              missing: '缺失',
              unsupported: '非文本资料',
              excerpt_mismatch: '摘录不在原文中',
              excerpt_missing: '计划摘录为空',
              source_mismatch: '网页快照与计划网址不匹配',
              locator_mismatch: '计划定位与原文实际位置不匹配',
              ready: '旧版 PC 未核对摘录',
              excerpt_matched: '摘录已匹配',
            }[source.status]
          }）`,
      )
      .join('、')
    attention.push({
      id: 'source-preparation',
      text: `${sourceProblems.length} 份计划引用资料尚不可用于原文追溯：${labels}${sourceProblems.length > 3 ? '等' : ''}；请检查当前文档附件。`,
    })
  } else if (project.sourcePreparationUnavailable)
    attention.push({
      id: 'source-preparation-unavailable',
      text: '当前文档引用资料状态暂不可读取；请重试项目状态后再判断原文追溯是否可用。',
    })
  const pages = (production?.pages ?? project.slides).map((slide) => {
    const plannedSlide = plan?.slides.find((item) => item.id === slide.id)
    const page = production?.pages.find((item) => item.id === slide.id)
    const index = production?.pages.findIndex((item) => item.id === slide.id) ?? -1
    const importedPage = importMatches && index >= 0 ? imported!.pages[index] : undefined
    const reviewedPage = qaMatches ? qa!.pages.find((item) => item.pageId === slide.id) : undefined
    const reportPage = reportMatches
      ? report!.pages.find((item) => item.pageId === slide.id)
      : undefined
    const pageIssues =
      reportPage?.issues.filter(
        (issue) => issue.disposition.state === 'open' || issue.disposition.stale,
      ).length ?? 0
    const evidenceText = !reportMatches
      ? '无当前任务内容报告'
      : !reportPage
        ? '当前报告缺少此页'
        : pageIssues
          ? `${pageIssues} 项证据问题待处理`
          : '无未决证据问题；来源仍需核验'
    const productionText = page
      ? {
          pending: '待编译',
          building: '编译中',
          failed: '编译失败',
          compiled: '已编译',
        }[page.state]
      : '待创建页任务'
    const importText = importedPage
      ? {
          pending: '待导入',
          uncertain: '写入待核查',
          complete: '已记录导入',
        }[importedPage.state]
      : '无当前任务导入记录'
    const qaText = reviewedPage
      ? reviewedPage.recheckRequired
        ? '历史检查已失效'
        : reviewedPage.visual.status === 'pass' &&
            reviewedPage.structure.status === 'passed' &&
            reviewedPage.screenshotRenderer === 'libreoffice'
          ? '备用预览已复核；宿主外观待核验'
          : reviewedPage.visual.status === 'pass' && reviewedPage.structure.status === 'passed'
            ? styleChangedSinceProduction
              ? '旧样式版本历史通过'
              : '历史结构与视觉通过'
            : reviewedPage.visual.status === 'needs_changes'
              ? '历史检查需修改'
              : '待完成页面复核'
      : '无当前任务 QA 记录'
    const pageNext = planChangedSinceProduction
      ? '先确认继续旧计划或选择新任务'
      : !page || page.state === 'pending'
        ? '制作页面'
        : page.state === 'building'
          ? '等待当前编译'
          : page.state === 'failed'
            ? '修复并重试此页'
            : !importedPage || importedPage.state === 'pending'
              ? '导入此页'
              : importedPage.state === 'uncertain'
                ? '检查宿主页后再继续'
                : !reviewedPage
                  ? '采集截图并复核'
                  : reviewedPage.recheckRequired
                    ? '重审此页'
                    : reviewedPage.visual.status === 'needs_changes'
                      ? '修改页面后重审'
                      : qaText === '历史结构与视觉通过'
                        ? pageIssues
                          ? '处理此页内容证据问题'
                          : '继续来源与保存重开核验'
                        : reviewedPage.screenshotRenderer === 'libreoffice'
                          ? 'PowerPoint 可用后核验宿主外观'
                          : '完成页面复核'
    return {
      id: slide.id,
      title: slide.title,
      ...(!planChangedSinceProduction && plannedSlide?.domainSection
        ? { section: sectionLabels.get(plannedSlide.domainSection) }
        : {}),
      production: productionText,
      imported: importText,
      qa: qaText,
      evidence: evidenceText,
      nextAction: pageNext,
    }
  })
  const stages: PresentationWorkflowSummary['stages'] = [
    {
      name: '目标与资料',
      status:
        sourceProblems.length || project.sourcePreparationUnavailable
          ? 'attention'
          : plan
            ? 'recorded'
            : 'pending',
      detail: plan
        ? `Brief 已保存；登记 ${plan.sources.length} 份资料、${plan.claims.length} 条主张${sourceStatus}，真实性仍需核验`
        : '尚无已保存的结构化计划',
    },
    {
      name: '故事线与样式',
      status: planChangedSinceProduction ? 'attention' : plan ? 'recorded' : 'pending',
      detail: plan
        ? `已保存 ${plan.slides.length} 页施工图和样式契约${domainProfile ? ` · ${domainProfile.title}结构` : ''} · 计划第 ${project.plan!.revision} 版`
        : '待保存逐页施工图与样式契约',
    },
    {
      name: '逐页制作',
      status: failed
        ? 'attention'
        : production
          ? production.compiledCount === production.total
            ? 'recorded'
            : 'working'
          : 'pending',
      detail: production
        ? `已编译 ${production.compiledCount}/${production.total} 页${failed ? ` · ${failed} 页失败待重试` : ''}；尚不代表导入或验收`
        : project.status === 'compiled'
          ? '旧编译成果可用；尚无逐页生产记录'
          : '尚未启动逐页生产',
    },
    {
      name: '导入 PowerPoint',
      status:
        (imported && !importMatches) || imported?.status === 'uncertain'
          ? 'attention'
          : importMatches
            ? imported!.completed === imported!.total
              ? 'recorded'
              : 'working'
            : 'pending',
      detail: importMatches
        ? `已记录 ${imported!.completed}/${imported!.total} 页；${imported!.status === 'uncertain' ? '写入结果不确定，需检查文档' : '导入不代表视觉验收'}`
        : imported
          ? '现有导入记录无法与当前页任务匹配，需核对'
          : '尚无当前页任务的导入记录',
    },
    {
      name: '页面审查',
      status:
        (qa && !qaMatches) ||
        recheck ||
        (qaMatches &&
          qa!.pages.some(
            (page) => page.visual.status === 'needs_changes' || page.structure.status !== 'passed',
          ))
          ? 'attention'
          : qaMatches
            ? reviewed === qa!.pages.length
              ? 'recorded'
              : 'working'
            : 'pending',
      detail: qaMatches
        ? `PowerPoint 宿主结构与视觉复核 ${reviewed}/${qa!.pages.length} 页通过；${fallbackReviewed} 页使用备用预览且宿主外观待核验；${qa!.pages.filter((page) => page.recheckRequired).length} 页需重审`
        : qa
          ? '现有 QA 记录无法与当前页任务匹配，需核对'
          : '尚无当前页任务的 QA 记录',
    },
    {
      name: '交付核验',
      status: openIssues ? 'attention' : reportMatches ? 'working' : 'pending',
      detail: `${reportMatches ? `内容证据报告有 ${openIssues} 项待处理；` : '尚无当前任务的内容证据报告；'}来源真实性、保存重开及真实 PowerPoint 验收尚不能由上述记录证明`,
    },
  ]
  // Rebuild from durable records. Undated entries are current checkpoints, not events.
  const timeline: PresentationWorkflowSummary['timeline'] = []
  if (project.planAcceptanceUnavailable)
    attention.push({
      id: 'plan-acceptance-unavailable',
      text: '计划与样式接受记录暂不可读取；普通制作可继续，不沿用未知接受结论。',
    })
  else if (project.planAcceptance?.projectId === project.projectId) {
    for (const record of project.planAcceptance.records) {
      const current = project.planAcceptanceCurrent?.decisionId === record.decisionId
      const key = JSON.stringify([project.projectId, record.decisionId])
      timeline.push({
        id: `plan-acceptance:${key}`,
        type: 'plan.approved',
        scope: 'explicit_user_decision',
        at: record.acceptedAt,
        text: `用户接受计划第 ${record.planRevision} 版${current ? '（当前版本）' : '（历史决定，不沿用到其它版本）'}；事实来源仍需核验`,
      })
      timeline.push({
        id: `style-acceptance:${key}`,
        type: 'style.approved',
        scope: 'explicit_user_decision',
        at: record.acceptedAt,
        text: `用户接受计划第 ${record.planRevision} 版的样式${current ? '（当前版本）' : '（历史决定）'}；页面视觉效果仍需审查`,
      })
    }
  }
  if (project.assetHistoryUnavailable)
    attention.push({
      id: 'asset-history-unavailable',
      text: '素材解析记录不可读，页面制作状态仍可查看。',
    })
  const assets = project.assetHistory
  if (
    assets &&
    !project.assetHistoryUnavailable &&
    assets.projectId === project.projectId &&
    assets.requestId === project.production?.requestId
  ) {
    const groups = new Map<string, typeof assets.events>()
    for (const event of assets.events) {
      const key = JSON.stringify([event.pageId, event.assetId])
      groups.set(key, [...(groups.get(key) ?? []), event])
    }
    const labels = {
      'asset.fetching': '开始解析（仅有开始记录不能证明仍在执行）',
      'asset.ready': '素材字节就绪（许可、版权与页面质量未核验）',
      'asset.rejected': '素材解析失败',
    }
    for (const [key, events] of groups) {
      const latest = events.at(-1)!
      const title =
        project.production?.pages.find((page) => page.id === latest.pageId)?.title ?? latest.pageId
      timeline.push({
        id: `asset:${assets.requestId}:${key}`,
        type: latest.type,
        scope: 'production_asset_resolution',
        at: latest.createdAt,
        text: `页面 ${title} · 素材 ${latest.assetId} · 第 ${latest.attempt} 次 · ${labels[latest.type]}`,
        recordsLabel: '素材解析记录',
        records: events.map((event) => ({
          id: `asset:${assets.requestId}:${key}:${event.sequence}`,
          at: event.createdAt,
          text: `第 ${event.attempt} 次 · ${labels[event.type]}${event.error ? ' · ' + { asset_unavailable: '素材不可用', output_too_large: '素材超出容量', aborted: '已停止' }[event.error] : ''}`,
        })),
      })
    }
    if (assets.revision > assets.events.length)
      attention.push({
        id: 'asset-history-truncated',
        text: '素材解析仅保留最近 128 条事件，较早记录未展示。',
      })
  }
  const research = project.researchSummary
  if (project.researchHistoryUnavailable)
    attention.push({
      id: 'research-history-unavailable',
      text: '独立研究整理历史暂不可读取；不能沿用未知记录状态或结论，请刷新项目状态后核对。',
    })
  else if (research?.projectId === project.projectId) {
    const binding = plan?.research
    const disclaimer = '仅为研究整理归档，不代表主张支持、来源权威性或时效通过'
    for (const record of research.records) {
      const id = `research-ledger:${JSON.stringify([research.documentId, research.projectId, record.id, record.sequence, record.draftDigest])}`
      const bound =
        binding?.ledgerId === record.id &&
        binding.sequence === record.sequence &&
        binding.draftDigest === record.draftDigest
      const result =
        record.state === 'completed'
          ? '整理已归档（待核验）'
          : record.state === 'failed'
            ? `整理未完成：${record.error === 'aborted' ? '整理已中断' : record.error === 'source_unavailable' ? '原文资料暂不可用' : '研究记录状态异常'}`
            : '缺少结束回执，不能证明仍在执行或已经中断'
      const records = [
        {
          id: `${id}:started`,
          at: record.startedAt,
          text: `开始整理研究 #${record.sequence} · 原记录 ${record.id} · 草稿摘要 ${record.draftDigest}；${disclaimer}`,
        },
      ]
      if (record.finishedAt)
        records.push({
          id: `${id}:finished`,
          at: record.finishedAt,
          text: `${result}；${disclaimer}`,
        })
      timeline.push({
        id,
        scope: 'research_ledger',
        type:
          record.state === 'completed'
            ? 'research.completed'
            : record.state === 'failed'
              ? 'research.failed'
              : 'research.started',
        at: record.finishedAt ?? record.startedAt,
        text: `研究 #${record.sequence} · ${result} · 来源 ${record.sourceCount} · 结论 ${record.factCount} · 冲突 ${record.conflictCount}；${bound ? '当前计划精确绑定' : '历史研究，不替代当前计划绑定'}；${disclaimer}`,
        recordsLabel: '研究整理记录',
        records,
      })
    }
    if (research.records.some((record) => record.state === 'running'))
      attention.push({
        id: 'research-unfinished',
        text: '研究整理有开始记录但缺少结束回执，不能证明后台仍在执行或已经中断；请只读核对原记录，不自动重跑。',
      })
    if (research.records.some((record) => record.state === 'failed'))
      attention.push({
        id: 'research-failed',
        text: '研究整理存在未完成记录；原草稿与记录保留，请读取原记录核对，不自动重跑。',
      })
    if (research.records.some((record) => record.conflictCount > 0))
      attention.push({
        id: 'research-conflicts',
        text: '研究整理归档记录包含冲突；请读取双方原结论与来源，历史记录不替代当前计划绑定或事实核验。',
      })
    if (research.totalRecords > research.records.length)
      attention.push({
        id: 'research-history-window',
        text: `研究整理现存 ${research.totalRecords} 条；仅展示最近 ${research.records.length} 条，较早记录不在此窗口内；序号间隔不代表缺失记录的状态。`,
      })
    if (
      binding &&
      !research.records.some(
        (record) =>
          record.id === binding.ledgerId &&
          record.sequence === binding.sequence &&
          record.draftDigest === binding.draftDigest,
      )
    )
      attention.push({
        id: 'research-binding-window',
        text: '当前计划精确绑定的研究记录未出现在最近窗口中；不能据此推断删除或完成，请按原记录 ID 读取核对，不替换绑定。',
      })
  }
  const sourceHistory = project.sourceAuditHistory
  if (project.sourceAuditHistoryUnavailable) {
    attention.push({
      id: 'source-audit-history-unavailable',
      text: '资料核对历史暂不可读取；项目与计划仍保留，请恢复记录后重试，不沿用未知核对结论。',
    })
  } else if (sourceHistory?.projectId === project.projectId) {
    const groups = new Map<string, PresentationSourceAuditHistory['runs']>()
    for (const run of sourceHistory.runs) {
      const key = JSON.stringify([project.projectId, run.planRevision, run.planDigest])
      const group = groups.get(key) ?? []
      group.push(run)
      groups.set(key, group)
    }
    for (const [key, group] of groups) {
      const latest = group.at(-1)!
      const text = (run: typeof latest) =>
        run.state === 'completed'
          ? `资料摘录核对结束：${run.foundCount}/${run.sourceCount} 份字面匹配`
          : run.state === 'failed'
            ? `资料摘录核对未完成：${run.error === 'aborted' ? '已停止' : run.error === 'invalid_state' ? '资料或记录状态异常' : '资料暂不可读取'}`
            : '核对已开始但无结果回执，不能证明仍在后台执行；可重新核对'
      timeline.push({
        id: `source-audit:${key}`,
        type:
          latest.state === 'completed'
            ? 'source_audit.completed'
            : latest.state === 'failed'
              ? 'source_audit.failed'
              : 'source_audit.started',
        scope: 'source_excerpt_audit',
        text: `计划第 ${latest.planRevision} 版 · ${text(latest)}；窗口内保留 ${group.length} 次核对，来源真实性、适用范围和时效未核验`,
        at: latest.finishedAt ?? latest.startedAt,
        records: group.map((run) => ({
          id: run.id,
          text: `开始 ${run.startedAt} · ${text(run)}`,
          at: run.finishedAt ?? run.startedAt,
        })),
      })
    }
    if (sourceHistory.runs.some((run) => run.state === 'running'))
      attention.push({
        id: 'source-audit-unfinished',
        text: '历史资料核对有开始记录但无结果回执，不代表后台仍在运行；当前计划可重新核对，旧记录保持原身份。',
      })
  }
  if (preparedSources)
    timeline.push({
      id: 'source-preparation',
      text: `当前文档资料检查点：${preparedSources.filter((source) => source.status === 'excerpt_matched').length}/${preparedSources.length} 份计划引用附件的摘录逐字匹配原文；来源真实性未核验`,
    })
  if (plan) {
    const revisions = project.plan!.revisions
    if (revisions?.length) {
      if (
        revisions[0]!.revision !== 1 ||
        revisions.some(
          (event, index) => index > 0 && event.revision !== revisions[index - 1]!.revision + 1,
        )
      )
        timeline.push({
          id: 'plan-history',
          text: '仅展示已保留的计划版本；不推断缺失版本的资料、主张或样式变化。',
        })
      for (const [index, event] of revisions.entries()) {
        const current = event.snapshot
        const previous =
          revisions[index - 1]?.revision === event.revision - 1
            ? revisions[index - 1]?.snapshot
            : undefined
        const changed =
          current && previous
            ? [
                current.sourcesDigest !== previous.sourcesDigest ? '已登记资料' : undefined,
                current.claimsDigest !== previous.claimsDigest ? '主张' : undefined,
                current.slidesDigest !== previous.slidesDigest ? '逐页计划' : undefined,
                current.styleDigest !== previous.styleDigest ? '样式规范' : undefined,
              ].filter(Boolean)
            : []
        const detail =
          current && previous
            ? `；${changed.length ? changed.join('、') : 'Brief 或其他计划字段'}有变化`
            : current
              ? `；登记 ${current.sourceCount} 份资料、${current.claimCount} 条主张、${current.slideCount} 页计划`
              : ''
        timeline.push({
          id: `plan-${event.revision}`,
          type: event.revision === 1 ? 'plan.proposed' : 'plan.revised',
          scope: 'saved_plan',
          text: `已保存计划第 ${event.revision} 版${detail}；来源真实性仍需核验`,
          at: event.createdAt,
        })
        if (
          current &&
          (!previous ||
            current.sourcesDigest !== previous.sourcesDigest ||
            current.claimsDigest !== previous.claimsDigest)
        )
          timeline.push({
            id: `research-${event.revision}`,
            text: `第 ${event.revision} 版已登记 ${current.sourceCount} 份资料、${current.claimCount} 条主张；尚需核对来源与结论`,
            at: event.createdAt,
          })
        if (current && (!previous || current.styleDigest !== previous.styleDigest))
          timeline.push({
            id: `style-${event.revision}`,
            type: 'style.proposed',
            scope: 'saved_style',
            text: `第 ${event.revision} 版样式规范已保存；页面视觉效果仍需审查`,
            at: event.createdAt,
          })
      }
    } else
      timeline.push({
        id: 'plan',
        text: `已保存计划第 ${project.plan!.revision} 版：${plan.slides.length} 页，${plan.sources.length} 份资料`,
      })
  }
  if (production) {
    timeline.push({
      id: 'production',
      text: `当前页任务 ${production.requestId}：已编译 ${production.compiledCount}/${production.total} 页`,
    })
    const grouped = presentationProductionEventRows(project)
    if (grouped) {
      timeline.push(...grouped.rows)
      if (grouped.truncated)
        timeline.push({
          id: `job-history:${JSON.stringify([project.projectId, production.requestId])}`,
          text: '仅展示最近保留的生产事件；更早历史已截断，此处不是完整审计记录或失败次数统计。',
        })
    }
  }
  for (const task of project.productionTasks?.slice(0, 20) ?? []) {
    if (!task.lastEvent || task.requestId === project.productionJob?.requestId) continue
    timeline.push({
      id: `task-${task.requestId}-latest`,
      text: `页任务 ${task.requestId}：${jobEventLabels[task.lastEvent.type]}${task.lastEvent.pageId ? ` · ${task.lastEvent.pageId}` : ''}；已编译 ${task.compiledCount}/${task.total} 页`,
      at: task.lastEvent.createdAt,
    })
  }
  if (importMatches) {
    timeline.push({
      id: 'import',
      text: `导入检查点：${imported!.completed}/${imported!.total} 页${imported!.status === 'uncertain' ? '，有写入待核查' : ''}`,
    })
    const timestamp = (value: unknown): value is string =>
      typeof value === 'string' &&
      /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(value) &&
      Number.isFinite(Date.parse(value)) &&
      new Date(value).toISOString() === value
    const hostId = (value: unknown): value is string =>
      typeof value === 'string' && value.length > 0 && value.length <= 256
    for (const page of imported!.pages) {
      if (production!.projectId !== project.projectId) continue
      const recorded = page.state === 'complete'
      if (
        (page.state !== 'complete' && page.state !== 'uncertain') ||
        (page.slideId !== undefined && !hostId(page.slideId)) ||
        (page.startedAt !== undefined && !timestamp(page.startedAt)) ||
        (recorded &&
          (!timestamp(page.completedAt) ||
            !hostId(page.slideId) ||
            (page.startedAt !== undefined && page.startedAt > page.completedAt))) ||
        (!recorded && (page.completedAt !== undefined || !timestamp(page.startedAt)))
      )
        continue
      const id = `host-import:${JSON.stringify([project.projectId, production!.requestId, page.id, page.slideId ?? null, ...(recorded ? [] : [page.startedAt])])}`
      const records: NonNullable<PresentationWorkflowSummary['timeline'][number]['records']> = []
      if (page.startedAt)
        records.push({
          id: `${id}:started`,
          type: 'host.import.started',
          at: page.startedAt,
          text: `页面 ${page.title} 的原宿主写入开始记录；无结束回执不能证明仍在执行。`,
        })
      if (recorded)
        records.push({
          id: `${id}:recorded`,
          type: 'host.import.recorded',
          at: page.completedAt!,
          text: `页面 ${page.title} 的导入结果已记录或核对认领；此为记录完成时间，不代表原宿主写入完成时刻或 QA 通过。`,
        })
      timeline.push({
        id,
        scope: 'host_page_import',
        type: recorded ? 'host.import.recorded' : 'host.import.uncertain',
        text: recorded
          ? `已记录导入页面 ${page.title} · 宿主页 ${page.slideId}；不代表视觉或 QA 验收通过`
          : `页面 ${page.title} 的写入结果待核查；检查宿主页后再继续，不自动重放写入`,
        at: recorded ? page.completedAt : page.startedAt,
        recordsLabel: '宿主页导入记录',
        records,
      })
    }
  }
  if (qaMatches) {
    timeline.push({
      id: 'qa',
      text: `历史页面审查：${reviewed}/${qa!.pages.length} 页结构与视觉通过`,
    })
  }
  if (
    production?.projectId === project.projectId &&
    qa?.source === 'production' &&
    qa.projectId === project.projectId &&
    qa.requestId === production.requestId &&
    validatePresentationQaRecord(qa) &&
    qa.pages.every((page) => pageIds?.includes(page.pageId)) &&
    (!reportMatches || qa.documentId === report!.documentId)
  ) {
    for (const page of qa.pages) {
      if (importMatches) {
        const host = imported!.pages.find((item) => item.id === page.pageId)?.slideId
        if (host !== page.hostSlideId) continue
      }
      const id = `page-qa:${JSON.stringify([qa.documentId, qa.projectId, qa.requestId, page.pageId, page.hostSlideId, page.capturedAt, page.screenshotDigest])}`
      const fallback = page.screenshotRenderer === 'libreoffice'
      const caveat = `${fallback ? 'LibreOffice 备用预览；宿主外观未验；' : ''}历史记录不代表当前宿主外观、专业事实或保存重开验收通过`
      const records: NonNullable<PresentationWorkflowSummary['timeline'][number]['records']> = [
        {
          id: `${id}:capture`,
          type: 'qa.capture.recorded',
          at: page.capturedAt,
          text: `已记录页面 ${page.title} 的历史${fallback ? 'LibreOffice 备用预览' : '宿主截图'}；${caveat}`,
        },
      ]
      if (page.visual.reviewedAt)
        records.push({
          id: `${id}:visual:${page.visual.reviewedAt}`,
          type: 'qa.visual.recorded',
          at: page.visual.reviewedAt,
          text: `Agent 历史视觉复核：${page.visual.status === 'pass' ? '通过' : '需修改'}；${page.visual.notes}；${caveat}`,
        })
      if (page.invalidatedAt)
        records.push({
          id: `${id}:invalidated:${page.invalidatedAt}`,
          type: 'qa.evidence.invalidated',
          at: page.invalidatedAt,
          text: `页面 ${page.title} 的截图与复核证据首次失效；需重新采集和核对，不代表原记录丢失。`,
        })
      const latest = records.at(-1)!
      timeline.push({
        id,
        scope: 'saved_page_qa',
        type: latest.type,
        at: latest.at,
        text: `页面 ${page.title}：${page.invalidatedAt ? 'QA 证据已记录失效' : page.visual.reviewedAt ? `已记录 Agent 历史视觉复核：${page.visual.status === 'pass' ? '通过' : '需修改'}` : '已记录历史截图'}${page.recheckRequired ? '；结果已失效，需重审' : ''}${page.recheckRequired && !page.invalidatedAt ? '（失效时间未知）' : ''}；${caveat}`,
        recordsLabel: '页面 QA 记录',
        records,
      })
    }
  }
  if (qaAttempts) {
    const attempts = qaAttempts.attempts ?? []
    const unavailable =
      qaAttempts.unavailable === true ||
      !Array.isArray(attempts) ||
      attempts.length > 64 ||
      attempts.some((attempt) => !validatePresentationQaAttempt(attempt)) ||
      new Set(attempts.map((attempt) => attempt?.id)).size !== attempts.length
    if (unavailable)
      attention.push({
        id: 'qa-attempt-history-unavailable',
        text: '截图尝试历史暂不可读取；不沿用未知尝试结果，请刷新核对。',
      })
    else if (production?.projectId === project.projectId) {
      const documentId = reportMatches
        ? report!.documentId
        : qa?.projectId === project.projectId &&
            qa.requestId === production.requestId &&
            validatePresentationQaRecord(qa)
          ? qa.documentId
          : undefined
      for (const attempt of attempts) {
        if (
          !validatePresentationQaAttempt(attempt) ||
          attempt.source !== 'production' ||
          attempt.projectId !== project.projectId ||
          attempt.requestId !== production.requestId ||
          !pageIds?.includes(attempt.pageId) ||
          (documentId !== undefined && attempt.documentId !== documentId)
        )
          continue
        if (
          importMatches &&
          imported!.pages.find((page) => page.id === attempt.pageId)?.slideId !==
            attempt.hostSlideId
        )
          continue
        const id = `qa-attempt:${JSON.stringify([attempt.documentId, attempt.projectId, attempt.requestId, attempt.pageId, attempt.hostSlideId, attempt.artifactDigest, attempt.id, attempt.startedAt])}`
        const text = {
          started: '开始记录尚未闭合，不能证明仍在执行；不自动重放',
          recorded: '已保存截图；不代表图片已收到、视觉或专业 QA 通过',
          waiting: '等待截图能力；可在能力恢复后按明确请求重试',
          failed: '采集未完成；请核对状态后再决定是否重试',
          cancelled: '采集已取消；未认证截图或页面状态',
          closed: '未决记录已结束；不代表宿主操作已取消、截图或 QA 通过',
        }[attempt.status]
        const records: NonNullable<PresentationWorkflowSummary['timeline'][number]['records']> = [
          {
            id: `${id}:started`,
            type: 'qa.attempt.started',
            at: attempt.startedAt,
            text: '截图采集尝试的真实开始记录；不证明当前仍在执行。',
          },
        ]
        if (attempt.finishedAt)
          records.push({
            id: `${id}:finished:${attempt.finishedAt}`,
            type: `qa.attempt.${attempt.status}`,
            at: attempt.finishedAt,
            text,
          })
        timeline.push({
          id,
          type: `qa.attempt.${attempt.status}`,
          scope: 'page_qa_attempt',
          at: attempt.finishedAt ?? attempt.startedAt,
          text: `页面 ${attempt.pageId} · ${text}；不认证当前宿主外观或保存重开结果`,
          recordsLabel: '截图尝试记录',
          records,
        })
      }
    }
  }
  if (reportMatches) {
    timeline.push({ id: 'report', text: `内容证据报告：${openIssues} 项问题待处理` })
    for (const review of report!.reviews.slice(-20)) {
      const outcome = {
        supported: '支持',
        contradicted: '冲突',
        insufficient_evidence: '证据不足',
      }[review.outcome]
      timeline.push({
        id: `evidence-${review.reviewId}`,
        text: `Agent 对页面 ${review.pageId} 的主张 ${review.claimId}、来源 ${review.sourceId} 记录历史判断：${outcome}；来源真实性仍需核验`,
        at: review.createdAt,
      })
    }
    const currentIssues = new Map(
      report!.pages.flatMap((page) =>
        page.issues.map((issue) => [issue.id, issue.digest] as const),
      ),
    )
    for (const action of report!.issueLedger.actions.slice(-20)) {
      const label = { open: '重新打开', deferred: '暂缓', explained: '记录解释' }[action.state]
      const stale = currentIssues.get(action.issueId) !== action.issueDigest
      timeline.push({
        id: `issue-${action.actionId}`,
        text: `内容问题 ${action.issueId}：${label}${stale ? '；对应证据已变化，需重新处理' : ''}；未验证来源真实性或结论`,
        at: action.createdAt,
      })
    }
  }
  if (delivery) {
    let unavailable = delivery.unavailable === true
    let bundles: PresentationDeliveryBundleReceipt[] = []
    if (!unavailable && delivery.bundles !== undefined) {
      try {
        if (!Array.isArray(delivery.bundles) || delivery.bundles.length > 32)
          throw Error('invalid_state')
        bundles = delivery.bundles.map(parsePresentationDeliveryBundleReceipt)
        if (
          new Set(bundles.map((bundle) => bundle.bundleId)).size !== bundles.length ||
          new Set(bundles.map((bundle) => bundle.documentId)).size > 1 ||
          bundles.some(
            (bundle) =>
              bundle.projectId !== project.projectId ||
              bundle.requestId !== production?.requestId ||
              bundle.manifest.planRevision !== production?.planRevision ||
              (reportMatches &&
                (bundle.documentId !== report!.documentId ||
                  bundle.manifest.inputDigest !== report!.inputDigest ||
                  bundle.manifest.planDigest !== report!.planDigest)),
          )
        )
          throw Error('invalid_state')
      } catch {
        unavailable = true
        bundles = []
      }
    }
    if (unavailable)
      attention.push({
        id: 'delivery-bundle-history-unavailable',
        text: '本机交付包历史暂不可读取；不沿用未知归档状态，已有文稿保留，请刷新核对。',
      })
    else {
      for (const bundle of bundles) {
        const id = `delivery-bundle:${JSON.stringify([bundle.documentId, bundle.projectId, bundle.requestId, bundle.bundleId, bundle.sha256])}`
        const ready = bundle.state === 'ready'
        const records = [
          {
            id: `${id}:upload`,
            at: bundle.createdAt,
            text: '本机交付包上传开始；没有结束回执不能证明仍在上传。',
          },
        ]
        if (bundle.completedAt)
          records.push({
            id: `${id}:ready`,
            at: bundle.completedAt,
            text: '本机交付包归档完成；不是专业内容或PowerPoint验收通过。',
          })
        timeline.push({
          id,
          type: ready ? 'delivery.bundle.ready' : 'delivery.bundle.started',
          scope: 'current_document_delivery_bundle',
          at: bundle.completedAt ?? bundle.createdAt,
          text: `${ready ? '本机交付包已归档' : '本机交付包上传未完成'} · 计划第 ${bundle.manifest.planRevision} 版 · PDF：${bundle.manifest.checks.pdf === 'included' ? '已包含宿主PDF' : bundle.manifest.checks.pdf === 'unavailable' ? '宿主PDF不可用' : '未请求'}；专业内容、来源权威性、时效、当前宿主验收及保存重开检查待完成`,
          recordsLabel: '交付包归档记录',
          records,
        })
      }
      if (bundles.some((bundle) => bundle.state === 'uploading'))
        attention.push({
          id: 'delivery-bundle-unfinished',
          text: '本机交付包有上传开始但未归档的记录；请只读刷新核对，不自动再次导出当前文稿。',
        })
    }
  }
  timeline.sort((a, b) => (a.at && b.at ? a.at.localeCompare(b.at) : a.at ? -1 : b.at ? 1 : 0))
  const nextAction = !plan
    ? '保存 Brief、资料、故事线与样式规范'
    : !production
      ? sourceProblems.length
        ? sourceProblems.some((source) =>
            ['excerpt_mismatch', 'excerpt_missing', 'source_mismatch', 'locator_mismatch'].includes(
              source.status,
            ),
          )
          ? '核对附件原文并修订计划摘录，再决定是否开始生产'
          : sourceProblems.some((source) => source.status === 'ready')
            ? '旧版 PC 尚未核对计划摘录；请升级 PC 或使用来源核对工具后再生产'
            : '检查并补齐计划引用的资料，再决定是否开始生产'
        : project.sourcePreparationUnavailable
          ? '资料状态暂不可读取；先刷新项目状态并核对计划引用资料'
          : '按已保存计划启动逐页生产'
      : planChangedSinceProduction
        ? '核对已保存的新计划，选择继续旧任务或按新计划重新生产'
        : production.revision
          ? '检查单页修订并按保存点确认宿主页替换'
          : project.productionJob &&
              ['running', 'pausing', 'cancelling'].includes(project.productionJob.state)
            ? '等待当前后台页任务完成或处理暂停/取消请求'
            : project.productionJob?.state === 'cancelled'
              ? '当前后台任务已取消；如需继续，请重新建立页生产任务'
              : failed
                ? '修复失败页后继续生产'
                : production.compiledCount < production.total
                  ? '继续生产剩余页面'
                  : !importMatches
                    ? '准备成果并确认逐页导入'
                    : imported!.status === 'uncertain'
                      ? '先检查写入结果不确定的页面'
                      : imported!.completed < imported!.total
                        ? '继续导入剩余页面'
                        : !qaMatches
                          ? '采集页面截图并记录审查'
                          : qa!.pages.some((page) => page.recheckRequired)
                            ? '重审受影响页面'
                            : fallbackReviewed && reviewed + fallbackReviewed === qa!.pages.length
                              ? '在 PowerPoint 中重新截图并核验备用预览页面的宿主外观'
                              : reviewed < qa!.pages.length
                                ? '处理未通过或待审页面'
                                : !reportMatches
                                  ? '读取当前任务的内容证据交付报告'
                                  : openIssues
                                    ? '处理内容证据报告中的待处理问题'
                                    : '继续来源核验与保存重开验收'
  let nextTool: PresentationWorkflowSummary['nextTool']
  if (plan && production && !production.revision && !planChangedSinceProduction) {
    if (
      (failed || production.compiledCount < production.total) &&
      ['paused', 'interrupted', 'failed'].includes(project.productionJob?.state ?? '')
    )
      nextTool = 'resume_job'
    else if ((failed || production.compiledCount < production.total) && !project.productionJob)
      nextTool = project.jobsUnavailable ? 'run_pages' : 'start_job'
    else if (production.status === 'compiled' && !importMatches) nextTool = 'prepare_import'
    else if (
      importMatches &&
      imported?.status === 'complete' &&
      qaMatches &&
      reviewed === qa!.pages.length &&
      !reportMatches
    )
      nextTool = 'read_report'
  }
  return { stages, timeline, attention, pages, nextAction, ...(nextTool ? { nextTool } : {}) }
}
