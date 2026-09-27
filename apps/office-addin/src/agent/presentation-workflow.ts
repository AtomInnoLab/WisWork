import type { PresentationProjectStatus } from '../skills/powerpoint/presentation-project.js'
import type { PresentationImportProgress } from '../skills/powerpoint/presentation-page-delivery.js'
import type { PresentationQaRecord } from '../skills/powerpoint/presentation-qa.js'
import type { PresentationDeliveryReport } from '@wiswork/pptx-engine/presentation-delivery-report'

export interface PresentationWorkflowSummary {
  stages: { name: string; detail: string }[]
  timeline: { id: string; text: string; at?: string }[]
  attention: { id: string; text: string }[]
  pages: {
    id: string
    title: string
    production: string
    imported: string
    qa: string
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
): PresentationWorkflowSummary | undefined {
  if (!project) return undefined
  const plan = project.plan?.value
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
  const pages = (production?.pages ?? project.slides).map((slide) => {
    const page = production?.pages.find((item) => item.id === slide.id)
    const index = production?.pages.findIndex((item) => item.id === slide.id) ?? -1
    const importedPage = importMatches && index >= 0 ? imported!.pages[index] : undefined
    const reviewedPage = qaMatches ? qa!.pages.find((item) => item.pageId === slide.id) : undefined
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
        : reviewedPage.visual.status === 'pass' && reviewedPage.structure.status === 'passed'
          ? '历史结构与视觉通过'
          : reviewedPage.visual.status === 'needs_changes'
            ? '历史检查需修改'
            : '待完成页面复核'
      : '无当前任务 QA 记录'
    const pageNext =
      !page || page.state === 'pending'
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
                        ? '继续来源与保存重开核验'
                        : '完成页面复核'
    return {
      id: slide.id,
      title: slide.title,
      production: productionText,
      imported: importText,
      qa: qaText,
      nextAction: pageNext,
    }
  })
  const stages = [
    {
      name: '目标与资料',
      detail: plan
        ? `Brief 已保存；登记 ${plan.sources.length} 份资料、${plan.claims.length} 条主张，真实性仍需核验`
        : '尚无已保存的结构化计划',
    },
    {
      name: '故事线与样式',
      detail: plan
        ? `已保存 ${plan.slides.length} 页施工图和样式契约 · 计划第 ${project.plan!.revision} 版`
        : '待保存逐页施工图与样式契约',
    },
    {
      name: '逐页制作',
      detail: production
        ? `已编译 ${production.compiledCount}/${production.total} 页${failed ? ` · ${failed} 页失败待重试` : ''}；尚不代表导入或验收`
        : project.status === 'compiled'
          ? '旧编译成果可用；尚无逐页生产记录'
          : '尚未启动逐页生产',
    },
    {
      name: '导入 PowerPoint',
      detail: importMatches
        ? `已记录 ${imported!.completed}/${imported!.total} 页；${imported!.status === 'uncertain' ? '写入结果不确定，需检查文档' : '导入不代表视觉验收'}`
        : imported
          ? '现有导入记录无法与当前页任务匹配，需核对'
          : '尚无当前页任务的导入记录',
    },
    {
      name: '页面审查',
      detail: qaMatches
        ? `历史结构与视觉复核 ${reviewed}/${qa!.pages.length} 页通过；${qa!.pages.filter((page) => page.recheckRequired).length} 页需重审`
        : qa
          ? '现有 QA 记录无法与当前页任务匹配，需核对'
          : '尚无当前页任务的 QA 记录',
    },
    {
      name: '交付核验',
      detail: `${reportMatches ? `内容证据报告有 ${openIssues} 项待处理；` : '尚无当前任务的内容证据报告；'}来源真实性、保存重开及真实 PowerPoint 验收尚不能由上述记录证明`,
    },
  ]
  // Rebuild from durable records. Undated entries are current checkpoints, not events.
  const timeline: PresentationWorkflowSummary['timeline'] = []
  if (plan) {
    const revisions = project.plan!.revisions
    if (revisions?.length) {
      for (const [index, event] of revisions.entries()) {
        const current = event.snapshot
        const previous = revisions[index - 1]?.snapshot
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
          text: `已保存计划第 ${event.revision} 版${detail}；来源真实性仍需核验`,
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
    if (
      project.productionJob?.requestId === production.requestId &&
      project.productionJob.projectId === project.projectId
    ) {
      const labels: Record<string, string> = {
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
      }
      for (const event of project.productionJob.events.slice(-20)) {
        const pageText = 'pageId' in event ? ` · ${event.pageId}（第 ${event.attempt} 次）` : ''
        timeline.push({
          id: `job-${event.sequence}`,
          text: `${labels[event.type]}${pageText}`,
          at: event.createdAt,
        })
      }
    }
  }
  if (importMatches) {
    timeline.push({
      id: 'import',
      text: `导入检查点：${imported!.completed}/${imported!.total} 页${imported!.status === 'uncertain' ? '，有写入待核查' : ''}`,
    })
    for (const page of imported!.pages)
      if (page.state === 'complete' && page.completedAt)
        timeline.push({
          id: `import-${page.id}`,
          text: `已记录导入页面 ${page.title}；尚未完成视觉验收`,
          at: page.completedAt,
        })
    for (const page of imported!.pages)
      if (page.state === 'uncertain' && page.startedAt)
        timeline.push({
          id: `import-uncertain-${page.id}`,
          text: `页面 ${page.title} 的写入结果待核查；检查宿主页后再继续`,
          at: page.startedAt,
        })
  }
  if (qaMatches) {
    timeline.push({
      id: 'qa',
      text: `历史页面审查：${reviewed}/${qa!.pages.length} 页结构与视觉通过`,
    })
    for (const page of qa!.pages) {
      timeline.push({
        id: `capture-${page.pageId}`,
        text: `已采集页面 ${page.title} 的历史截图${page.recheckRequired ? '；需重审' : ''}`,
        at: page.capturedAt,
      })
      if (page.visual.reviewedAt)
        timeline.push({
          id: `review-${page.pageId}`,
          text: `已记录页面 ${page.title} 的历史视觉复核：${page.visual.status === 'pass' ? '通过' : '需修改'}${page.recheckRequired ? '；结果已失效' : ''}`,
          at: page.visual.reviewedAt,
        })
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
  timeline.sort((a, b) => (a.at && b.at ? a.at.localeCompare(b.at) : a.at ? -1 : b.at ? 1 : 0))
  const nextAction = !plan
    ? '保存 Brief、资料、故事线与样式规范'
    : !production
      ? '按已保存计划启动逐页生产'
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
                          : reviewed < qa!.pages.length
                            ? '处理未通过或待审页面'
                            : !reportMatches
                              ? '读取当前任务的内容证据交付报告'
                              : openIssues
                                ? '处理内容证据报告中的待处理问题'
                                : '继续来源核验与保存重开验收'
  let nextTool: PresentationWorkflowSummary['nextTool']
  if (plan && production && !production.revision) {
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
