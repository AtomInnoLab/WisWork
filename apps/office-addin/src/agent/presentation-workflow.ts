import type { PresentationProjectStatus } from '../skills/powerpoint/presentation-project.js'
import type { PresentationImportProgress } from '../skills/powerpoint/presentation-page-delivery.js'
import type { PresentationQaRecord } from '../skills/powerpoint/presentation-qa.js'
import type { PresentationDeliveryReport } from '@wiswork/pptx-engine/presentation-delivery-report'

export interface PresentationWorkflowSummary {
  stages: { name: string; detail: string }[]
  pages: { id: string; title: string; production: string; imported: string; qa: string; nextAction: string }[]
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
  const importMatches = Boolean(production && imported?.source === 'production' &&
    imported.projectId === project.projectId && imported.requestId === production.requestId &&
    imported.total === production.total &&
    JSON.stringify(imported.pages.map((page) => page.id)) === JSON.stringify(pageIds))
  const qaMatches = Boolean(production && qa?.source === 'production' &&
    qa.projectId === project.projectId && qa.requestId === production.requestId &&
    qa.pages.length === production.total &&
    qa.pages.every((page) => pageIds?.includes(page.pageId)))
  const reviewed = qaMatches ? qa!.pages.filter((page) =>
    !page.recheckRequired && page.structure.status === 'passed' && page.visual.status === 'pass').length : 0
  const reportMatches = Boolean(production && report &&
    report.projectId === project.projectId && report.requestId === production.requestId &&
    report.planRevision === production.planRevision)
  const openIssues = reportMatches ? report!.pages.reduce((sum, page) =>
    sum + page.issues.filter((issue) => issue.disposition.state === 'open' || issue.disposition.stale).length, 0) : 0
  const failed = production?.pages.filter((page) => page.state === 'failed').length ?? 0
  const pages = (production?.pages ?? project.slides).map((slide) => {
    const page = production?.pages.find((item) => item.id === slide.id)
    const index = production?.pages.findIndex((item) => item.id === slide.id) ?? -1
    const importedPage = importMatches && index >= 0 ? imported!.pages[index] : undefined
    const reviewedPage = qaMatches ? qa!.pages.find((item) => item.pageId === slide.id) : undefined
    const productionText = page ? {
      pending: '待编译', building: '编译中', failed: '编译失败', compiled: '已编译',
    }[page.state] : '待创建页任务'
    const importText = importedPage ? {
      pending: '待导入', uncertain: '写入待核查', complete: '已记录导入',
    }[importedPage.state] : '无当前任务导入记录'
    const qaText = reviewedPage
      ? reviewedPage.recheckRequired ? '历史检查已失效' :
        reviewedPage.visual.status === 'pass' && reviewedPage.structure.status === 'passed'
          ? '历史结构与视觉通过' : reviewedPage.visual.status === 'needs_changes'
            ? '历史检查需修改' : '待完成页面复核'
      : '无当前任务 QA 记录'
    const pageNext = !page || page.state === 'pending' ? '制作页面' :
      page.state === 'building' ? '等待当前编译' :
      page.state === 'failed' ? '修复并重试此页' :
      !importedPage || importedPage.state === 'pending' ? '导入此页' :
      importedPage.state === 'uncertain' ? '检查宿主页后再继续' :
      !reviewedPage ? '采集截图并复核' :
      reviewedPage.recheckRequired ? '重审此页' :
      reviewedPage.visual.status === 'needs_changes' ? '修改页面后重审' :
      qaText === '历史结构与视觉通过' ? '继续来源与保存重开核验' : '完成页面复核'
    return { id: slide.id, title: slide.title, production: productionText, imported: importText,
      qa: qaText, nextAction: pageNext }
  })
  const stages = [
    { name: '目标与资料', detail: plan
      ? `Brief 已保存；登记 ${plan.sources.length} 份资料、${plan.claims.length} 条主张，真实性仍需核验`
      : '尚无已保存的结构化计划' },
    { name: '故事线与样式', detail: plan
      ? `已保存 ${plan.slides.length} 页施工图和样式契约 · 计划第 ${project.plan!.revision} 版`
      : '待保存逐页施工图与样式契约' },
    { name: '逐页制作', detail: production
      ? `已编译 ${production.compiledCount}/${production.total} 页${failed ? ` · ${failed} 页失败待重试` : ''}；尚不代表导入或验收`
      : project.status === 'compiled' ? '旧编译成果可用；尚无逐页生产记录' : '尚未启动逐页生产' },
    { name: '导入 PowerPoint', detail: importMatches
      ? `已记录 ${imported!.completed}/${imported!.total} 页；${imported!.status === 'uncertain' ? '写入结果不确定，需检查文档' : '导入不代表视觉验收'}`
      : imported ? '现有导入记录无法与当前页任务匹配，需核对' : '尚无当前页任务的导入记录' },
    { name: '页面审查', detail: qaMatches
      ? `历史结构与视觉复核 ${reviewed}/${qa!.pages.length} 页通过；${qa!.pages.filter((page) => page.recheckRequired).length} 页需重审`
      : qa ? '现有 QA 记录无法与当前页任务匹配，需核对' : '尚无当前页任务的 QA 记录' },
    { name: '交付核验', detail: `${reportMatches ? `内容证据报告有 ${openIssues} 项待处理；` : '尚无当前任务的内容证据报告；'}来源真实性、保存重开及真实 PowerPoint 验收尚不能由上述记录证明` },
  ]
  const nextAction = !plan
    ? '保存 Brief、资料、故事线与样式规范'
      : !production
        ? '按已保存计划启动逐页生产'
        : production.revision
          ? '检查单页修订并按保存点确认宿主页替换'
          : project.productionJob && ['running', 'pausing', 'cancelling'].includes(project.productionJob.state)
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
    if ((failed || production.compiledCount < production.total) &&
      ['paused', 'interrupted', 'failed'].includes(project.productionJob?.state ?? ''))
      nextTool = 'resume_job'
    else if ((failed || production.compiledCount < production.total) && !project.productionJob)
      nextTool = project.jobsUnavailable ? 'run_pages' : 'start_job'
    else if (production.status === 'compiled' && !importMatches)
      nextTool = 'prepare_import'
    else if (importMatches && imported?.status === 'complete' && qaMatches &&
      reviewed === qa!.pages.length && !reportMatches)
      nextTool = 'read_report'
  }
  return { stages, pages, nextAction, ...(nextTool ? { nextTool } : {}) }
}
