import { useSyncExternalStore } from 'react'
import type { PresentationProjectController } from '../skills/powerpoint/presentation-project.js'
import type { PresentationImportProgressController } from './presentation-import-progress.js'
import type { PresentationQaController } from './presentation-qa-card.js'
import { presentationWorkflowSummary } from './presentation-workflow.js'

const noSubscribe = () => () => {}
const zero = () => 0

export function PresentationWorkflowCard({
  project,
  imported,
  qa,
  disabled = false,
}: {
  project: PresentationProjectController
  imported?: PresentationImportProgressController
  qa?: PresentationQaController
  disabled?: boolean
}) {
  const snapshot = useSyncExternalStore(project.subscribe, project.snapshot, project.snapshot)
  useSyncExternalStore(
    imported?.subscribe ?? noSubscribe,
    imported?.revision ?? zero,
    imported?.revision ?? zero,
  )
  useSyncExternalStore(qa?.subscribe ?? noSubscribe, qa?.revision ?? zero, qa?.revision ?? zero)
  let importRecord: ReturnType<PresentationImportProgressController['read']>
  let qaRecord: ReturnType<PresentationQaController['read']>
  try {
    importRecord = imported?.read()
  } catch {
    /* The dedicated import card shows the read error. */
  }
  try {
    qaRecord = qa?.read()
  } catch {
    /* The dedicated QA card shows the read error. */
  }
  const workflow = presentationWorkflowSummary(
    snapshot.project,
    importRecord,
    qaRecord,
    snapshot.deliveryReport,
  )
  if (!workflow) return null
  const runNext = () => {
    const requestId = snapshot.project?.production?.requestId
    switch (workflow.nextTool) {
      case 'start_job':
        if (requestId) void project.startProductionJob(requestId)
        break
      case 'resume_job':
        if (requestId) void project.resumeProductionJob(requestId)
        break
      case 'run_pages':
        if (requestId) void project.runProduction(requestId)
        break
      case 'prepare_import':
        void project.prepareProduction()
        break
      case 'read_report':
        void project.readDeliveryReport()
        break
    }
  }
  const actionLabels = {
    start_job: '后台制作剩余页面',
    resume_job: '继续后台任务',
    run_pages: '继续页任务',
    prepare_import: '准备逐页导入',
    read_report: '读取内容证据报告',
  }
  return (
    <section className="presentation-project" aria-label="演示文稿制作阶段">
      <strong>制作流程</strong>
      <p role="status">下一步：{workflow.nextAction}</p>
      {workflow.nextTool && (
        <button type="button" disabled={disabled || snapshot.phase !== 'idle'} onClick={runNext}>
          {actionLabels[workflow.nextTool]}
        </button>
      )}
      {workflow.attention.length > 0 && (
        <section aria-label="待处理问题">
          <strong>待处理 · {workflow.attention.length} 项</strong>
          <ul>
            {workflow.attention.map((item) => (
              <li key={item.id}>{item.text}</li>
            ))}
          </ul>
        </section>
      )}
      <ol>
        {workflow.stages.map((stage) => (
          <li key={stage.name}>
            <strong>{stage.name}</strong>：{stage.detail}
          </li>
        ))}
      </ol>
      <details>
        <summary>恢复记录 · {workflow.timeline.length} 项</summary>
        <p>
          根据已保存记录重建；计划修订、页任务、逐页导入、历史
          QA、证据判断和内容问题处置带有记录时间，其他条目是当前检查点。
        </p>
        <ol>
          {workflow.timeline.map((event) => (
            <li key={event.id}>
              {event.at ? <time dateTime={event.at}>{event.at}</time> : null} {event.text}
            </li>
          ))}
        </ol>
      </details>
      <details>
        <summary>逐页状态 · {workflow.pages.length} 页</summary>
        <ol>
          {workflow.pages.map((page) => (
            <li key={page.id}>
              <strong>{page.title}</strong> · {page.id}：{page.production} / {page.imported} /{' '}
              {page.qa} / {page.evidence}；下一步：{page.nextAction}
            </li>
          ))}
        </ol>
      </details>
    </section>
  )
}
