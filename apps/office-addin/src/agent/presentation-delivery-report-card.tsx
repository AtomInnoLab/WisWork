import { useState } from 'react'
import type {
  PresentationDeliveryReport,
  DeliveryIssue,
} from '@wiswork/pptx-engine/presentation-delivery-report'
import type { PresentationProjectController } from '../skills/powerpoint/presentation-project.js'

function IssueForm({
  issue,
  disabled,
  controller,
}: {
  issue: DeliveryIssue
  disabled: boolean
  controller: PresentationProjectController
}) {
  const [state, setState] = useState<'open' | 'deferred' | 'explained'>('open')
  const [note, setNote] = useState('')
  return (
    <form
      onSubmit={(event) => {
        event.preventDefault()
        if (note.trim())
          void controller.recordIssueAction({
            actionId: crypto.randomUUID(),
            issueId: issue.id,
            issueDigest: issue.digest,
            state,
            note: note.trim(),
          })
      }}
    >
      <label>
        处置状态
        <select
          aria-label={`处置状态 ${issue.id}`}
          value={state}
          disabled={disabled}
          onChange={(event) => setState(event.target.value as typeof state)}
        >
          <option value="open">待处理</option>
          <option value="deferred">暂缓</option>
          <option value="explained">已说明</option>
        </select>
      </label>
      <label>
        处置理由
        <textarea
          aria-label={`处置理由 ${issue.id}`}
          value={note}
          maxLength={2000}
          required
          disabled={disabled}
          onChange={(event) => setNote(event.target.value)}
        />
      </label>
      <button type="submit" disabled={disabled || !note.trim()}>
        记录处置
      </button>
    </form>
  )
}
function IssueList({
  category,
  issues,
  report,
  controller,
  disabled,
}: {
  category: DeliveryIssue['category']
  issues: DeliveryIssue[]
  report: PresentationDeliveryReport
  controller: PresentationProjectController
  disabled: boolean
}) {
  const [limit, setLimit] = useState(20)
  return (
    <div>
      <h5>
        {category === 'needs_human' ? '待人工判断' : '无法核验'} · {issues.length}
      </h5>
      <ul>
        {issues.slice(0, limit).map((issue) => (
          <li key={issue.id}>
            <p>
              {issue.code} · 主张 {issue.claimId}：
              {report.plan.claims.find((claim) => claim.id === issue.claimId)?.statement}
              {issue.sourceId && (
                <>
                  {' · 来源 '}
                  <a href={`#evidence-source-${report.requestId}-${issue.sourceId}`}>
                    {issue.sourceId}
                  </a>
                </>
              )}{' '}
              · {issue.disposition.state}
              {issue.disposition.stale ? ' · 原处置已过期，当前待处理' : ''}
            </p>
            <IssueForm
              key={`${issue.id}-${issue.digest}-${report.issueLedger.revision}`}
              issue={issue}
              controller={controller}
              disabled={disabled}
            />
          </li>
        ))}
      </ul>
      {issues.length > limit && (
        <button type="button" onClick={() => setLimit(Math.min(limit + 20, issues.length))}>
          显示更多问题 · 剩余 {issues.length - limit} 项
        </button>
      )}
    </div>
  )
}
export function PresentationDeliveryReportCard({
  report,
  controller,
  disabled,
}: {
  report: PresentationDeliveryReport
  controller: PresentationProjectController
  disabled: boolean
}) {
  return (
    <section aria-label="内容证据交付报告">
      <h3>内容证据交付报告 · {report.requestId}</h3>
      <p>
        仅对应冻结计划第 {report.planRevision}{' '}
        版。已核验仅表示算术复现；来源真实性、时效未核验，宿主检查与 Office
        往返未执行。附件字面核对反映读取报告时的当前文档，不属于冻结生产快照。来源复核是历史 Agent
        判断。已说明不会关闭机器发现。
      </p>
      {report.pages.map((page) => (
        <section key={page.pageId} aria-label={`证据页面 ${page.title}`}>
          <h4>
            {page.title} · {page.pageId} · {page.productionState}
          </h4>
          <p>
            已核验（仅算术复现）：
            {page.calculations.filter((result) => result.status === 'reproduced').length}
          </p>
          <ul>
            {page.calculations.map((result) => (
              <li key={result.claimId}>
                主张 {result.claimId}：{result.status}
                {result.actual !== undefined
                  ? ` · 复算 ${result.actual} · 声明 ${result.expected} · 容差 ${result.tolerance}`
                  : ''}
              </li>
            ))}
          </ul>
          {(['needs_human', 'unverifiable'] as const).map((category) => {
            const issues = page.issues.filter((issue) => issue.category === category)
            return (
              <IssueList
                key={`${report.requestId}-${report.planRevision}-${page.pageId}-${category}`}
                category={category}
                issues={issues}
                report={report}
                controller={controller}
                disabled={disabled}
              />
            )
          })}
        </section>
      ))}
      <details>
        <summary>来源目录 · {report.plan.sources.length} 条</summary>
        {report.plan.sources.map((source) => (
          <section key={source.id} id={`evidence-source-${report.requestId}-${source.id}`}>
            <h4>
              {source.title} · {source.id}
            </h4>
            <p>来源：{source.uri}</p>
            <p>原文位置：{source.locator?.trim() || '未提供'}</p>
            <p>资料时点：{source.asOf?.trim() || '未提供'}</p>
            <p>原文片段：{source.excerpt.trim() || '未提供'}</p>
            {report.sourceAudit?.find((item) => item.sourceId === source.id) && (
              <p>
                当前附件字面核对：
                {
                  {
                    found: '找到原文片段',
                    not_found: '完整原文中未找到片段',
                    empty_excerpt: '计划未填写原文片段',
                    not_ready: '附件尚未解析就绪',
                    unsupported: '附件不是可读文本',
                    missing: '当前文档缺少附件',
                  }[report.sourceAudit.find((item) => item.sourceId === source.id)!.status]
                }
                。仅核对字面存在，不核验事实支持、来源权威性或时效。
              </p>
            )}
          </section>
        ))}
      </details>
      <details>
        <summary>完整处置历史 · {report.issueLedger.actions.length} 条</summary>
        <ul>
          {report.issueLedger.actions.map((action) => (
            <li key={action.actionId}>
              {action.sequence} · {action.issueId} · {action.state} · {action.createdAt} ·{' '}
              {action.note}
            </li>
          ))}
        </ul>
      </details>
    </section>
  )
}
