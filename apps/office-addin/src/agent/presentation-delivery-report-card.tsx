import { useState } from 'react'
import {
  presentationProfessionalContextMissingFields,
  type PresentationProfessionalContext,
} from '@wiswork/project-store/presentation-professional-context'
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
          <option value="open">待处理（重新打开）</option>
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
const researchReasons: Record<string, string> = {
  research_unmapped_claim: '计划主张未映射研究',
  research_conflict_partner_omitted: '未选用冲突另一方',
  research_source_reference_unselected: '研究引用未选入计划',
  research_source_unavailable: '原文尚不可用',
  research_claim_conflict: '研究结论存在冲突',
}
const sourceAssessmentReasons: Record<string, string> = {
  source_authority_review_missing: '缺少来源权威性判断',
  source_authority_review_uncertain: '来源权威性尚不确定',
  source_authority_review_insufficient: '来源权威性不足以支持该主张',
  source_authority_review_mixed: '来源权威性判断或声明级别不同',
  source_timeliness_review_missing: '缺少来源时效判断',
  source_timeliness_review_uncertain: '来源时效尚不确定',
  source_timeliness_review_historical_only: '来源仅适用于历史时点',
  source_timeliness_review_superseded: '来源已被后续资料取代',
  source_timeliness_review_mixed: '来源时效判断或比较框架不同',
  source_jurisdiction_review_missing: '缺少来源适用范围判断',
  source_jurisdiction_review_uncertain: '来源适用范围尚不确定',
  source_jurisdiction_review_mismatch: '来源与主张适用范围不匹配',
  source_jurisdiction_review_mixed: '来源适用范围判断不同',
}
function SourceAssessmentHistory({
  report,
  pageId,
}: {
  report: PresentationDeliveryReport
  pageId: string
}) {
  const reviews = (report.reviews ?? []).filter(
    (review) =>
      review.pageId === pageId && review.requestId === report.requestId && review.sourceAssessment,
  )
  const groups = new Map<string, typeof reviews>()
  for (const review of reviews) {
    const key = `${review.claimId}:${review.sourceId}`
    const group = groups.get(key) ?? []
    group.push(review)
    groups.set(key, group)
  }
  return (
    <>
      {Array.from(groups, ([key, items]) => (
        <details key={key} aria-label={`来源评估历史 ${items[0]!.claimId} ${items[0]!.sourceId}`}>
          <summary>
            来源评估历史 · 主张 {items[0]!.claimId} · 来源 {items[0]!.sourceId} · {items.length} 条
          </summary>
          <p>
            全部为历史 Agent
            判断，保留每个原文窗口，不以最新判断覆盖旧判断。正向判断仍不证明事实成立或来源已认证，全局权威性与时效未核验。
          </p>
          <p>判断或比较框架不同，不代表事实矛盾；说明或暂缓不会关闭发现的问题。</p>
          {items.map((review) => {
            const assessment = review.sourceAssessment!
            return (
              <section key={review.reviewId}>
                <p>
                  原复核 ID：{review.reviewId} · {review.createdAt}
                </p>
                <p>
                  窗口 UTF-16 {review.offset} · 最多 {review.maxChars} 字符
                </p>
                <p>评估范围：{assessment.scope}</p>
                <p>
                  权威性：
                  {
                    {
                      appropriate_for_claim: '适合该主张',
                      insufficient_authority: '权威性不足',
                      uncertain: '不确定',
                    }[assessment.authority.outcome]
                  }
                  ； 声明来源级别：
                  {
                    {
                      primary: '一手来源',
                      authoritative_secondary: '权威二手来源',
                      secondary: '二手来源',
                      unverified: '未核验',
                    }[assessment.authority.sourceTier]
                  }
                  ； 理由：{assessment.authority.reason}
                </p>
                <p>
                  时效：
                  {
                    {
                      current_for_claim: '适用于该主张时点',
                      historical_only: '仅适用于历史时点',
                      superseded: '已被后续资料取代',
                      uncertain: '不确定',
                    }[assessment.timeliness.outcome]
                  }
                  ； 比较日期：{assessment.timeliness.referenceDate}；理由：
                  {assessment.timeliness.reason}
                </p>
                <p>
                  主张数据时点：{assessment.timeliness.claimAsOf ?? '未提供'}；资料数据时点：
                  {assessment.timeliness.sourceAsOf ?? '未提供'}
                </p>
                {assessment.jurisdiction && (
                  <p>
                    主张适用范围：{assessment.jurisdiction.claimJurisdiction}； 范围判断：
                    {
                      { applicable: '适用', mismatch: '不匹配', uncertain: '不确定' }[
                        assessment.jurisdiction.outcome
                      ]
                    }
                    ； 理由：{assessment.jurisdiction.reason}
                  </p>
                )}
                <p>原文依据只证明字面存在，不证明来源真实或结论正确。</p>
                {!assessment.basis.length && <p>未提供可核验原文依据。</p>}
                {assessment.basis.map((basis, index) => (
                  <blockquote key={index}>
                    UTF-16 {basis.offset}：{basis.text}
                  </blockquote>
                ))}
              </section>
            )
          })}
        </details>
      ))}
    </>
  )
}
const professionalReasons: Record<string, string> = {
  professional_context_missing: '专业领域主张缺少专业上下文',
  professional_context_incomplete: '专业上下文尚有缺失',
  professional_source_secondary: '专业结论来源为二手或未核验资料',
  professional_legal_rule_inactive: '法律材料不在明确适用日期范围内',
  professional_jurisdiction_mismatch: '通用与专业适用范围不同',
  professional_financial_time_mixed: '通用与专业数据时点不同',
  professional_financial_unit_mismatch: '计算与专业上下文单位不同',
  professional_financial_currency_mismatch: '计算与专业上下文币种不同',
}
const professionalLabels: Record<string, string> = {
  materialKind: '材料类型',
  publicationId: '出版或发布标识',
  version: '版本',
  sample: '样本',
  method: '方法',
  statisticalBasis: '统计依据',
  limitations: '局限',
  jurisdiction: '适用范围',
  effectLevel: '效力层级',
  effectiveFrom: '生效日期',
  effectiveUntil: '失效日期',
  applicabilityDate: '适用日期',
  caseNumber: '案号',
  originalLocation: '原文位置',
  reportingPeriod: '报告期间',
  asOf: '数据时点',
  currency: '币种',
  unit: '单位',
  accountingBasis: '会计口径',
  formula: '公式',
}
const professionalKinds: Record<string, string> = {
  paper: '论文',
  dataset: '数据集',
  standard: '标准',
  institution: '机构资料',
  statute: '法律条文',
  case: '判例',
  regulation: '法规',
  contract: '合同',
  disclosure: '披露资料',
  financial_statement: '财务报表',
  ir: '投资者关系资料',
  market_data: '市场数据',
}
function ProfessionalContext({
  context,
  claimId,
  claimType,
}: {
  context?: PresentationProfessionalContext
  claimId: string
  claimType?: string
}) {
  if (!context) return null
  const missing = presentationProfessionalContextMissingFields(context, claimType)
  return (
    <details aria-label={`专业上下文 ${claimId}`}>
      <summary>
        专业上下文 · {claimId} · {{ science: '科研', law: '法律', finance: '金融' }[context.domain]}
      </summary>
      <p>
        保留原专业限定，不补猜缺失值。完整字段仍不代表事实支持、权威认证或时效核验；说明与暂缓不会关闭问题。
      </p>
      {context.domain === 'law' && <p>未提供失效日期时，不据此推断法律材料仍然有效。</p>}
      {Object.entries(context)
        .filter(([field]) => field !== 'domain')
        .map(([field, value]) => (
          <p key={field}>
            {professionalLabels[field]}：
            {field === 'materialKind' ? professionalKinds[value] : value}
          </p>
        ))}
      <p>
        {missing.length
          ? `专业字段缺失：${missing.map((field) => professionalLabels[field]).join('、')}`
          : '未发现必要专业字段缺失，仍需人工核对。'}
      </p>
    </details>
  )
}
function ResearchIssueContext({
  issue,
  report,
}: {
  issue: DeliveryIssue
  report: PresentationDeliveryReport
}) {
  const binding = issue.research
  const record = report.research?.record
  if (
    !binding ||
    !record ||
    binding.ledgerId !== record.id ||
    binding.sequence !== record.sequence ||
    binding.draftDigest !== record.draftDigest
  )
    return null
  const claimIds = [binding.researchClaimId, ...binding.relatedClaimIds].filter(Boolean)
  const facts = record.draft.facts.filter((fact) => claimIds.includes(fact.claimId))
  const sourceIds = new Set([...binding.sourceIds, ...facts.flatMap((fact) => fact.sourceRefs)])
  return (
    <details aria-label="逐页研究问题上下文">
      <summary>原研究结论、冲突与来源</summary>
      <p>
        原研究 ID：{binding.ledgerId} · 版本 #{binding.sequence}
      </p>
      <p>已说明或暂缓仍保留研究缺口；重新打开继续审查，不代表事实支持、来源认证或核验完成。</p>
      {!binding.researchClaimId && <p>此计划主张没有绑定原研究结论，需人工核对。</p>}
      {facts.map((fact) => (
        <section key={fact.claimId}>
          <p>
            {fact.claimId === binding.researchClaimId ? '对应原结论' : '相关或相反结论'}{' '}
            {fact.claimId}：{fact.statement}
          </p>
          <p>
            类型：
            {
              {
                fact: '事实主张',
                quote: '引文',
                calculation: '计算',
                judgment: '判断',
                assumption: '假设',
              }[fact.type]
            }
            ； 声明来源级别：
            {
              {
                primary: '一手来源',
                authoritative_secondary: '权威二手来源',
                secondary: '二手来源',
                unverified: '未核验',
              }[fact.sourceTier]
            }
            ； 声明可信度：{{ high: '高', medium: '中', low: '低' }[fact.confidence]}
            ，仍待独立审查。
          </p>
          <p>原研究来源：{fact.sourceRefs.join('、') || '未提供'}</p>
          {fact.asOf && <p>数据时点：{fact.asOf}</p>}
          <ProfessionalContext
            context={fact.professionalContext}
            claimId={fact.claimId}
            claimType={fact.type}
          />
          {fact.jurisdiction && <p>适用范围：{fact.jurisdiction}</p>}
          {fact.calculation && (
            <p>
              计算：{fact.calculation.formula}；输入：{fact.calculation.inputs.join('、')}
              {fact.calculation.unit && `；单位：${fact.calculation.unit}`}
              {fact.calculation.currency && `；币种：${fact.calculation.currency}`}
            </p>
          )}
        </section>
      ))}
      {record.draft.sources
        .filter((source) => sourceIds.has(source.id))
        .map((source) => (
          <section key={source.id}>
            <p>
              原研究来源 {source.id}：
              {/^https?:/.test(source.uri) ? (
                <a href={source.uri} target="_blank" rel="noreferrer">
                  {source.title}
                </a>
              ) : (
                source.title
              )}
            </p>
            <blockquote>{source.excerpt || '未提供原文摘录'}</blockquote>
            <p>
              {record.sources?.find((item) => item.sourceId === source.id)?.status === 'found'
                ? '找到原文摘录，不证明主张成立'
                : '原文尚不可用，仍需核对完整资料'}
            </p>
            {source.locator && <p>原文位置：{source.locator}</p>}
            {source.asOf && <p>资料数据时点：{source.asOf}</p>}
          </section>
        ))}
    </details>
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
              {issue.code === 'source_original_not_frozen'
                ? '项目未保存该来源原文'
                : issue.code === 'source_url_mismatch'
                  ? '网页快照与计划网址不匹配'
                  : issue.code === 'source_locator_mismatch'
                    ? '计划定位与原文实际位置不匹配'
                    : (professionalReasons[issue.code] ??
                      sourceAssessmentReasons[issue.code] ??
                      researchReasons[issue.code] ??
                      issue.code)}{' '}
              · 主张 {issue.claimId}：
              {report.plan.claims.find((claim) => claim.id === issue.claimId)?.statement}
              {issue.sourceId && (
                <>
                  {' · 来源 '}
                  <a href={`#evidence-source-${report.requestId}-${issue.sourceId}`}>
                    {issue.sourceId}
                  </a>
                </>
              )}{' '}
              · {{ open: '待处理', deferred: '暂缓', explained: '已说明' }[issue.disposition.state]}
              {issue.disposition.stale ? ' · 原处置已过期，当前待处理' : ''}
            </p>
            <ResearchIssueContext issue={issue} report={report} />
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
const workflowTools: Record<string, string> = {
  read_research_ledger: '读取指定研究记录',
  save_presentation_plan: '保存演示计划',
  read_presentation_claim_evidence: '读取主张来源原文',
  record_presentation_claim_review: '记录历史来源判断',
  check_presentation_page_content: '检查页面内容',
  read_presentation_delivery_report: '读取内容证据报告',
}
function ProfessionalWorkflow({
  workflow,
}: {
  workflow?: PresentationDeliveryReport['professionalWorkflow']
}) {
  if (!workflow) return null
  return (
    <details aria-label="专业制作工作流" key={workflow.domain}>
      <summary>
        专业制作工作流 · {{ science: '科研', law: '法律', finance: '金融' }[workflow.domain]}
      </summary>
      <h4>来源优先级</h4>
      <ol>
        {workflow.sourcePriority.map((item) => (
          <li key={item}>{item}</li>
        ))}
      </ol>
      <h4>专业上下文字段</h4>
      <ul>
        {workflow.contextFields.map((item) => (
          <li key={item}>{professionalLabels[item] ?? item}</li>
        ))}
      </ul>
      <h4>复核步骤</h4>
      <ol>
        {workflow.reviewSteps.map((step) => (
          <li key={step.id}>
            <p>{step.title}</p>
            <p>{step.instruction}</p>
            <p>复核操作：{step.tools.map((tool) => workflowTools[tool] ?? tool).join('、')}</p>
          </li>
        ))}
      </ol>
      <h4>人工检查</h4>
      <ul>
        {workflow.manualChecks.map((item) => (
          <li key={item}>{item}</li>
        ))}
      </ul>
      <h4>范围说明</h4>
      <p>{workflow.disclosure}</p>
    </details>
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
      <ProfessionalWorkflow workflow={report.professionalWorkflow} />
      {report.research && (
        <details aria-label="冻结计划绑定研究">
          <summary>冻结计划绑定研究 #{report.research.record.sequence}</summary>
          <p>
            此处是冻结计划明确绑定的历史研究，不自动替换为最近整理结果；完成记录和引用映射不代表事实支持或
            QA 通过，来源权威性与时效仍未核验。
          </p>
          <p>研究范围：{report.research.record.draft.scope}</p>
          <details>
            <summary>原研究身份</summary>
            <p>
              原记录 ID：{report.research.record.id}；研究摘要：{report.research.record.draftDigest}
            </p>
          </details>
          <h4>引用范围与缺口 · {report.research.findings.length} 项</h4>
          {!report.research.findings.length && <p>未发现引用范围缺口，仍不代表主张已核验。</p>}
          <ul>
            {report.research.findings.map((finding, index) => (
              <li key={index}>
                {
                  {
                    unmapped_claim: '计划主张未映射研究',
                    omitted_conflict_partner: '未选用冲突另一方',
                    unselected_source_ref: '研究引用未选入计划',
                    source_unavailable: '原文尚不可用',
                  }[finding.code]
                }
                <p>
                  计划主张：{finding.claimId}
                  {finding.researchClaimId && `；原研究主张：${finding.researchClaimId}`}
                  {finding.relatedResearchClaimId &&
                    `；冲突另一方原 ID：${finding.relatedResearchClaimId}`}
                  {finding.sourceId && `；原研究来源：${finding.sourceId}`}
                </p>
              </li>
            ))}
          </ul>
          <details>
            <summary>完整研究结论与冲突双方</summary>
            <ol>
              {report.research.record.draft.facts.map((fact) => (
                <li key={fact.claimId}>
                  <p>
                    原主张 {fact.claimId} ·{' '}
                    {
                      {
                        fact: '事实主张',
                        quote: '引文',
                        calculation: '计算',
                        judgment: '判断',
                        assumption: '假设',
                      }[fact.type]
                    }
                    （待审查）：{fact.statement}
                  </p>
                  <p>
                    声明来源级别：
                    {
                      {
                        primary: '一手来源',
                        authoritative_secondary: '权威二手来源',
                        secondary: '二手来源',
                        unverified: '未核验',
                      }[fact.sourceTier]
                    }
                    ；声明可信度：{{ high: '高', medium: '中', low: '低' }[fact.confidence]}
                    ，不作为独立认证。
                  </p>
                  <p>原研究来源：{fact.sourceRefs.join('、') || '未提供'}</p>
                  {!!fact.conflictsWith.length && (
                    <p>
                      冲突：
                      {fact.conflictsWith
                        .map(
                          (id) =>
                            `${id}：${report.research!.record.draft.facts.find((other) => other.claimId === id)?.statement ?? '待读取'}`,
                        )
                        .join('；')}
                    </p>
                  )}
                  {fact.asOf && <p>数据时点：{fact.asOf}</p>}
                  <ProfessionalContext
                    context={fact.professionalContext}
                    claimId={fact.claimId}
                    claimType={fact.type}
                  />
                  {fact.jurisdiction && <p>适用范围：{fact.jurisdiction}</p>}
                  {fact.calculation && (
                    <p>
                      计算：{fact.calculation.formula}；输入：{fact.calculation.inputs.join('、')}
                      {fact.calculation.unit && `；单位：${fact.calculation.unit}`}
                      {fact.calculation.currency && `；币种：${fact.calculation.currency}`}
                    </p>
                  )}
                </li>
              ))}
            </ol>
          </details>
          <details>
            <summary>完整研究来源与原文缺口</summary>
            <ol>
              {report.research.record.draft.sources.map((source) => {
                const evidence = report.research!.record.sources?.find(
                  (item) => item.sourceId === source.id,
                )
                return (
                  <li key={source.id}>
                    <p>
                      原来源 {source.id}：
                      {/^https?:/.test(source.uri) ? (
                        <a href={source.uri} target="_blank" rel="noreferrer">
                          {source.title}
                        </a>
                      ) : (
                        source.title
                      )}
                    </p>
                    <blockquote>{source.excerpt || '未提供原文摘录'}</blockquote>
                    <p>
                      {evidence?.status === 'found'
                        ? '找到原文摘录，不证明主张成立'
                        : '原文仍有缺口，需核对完整研究资料'}
                    </p>
                    {source.locator && <p>原文位置：{source.locator}</p>}
                    {source.asOf && <p>资料数据时点：{source.asOf}</p>}
                  </li>
                )
              })}
            </ol>
          </details>
        </details>
      )}
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
          {report.plan.slides
            ?.find((slide) => slide.id === page.pageId)
            ?.claimIds.map((claimId) => {
              const claim = report.plan.claims.find((claim) => claim.id === claimId)
              if (!claim?.professionalContext && report.professionalWorkflow)
                return (
                  <details key={claimId} aria-label={`专业上下文缺失 ${claimId}`}>
                    <summary>主张 {claimId} · 缺少专业上下文</summary>
                    <p>{claim?.statement}</p>
                    <p>下一步：读取该主张来源原文，再补充专业限定；无法确认的字段保持未知。</p>
                    <p>说明或暂缓仍保留缺口，不代表专业审查或来源核验通过。</p>
                  </details>
                )
              return (
                <ProfessionalContext
                  key={claimId}
                  context={claim?.professionalContext}
                  claimId={claimId}
                  claimType={claim?.type}
                />
              )
            })}
          <SourceAssessmentHistory report={report} pageId={page.pageId} />
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
                    source_mismatch: '网页快照与计划网址不匹配',
                  }[report.sourceAudit.find((item) => item.sourceId === source.id)!.status]
                }
                {report.sourceAudit.find((item) => item.sourceId === source.id)?.locator &&
                  `；实际位置：${report.sourceAudit.find((item) => item.sourceId === source.id)!.locator}`}
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
