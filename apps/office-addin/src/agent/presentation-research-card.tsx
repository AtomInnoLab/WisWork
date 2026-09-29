import { useEffect, useState, useSyncExternalStore } from 'react'
import type { PresentationResearchController } from './presentation-research.js'
const kinds = {
  fact: '事实主张（待核验）',
  quote: '引文（待核验）',
  calculation: '计算（待核验）',
  judgment: '判断',
  assumption: '假设',
}
const tiers = {
  primary: '一手来源',
  authoritative_secondary: '权威二手来源',
  secondary: '二手来源',
  unverified: '未核验',
}
const confidence = { high: '高', medium: '中', low: '低' }
const gaps = {
  found: '找到原文摘录（不代表主张成立）',
  not_found: '未找到对应摘录',
  empty_excerpt: '摘录为空',
  not_ready: '原文尚未准备好',
  unsupported: '原文格式暂不支持',
  missing: '原文缺失',
  source_mismatch: '原文与来源不匹配',
}
export function PresentationResearchCard({
  controller,
  disabled,
}: {
  controller: PresentationResearchController
  disabled: boolean
}) {
  const snapshot = useSyncExternalStore(
    (listener) => controller.subscribe(listener),
    () => controller.snapshot(),
    () => controller.snapshot(),
  )
  const [project, setProject] = useState('')
  useEffect(() => {
    void controller.refresh()
    return () => controller.cancel()
  }, [controller])
  if (snapshot.available !== true) return null
  const busy = snapshot.phase !== 'idle'
  const blocked = disabled || busy
  const record = snapshot.record
  return (
    <section aria-label="资料研究账本" aria-busy={busy}>
      <details>
        <summary>资料研究账本</summary>
        <p>研究可先于制作计划整理；记录完成不代表主张支持；不代表来源权威性或时效通过。</p>
        <p>记录时间是研究整理操作时间；网页快照时间来自原资料，不表示本次重新抓取。</p>
        {busy && <p role="status">正在读取或导出研究记录…</p>}
        {snapshot.error && <p role="alert">{snapshot.error}</p>}
        {snapshot.notice && <p role="status">{snapshot.notice}</p>}
        <button type="button" disabled={blocked} onClick={() => void controller.refresh()}>
          刷新研究记录
        </button>
        {busy && (
          <button type="button" onClick={() => controller.cancel()}>
            停止等待研究记录
          </button>
        )}
        <details>
          <summary>选择已有研究项目</summary>
          <form
            onSubmit={(event) => {
              event.preventDefault()
              if (!blocked && /^[A-Za-z0-9_-]{1,128}$/.test(project))
                void controller.selectProject(project)
            }}
          >
            <label>
              研究项目 ID
              <input
                value={project}
                onChange={(event) => setProject(event.currentTarget.value)}
                maxLength={128}
                required
                pattern="[A-Za-z0-9_-]+"
              />
            </label>
            <button type="submit" disabled={blocked || !project}>
              读取所选研究项目
            </button>
          </form>
          {snapshot.projectId && <p>所选项目：{snapshot.projectId}</p>}
        </details>
        {!snapshot.projectId && <p>可先让 Agent 整理资料，或选择已有研究项目读取。</p>}
        {snapshot.summary && (
          <>
            <p>
              累计 {snapshot.summary.totalRecords} 条整理记录；显示最近{' '}
              {snapshot.summary.records.length} 条
              {snapshot.summary.totalRecords > snapshot.summary.records.length
                ? '，更早记录可请 Agent 按记录 ID 读取'
                : ''}
              。
            </p>
            {!snapshot.summary.records.length && (
              <p>当前项目暂无研究记录，不需要先建立制作计划。</p>
            )}
            <ol>
              {[...snapshot.summary.records].reverse().map((item) => (
                <li key={item.id}>
                  研究 #{item.sequence} ·{' '}
                  {item.state === 'completed'
                    ? '整理已保存（待核验）'
                    : item.state === 'failed'
                      ? '整理未完成，已有结论与冲突保留'
                      : '缺少结束回执（不表示后台仍在运行）'}
                  <p>
                    来源 {item.sourceCount} · 结论 {item.factCount} · 冲突 {item.conflictCount}
                  </p>
                  <p>
                    <time dateTime={item.startedAt}>{item.startedAt}</time>
                    {item.finishedAt && (
                      <>
                        {' '}
                        至 <time dateTime={item.finishedAt}>{item.finishedAt}</time>
                      </>
                    )}
                  </p>
                  {item.error && (
                    <p>
                      {item.error === 'aborted'
                        ? '整理等待已中断'
                        : item.error === 'source_unavailable'
                          ? '原文资料暂不可用'
                          : '研究记录状态异常'}
                      ；请明确提交新尝试，读取不会自动重放。
                    </p>
                  )}
                  <button
                    type="button"
                    disabled={blocked}
                    onClick={() => void controller.read(item.id)}
                  >
                    读取完整研究记录
                  </button>
                  <button
                    type="button"
                    disabled={blocked}
                    onClick={() => void controller.export(item.id)}
                  >
                    导出研究 JSON 与 Markdown
                  </button>
                  <details>
                    <summary>记录详情</summary>
                    <p>记录 ID：{item.id}</p>
                  </details>
                </li>
              ))}
            </ol>
          </>
        )}
        {record && (
          <details>
            <summary>研究结论、冲突与来源详情</summary>
            <h3>{record.draft.scope}</h3>
            <p>所有结论仍需审查；可信度与来源级别是声明，未经独立认证。</p>
            <ol>
              {record.draft.facts.map((fact) => (
                <li key={fact.claimId}>
                  <strong>
                    {kinds[fact.type]}：{fact.statement}
                  </strong>
                  <p>
                    声明来源级别：{tiers[fact.sourceTier]}；声明可信度：
                    {confidence[fact.confidence]}；待审查。
                  </p>
                  {!!fact.sourceRefs.length && (
                    <p>
                      来源：
                      {fact.sourceRefs
                        .map(
                          (id) =>
                            record.draft.sources.find((source) => source.id === id)?.title ?? id,
                        )
                        .join('、')}
                    </p>
                  )}
                  {!fact.sourceRefs.length && <p>来源缺口：未提供原文依据。</p>}
                  {!!fact.conflictsWith.length && (
                    <p>
                      冲突：
                      {fact.conflictsWith
                        .map(
                          (id) =>
                            record.draft.facts.find((other) => other.claimId === id)?.statement ??
                            id,
                        )
                        .join('；')}
                    </p>
                  )}
                  {fact.asOf && <p>结论数据时点：{fact.asOf}</p>}
                  {fact.jurisdiction && <p>适用范围：{fact.jurisdiction}</p>}
                  {fact.calculation && (
                    <p>
                      计算：{fact.calculation.formula}；输入：{fact.calculation.inputs.join('、')}
                      {fact.calculation.unit && `；单位：${fact.calculation.unit}`}
                      {fact.calculation.currency && `；币种：${fact.calculation.currency}`}
                    </p>
                  )}
                  {!!fact.slideRefs.length && (
                    <p>建议使用页面：{fact.slideRefs.join('、')}（不证明宿主页已存在）。</p>
                  )}
                </li>
              ))}
            </ol>
            <h4>来源与原文缺口</h4>
            <ol>
              {record.draft.sources.map((source) => {
                const evidence = record.sources?.find((item) => item.sourceId === source.id)
                return (
                  <li key={source.id}>
                    <p>
                      {/^https?:/.test(source.uri) ? (
                        <a href={source.uri} target="_blank" rel="noreferrer">
                          {source.title}
                        </a>
                      ) : (
                        source.title
                      )}
                    </p>
                    <blockquote>{source.excerpt || '未提供原文摘录'}</blockquote>
                    <p>{evidence ? gaps[evidence.status] : '尚无原文匹配回执'}</p>
                    {evidence?.provenance && (
                      <p>
                        原文身份：
                        {evidence.provenance === 'user_supplied'
                          ? '用户提供的材料'
                          : evidence.provenance === 'fetched_url_matched'
                            ? '网址与已保存快照匹配'
                            : '尚不可用'}
                        。
                      </p>
                    )}
                    {source.locator && <p>原文位置：{source.locator}</p>}
                    {source.asOf && <p>资料数据时点：{source.asOf}</p>}
                    {evidence?.retrievedAt && <p>原快照获取时间：{evidence.retrievedAt}</p>}
                  </li>
                )
              })}
            </ol>
          </details>
        )}
      </details>
    </section>
  )
}
