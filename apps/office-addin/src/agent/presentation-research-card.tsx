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
  const [cleanupTarget, setCleanupTarget] = useState<{
    projectId: string
    documentId: string
    ledgerId: string
    draftDigest: string
  }>()
  const [abandonTarget, setAbandonTarget] = useState<{
    projectId: string
    documentId: string
    ledgerId: string
    draftDigest: string
  }>()
  const [forgetConfirm, setForgetConfirm] = useState(false)
  const [abandonRetryConfirm, setAbandonRetryConfirm] = useState(false)
  useEffect(() => {
    setAbandonTarget((target) =>
      target &&
      snapshot.recoveryAvailable &&
      target.projectId === snapshot.projectId &&
      target.documentId === snapshot.summary?.documentId &&
      snapshot.summary.records.some(
        (item) =>
          item.id === target.ledgerId &&
          item.draftDigest === target.draftDigest &&
          item.state === 'running',
      )
        ? target
        : undefined,
    )
    setAbandonRetryConfirm(false)
    setForgetConfirm(false)
  }, [snapshot.projectId, snapshot.summary, snapshot.recoveryAvailable, snapshot.abandonAttempt])
  const [retryConfirm, setRetryConfirm] = useState(false)
  useEffect(() => {
    setCleanupTarget((target) =>
      target &&
      snapshot.cleanupAvailable &&
      target.projectId === snapshot.projectId &&
      target.documentId === snapshot.summary?.documentId &&
      snapshot.summary.records.some(
        (item) =>
          item.id === target.ledgerId &&
          item.draftDigest === target.draftDigest &&
          item.state !== 'running',
      )
        ? target
        : undefined,
    )
    setRetryConfirm(false)
  }, [snapshot.projectId, snapshot.summary, snapshot.cleanupAvailable, snapshot.deleteAttempt])
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
        {busy && (
          <p role="status">
            {snapshot.phase === 'abandoning'
              ? '正在提交结束研究…'
              : snapshot.phase === 'checkingAbandon'
                ? '正在只读核对原研究记录…'
                : snapshot.phase === 'deleting'
                  ? '正在提交研究归档清理…'
                  : snapshot.phase === 'checkingDelete'
                    ? '正在只读核对删除回执…'
                    : '正在读取或导出研究记录…'}
          </p>
        )}
        {(snapshot.phase === 'deleting' || snapshot.phase === 'checkingDelete') && (
          <p>停止等待不承诺撤销已提交的清理；可随后读取删除回执，勿自动重发。</p>
        )}
        {(snapshot.phase === 'abandoning' || snapshot.phase === 'checkingAbandon') && (
          <p>停止等待只停止界面等待，不承诺撤销已提交的结束；随后只读核对原记录。</p>
        )}
        {snapshot.recoveryAvailable && snapshot.abandonAttempt && (
          <div aria-label="待核对研究结束尝试">
            <p>结束结果尚待核对；查询只读取原研究记录，不会结束或重跑。</p>
            <button
              type="button"
              disabled={blocked}
              onClick={() => void controller.checkAbandonStatus()}
            >
              读取结束状态
            </button>
            <button type="button" disabled={blocked} onClick={() => setAbandonRetryConfirm(true)}>
              重试同一次结束
            </button>
            {abandonRetryConfirm && (
              <div>
                <p>
                  明确重试使用原研究身份、摘要与版本；不会重跑研究。草稿和附件保留，活跃研究正常完成时不会覆盖其结果。
                </p>
                <button
                  type="button"
                  disabled={blocked}
                  onClick={() => {
                    setAbandonRetryConfirm(false)
                    void controller.retryAbandon()
                  }}
                >
                  确认重试同一次结束
                </button>
                <button type="button" onClick={() => setAbandonRetryConfirm(false)}>
                  取消重试结束
                </button>
              </div>
            )}
            <button type="button" disabled={blocked} onClick={() => setForgetConfirm(true)}>
              仅忘记本机恢复身份
            </button>
            {forgetConfirm && (
              <div>
                <p>
                  仅清除本机恢复身份，不修改 PC
                  研究记录，不结束、不删除、不重跑；此界面将失去原尝试核对身份，请先读取原记录。
                </p>
                <button
                  type="button"
                  disabled={blocked}
                  onClick={() => {
                    setForgetConfirm(false)
                    void controller.forgetAbandon()
                  }}
                >
                  确认忘记恢复身份
                </button>
                <button type="button" onClick={() => setForgetConfirm(false)}>
                  取消忘记
                </button>
              </div>
            )}
            <details>
              <summary>结束恢复身份</summary>
              <p>
                原研究 #{snapshot.abandonAttempt.sequence} · {snapshot.abandonAttempt.ledgerId}
                ；原版本 {snapshot.abandonAttempt.expectedRevision}
              </p>
            </details>
          </div>
        )}
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
        {snapshot.cleanupAvailable && snapshot.deleteAttempt && (
          <div aria-label="待核对清理尝试">
            <p>此清理尝试尚待核对。读取回执只查本机状态，不会再次删除。</p>
            <button
              type="button"
              disabled={blocked}
              onClick={() => void controller.checkDeleteStatus()}
            >
              读取删除回执
            </button>
            <button type="button" disabled={blocked} onClick={() => setRetryConfirm(true)}>
              重试同一次清理
            </button>
            {retryConfirm && (
              <div>
                <p>
                  明确重试仍使用原记录、删除身份与版本；只清理本机研究归档，原附件、PowerPoint
                  文稿、交付包与导出副本仍保留。
                </p>
                <button
                  type="button"
                  disabled={blocked}
                  onClick={() => {
                    setRetryConfirm(false)
                    void controller.retryDelete()
                  }}
                >
                  确认重试同一次清理
                </button>
                <button type="button" onClick={() => setRetryConfirm(false)}>
                  取消重试
                </button>
              </div>
            )}
            <details>
              <summary>清理恢复身份</summary>
              <p>
                原研究 #{snapshot.deleteAttempt.sequence} · {snapshot.deleteAttempt.ledgerId}
                ；删除身份：{snapshot.deleteAttempt.deleteId}
              </p>
            </details>
          </div>
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
              现存 {snapshot.summary.totalRecords} 条整理记录；显示最近{' '}
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
                  {snapshot.recoveryAvailable && item.state === 'running' && (
                    <>
                      <button
                        type="button"
                        disabled={blocked || !!snapshot.abandonAttempt}
                        onClick={() =>
                          setAbandonTarget({
                            projectId: snapshot.projectId!,
                            documentId: snapshot.summary!.documentId,
                            ledgerId: item.id,
                            draftDigest: item.draftDigest,
                          })
                        }
                      >
                        结束未完成研究
                      </button>
                      {abandonTarget?.ledgerId === item.id && (
                        <div aria-label="确认结束研究">
                          <p>
                            仅结束这条未完成研究，保留草稿、原附件、PowerPoint
                            文稿与导出副本，不会重跑。缺少结束回执不能证明后台已中断；活跃研究正常完成时不会覆盖其结果。
                          </p>
                          <button
                            type="button"
                            disabled={blocked || !!snapshot.abandonAttempt}
                            onClick={() => {
                              const target = abandonTarget
                              setAbandonTarget(undefined)
                              if (
                                target.projectId === snapshot.projectId &&
                                target.documentId === snapshot.summary?.documentId
                              )
                                void controller.abandonRecord(target.ledgerId, target.draftDigest)
                            }}
                          >
                            确认结束此研究
                          </button>
                          <button type="button" onClick={() => setAbandonTarget(undefined)}>
                            取消结束
                          </button>
                        </div>
                      )}
                    </>
                  )}
                  {snapshot.cleanupAvailable &&
                    (item.state === 'running' ? (
                      <p>未收到结束回执，不能清理此归档。</p>
                    ) : (
                      <>
                        <button
                          type="button"
                          disabled={blocked || !!snapshot.deleteAttempt}
                          onClick={() =>
                            setCleanupTarget({
                              projectId: snapshot.projectId!,
                              documentId: snapshot.summary!.documentId,
                              ledgerId: item.id,
                              draftDigest: item.draftDigest,
                            })
                          }
                        >
                          清理本机研究归档
                        </button>
                        {cleanupTarget?.ledgerId === item.id && (
                          <div aria-label="确认研究归档清理">
                            <p>
                              只删除本机这条已结束且未被引用的研究归档；若仍被计划或冻结任务引用，本机会拒绝清理。原附件、PowerPoint
                              文稿、交付包与导出副本仍保留。
                            </p>
                            <button
                              type="button"
                              disabled={blocked || !!snapshot.deleteAttempt}
                              onClick={() => {
                                const target = cleanupTarget
                                setCleanupTarget(undefined)
                                if (
                                  target.projectId === snapshot.projectId &&
                                  target.documentId === snapshot.summary?.documentId
                                )
                                  void controller.deleteRecord(target.ledgerId, target.draftDigest)
                              }}
                            >
                              确认清理此归档
                            </button>
                            <button type="button" onClick={() => setCleanupTarget(undefined)}>
                              取消清理
                            </button>
                          </div>
                        )}
                      </>
                    ))}
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
