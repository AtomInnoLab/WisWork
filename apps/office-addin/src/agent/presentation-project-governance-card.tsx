import { useEffect, useState, useSyncExternalStore } from 'react'
import type { createPresentationProjectGovernanceController } from './presentation-project-governance'
import { downloadLocalFile } from './session-download.js'
type Controller = ReturnType<typeof createPresentationProjectGovernanceController>
export function PresentationProjectGovernanceCard({ controller }: { controller: Controller }) {
  const snapshot = useSyncExternalStore(
    controller.subscribe,
    controller.snapshot,
    controller.snapshot,
  )
  const [confirm, setConfirm] = useState(false),
    [content, setContent] = useState(''),
    [audit, setAudit] = useState('')
  useEffect(() => {
    setConfirm(false)
    setContent(snapshot.lifecycle?.policy.contentRetentionDays?.toString() ?? '')
    setAudit(snapshot.lifecycle?.policy.auditRetentionDays?.toString() ?? '')
  }, [
    snapshot.scope?.documentId,
    snapshot.scope?.projectId,
    snapshot.lifecycle?.revision,
    snapshot.lifecycle?.policy.contentRetentionDays,
    snapshot.lifecycle?.policy.auditRetentionDays,
  ])
  const busy = snapshot.phase === 'busy',
    enabled = snapshot.available && !busy
  const preview = snapshot.available ? snapshot.preview : undefined,
    lifecycle = snapshot.available ? snapshot.lifecycle : undefined
  return (
    <section aria-label="本机项目资料治理">
      <h3>本机项目资料治理</h3>
      <p>仅处理这台电脑上的项目资料，不修改当前 PowerPoint 内容。保留策略不会自动执行清理。</p>
      {!snapshot.available ? (
        <p>当前连接不支持本机项目治理。</p>
      ) : (
        <>
          <button
            disabled={busy}
            onClick={() => {
              setConfirm(false)
              void controller.refresh()
            }}
          >
            读取治理状态
          </button>
          <p>
            正文保留：{lifecycle?.policy.contentRetentionDays ?? '不设期限'}；匿名审计保留：
            {lifecycle?.policy.auditRetentionDays ?? '不设期限'}
          </p>
          {!lifecycle && preview && (
            <button
              disabled={!enabled}
              onClick={() => {
                void controller.initializePolicy()
              }}
            >
              启用本机项目治理
            </button>
          )}
          {lifecycle?.state === 'active' && (
            <fieldset disabled={!enabled}>
              <legend>保存保留策略</legend>
              <label>
                正文保留天数（留空表示不设期限）
                <input
                  type="number"
                  min="1"
                  max="36500"
                  value={content}
                  onChange={(e) => setContent(e.target.value)}
                />
              </label>
              <label>
                匿名审计保留天数（留空表示不设期限）
                <input
                  type="number"
                  min="1"
                  max="36500"
                  value={audit}
                  onChange={(e) => setAudit(e.target.value)}
                />
              </label>
              <button
                onClick={() => {
                  const days = (v: string) => (v === '' ? null : Number(v))
                  void controller.setPolicy({
                    contentRetentionDays: days(content),
                    auditRetentionDays: days(audit),
                  })
                }}
              >
                保存策略
              </button>
            </fieldset>
          )}
          <button
            disabled={!enabled || Boolean(snapshot.attempt)}
            onClick={() => {
              setConfirm(false)
              void controller.preview()
            }}
          >
            预览本机项目删除范围
          </button>
          {preview && (
            <>
              <p>
                候选资料 {preview.resources.filter((r) => r.disposition === 'candidate').length}{' '}
                项；保留资料 {preview.resources.filter((r) => r.disposition === 'retained').length}{' '}
                项。其中共享{' '}
                {
                  preview.resources.filter((r) =>
                    ['document_shared', 'global_shared'].includes(r.ownership),
                  ).length
                }{' '}
                项，归属未证明 {preview.resources.filter((r) => r.ownership === 'unproven').length}{' '}
                项。治理记录与匿名审计保留。
              </p>
              <label>
                <input
                  type="checkbox"
                  checked={confirm}
                  onChange={(e) => setConfirm(e.target.checked)}
                />
                我确认仅删除上述本机项目候选资料
              </label>
              <button
                disabled={!enabled || !confirm}
                onClick={() => {
                  setConfirm(false)
                  void controller.confirmDeletion()
                }}
              >
                确认删除本机项目资料
              </button>
            </>
          )}
          {snapshot.attempt && (
            <>
              <p>
                {snapshot.phase === 'deleted'
                  ? '原删除请求已完成。'
                  : snapshot.phase === 'partial'
                    ? '原删除请求尚有资料保留或未完成。'
                    : '无法确认原请求已受理，暂不继续删除。不会自动重发删除。'}
              </p>
              <button
                disabled={!enabled}
                onClick={() => {
                  void controller.checkAttempt()
                }}
              >
                核对原删除请求
              </button>
              {snapshot.lifecycle?.state === 'deleting' &&
                snapshot.lifecycle.deletion?.deletionId === snapshot.attempt.deletionId && (
                  <button
                    disabled={!enabled}
                    onClick={() => {
                      void controller.resumeDeletion()
                    }}
                  >
                    确认继续原删除请求
                  </button>
                )}
            </>
          )}
          {snapshot.available && snapshot.deletion?.projectContentRetained && (
            <p>本机项目正文仍保留。</p>
          )}
          {snapshot.deletion && (
            <p>
              已移除 {snapshot.deletion.counts.removed} 项，待处理{' '}
              {snapshot.deletion.counts.pending} 项，失败 {snapshot.deletion.counts.failed} 项，保留{' '}
              {snapshot.deletion.counts.retained} 项。
            </p>
          )}
          <button
            disabled={!enabled}
            onClick={() => {
              void controller.exportAudit()
            }}
          >
            读取匿名审计
          </button>
          {snapshot.audit && (
            <button
              disabled={!enabled}
              onClick={() => {
                const bytes = new TextEncoder().encode(JSON.stringify(snapshot.audit, null, 2))
                downloadLocalFile(bytes, '项目匿名审计.json', 'application/json')
              }}
            >
              下载匿名审计 JSON
            </button>
          )}
        </>
      )}
      {snapshot.error && snapshot.available && <p role="alert">{snapshot.error}</p>}
    </section>
  )
}
