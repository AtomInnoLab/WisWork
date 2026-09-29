import { useEffect, useState, useSyncExternalStore } from 'react'
import type { OfficeTeamConnection, OfficeTeamConnectionSnapshot } from './team-connection.js'
import type { PresentationTeamController } from './presentation-team-controller.js'
const signedOutSnapshot: OfficeTeamConnectionSnapshot = Object.freeze({ phase: 'signed_out' })
const roles = { owner: '项目所有者', reviewer: '审阅成员', viewer: '只读成员' }
const kinds = { slide: '页面', claim: '结论', source: '来源' }
function errorText(code: string) {
  const messages: Record<string, string> = {
    access_denied: '没有团队访问权限，内容已停止显示；请向所有者核对权限。',
    not_found: '团队记录不存在，请核对团队 ID。',
    revision_conflict: '团队记录已变化，请刷新后重新创建提案。',
    presentation_team_account_changed: '团队账号已变化，请刷新身份后重新操作。',
    presentation_document_changed: '当前文档已变化，请刷新后重新操作。',
    presentation_team_unavailable: '团队连接不可用，请检查真实团队连接后刷新。',
    presentation_response_invalid: '团队回执无法核对，请刷新读取真实记录。',
    presentation_team_source_changed: '已保存计划在确认前发生变化，请重新读取后发起共享。',
    invalid_tool_input: '请核对输入的项目、修订、成员或评论目标。',
    invalid_request: '输入暂不符合团队操作要求，请核对后重试。',
    invalid_state: '团队保存记录异常，暂不能继续操作。',
    quota_exceeded: '团队保存容量已满，未自动清理或重试。',
    cancelled: '操作等待已停止；已提交结果请刷新核对。',
  }
  return messages[code] ?? '团队操作未完成，请刷新核对权限和真实记录。'
}
export function PresentationTeamCard({
  controller,
  disabled,
  account,
}: {
  controller: PresentationTeamController
  disabled: boolean
  account?: OfficeTeamConnection
}) {
  const snapshot = useSyncExternalStore(
    (listener) => controller.subscribe(listener),
    () => controller.snapshot(),
    () => controller.snapshot(),
  )
  const accountState = useSyncExternalStore(
    (listener) => account?.subscribe(listener) ?? (() => {}),
    () => account?.snapshot() ?? signedOutSnapshot,
    () => account?.snapshot() ?? signedOutSnapshot,
  )
  const [teamId, setTeamId] = useState('')
  const [projectId, setProjectId] = useState('')
  const [revision, setRevision] = useState('')
  const [member, setMember] = useState('')
  const [memberRole, setMemberRole] = useState<'reviewer' | 'viewer'>('reviewer')
  const [kind, setKind] = useState<'slide' | 'claim' | 'source'>('slide')
  const [target, setTarget] = useState('')
  const [text, setText] = useState('')
  const [localError, setLocalError] = useState(false)
  const sharedAvailable =
    snapshot.available && (!account || (accountState.phase === 'ready' && account.available()))
  const team = sharedAvailable ? snapshot.team : undefined
  useEffect(() => {
    if (snapshot.team?.teamId && /^team_[a-f0-9]{64}$/.test(snapshot.team.teamId))
      setTeamId(snapshot.team.teamId)
  }, [snapshot.team?.teamId])
  useEffect(() => {
    setTarget('')
    setText('')
    setMember('')
    setRevision('')
  }, [team?.teamId, team?.publishedPlan.revision])
  const busy = snapshot.phase !== 'idle'
  const blocked = disabled || busy || !sharedAvailable
  const readBlocked = disabled || busy || (!snapshot.available && !account?.available())
  const accountBusy = accountState.phase === 'signing_in' || accountState.phase === 'pairing'
  const accountRun = (action: () => Promise<void>) => {
    void action().catch(() => setLocalError(true))
  }
  const owner =
    sharedAvailable &&
    (snapshot.role === 'owner' ||
      (!team &&
        snapshot.identity?.actorSubject === snapshot.identity?.pcSubject &&
        Boolean(snapshot.identity)))
  const canComment = Boolean(team && (snapshot.role === 'owner' || snapshot.role === 'reviewer'))
  const targets = team
    ? kind === 'slide'
      ? team.publishedPlan.plan.slides
      : kind === 'claim'
        ? team.publishedPlan.plan.claims
        : team.publishedPlan.plan.sources
    : []
  const targetId = targets.some((item) => item.id === target) ? target : (targets[0]?.id ?? '')
  const validRevision = /^[1-9][0-9]*$/.test(revision) && Number.isSafeInteger(Number(revision))
  const run = (action: () => Promise<unknown>) => {
    if (!blocked) void action().catch(() => setLocalError(true))
  }
  return (
    <section aria-label="团队审阅工作台" aria-busy={busy}>
      <details>
        <summary>团队审阅工作台</summary>
        {account && (
          <div aria-label="团队账号连接">
            {accountState.account?.loggedIn && (
              <p>
                登录账号：{accountState.account.email ?? accountState.account.userId ?? '已登录'}
                。团队权限以验证后的账号身份为准。
              </p>
            )}
            {accountState.phase === 'signing_in' && <p role="status">正在登录团队账号…</p>}
            {accountState.phase === 'pairing' && (
              <p role="status">
                等待配对 PC 批准团队连接。
                {accountState.verificationCode && <>验证码：{accountState.verificationCode}</>}
              </p>
            )}
            {accountState.phase === 'ready' && (
              <p role="status">团队连接已建立；请刷新身份与权限。</p>
            )}
            {accountState.error && (
              <p role="alert">
                团队登录或连接未完成，请核对账号或重新连接。
                {accountState.error === 'team_auth_dialog_cancelled' ||
                accountState.error === 'team_connection_cancelled'
                  ? '登录已取消。'
                  : accountState.error === 'team_auth_dialog_timeout'
                    ? '登录等待已超时。'
                    : ''}
              </p>
            )}
            {!accountState.account?.loggedIn && (
              <button
                type="button"
                disabled={disabled || accountBusy}
                onClick={() => accountRun(() => account.signIn())}
              >
                登录团队账号
              </button>
            )}
            {accountState.account?.loggedIn && (
              <button
                type="button"
                disabled={disabled || accountBusy || accountState.phase === 'ready'}
                onClick={() => accountRun(() => account.connect())}
              >
                连接团队
              </button>
            )}
            {(accountState.account?.loggedIn || accountBusy) && (
              <button type="button" onClick={() => accountRun(() => account.signOut())}>
                退出团队账号
              </button>
            )}
          </div>
        )}
        <p>显示已明确共享的计划。私人修改不会自动发布；评论不表示内容已核验。</p>
        {!sharedAvailable && <p>团队协作暂不可用。请先连接团队账号。</p>}
        {sharedAvailable && snapshot.identity && (
          <div>
            <p>当前已验证账号 ID：{snapshot.identity.actorSubject}</p>
            <p>配对 PC 账号 ID：{snapshot.identity.pcSubject}</p>
            <p>
              当前角色：
              {snapshot.role
                ? roles[snapshot.role]
                : owner
                  ? '项目所有者（尚未读取团队）'
                  : '尚未读取团队权限'}
            </p>
          </div>
        )}
        {snapshot.error && <p role="alert">{errorText(snapshot.error)}</p>}
        {localError && <p role="alert">操作未完成，请刷新核对真实团队记录。</p>}
        {sharedAvailable && snapshot.notice && <p role="status">{snapshot.notice}</p>}
        {sharedAvailable && snapshot.phase === 'awaiting_confirmation' && (
          <p role="status">提案等待确认，请在确认面板审阅并决定。</p>
        )}
        <label>
          团队 ID{' '}
          <input
            aria-label="团队 ID"
            value={teamId}
            onChange={(e) => setTeamId(e.target.value)}
            disabled={readBlocked}
          />
        </label>
        <button
          type="button"
          disabled={readBlocked || !/^team_[a-f0-9]{64}$/.test(teamId)}
          onClick={() => {
            if (!readBlocked) void controller.refresh(teamId).catch(() => setLocalError(true))
          }}
        >
          读取团队
        </button>
        <button
          type="button"
          disabled={readBlocked}
          onClick={() => {
            if (!readBlocked)
              void controller
                .refresh(/^team_[a-f0-9]{64}$/.test(teamId) ? teamId : undefined)
                .catch(() => setLocalError(true))
          }}
        >
          刷新身份与权限
        </button>
        {owner && (
          <div>
            {!team && (
              <label>
                共享项目 ID{' '}
                <input
                  aria-label="共享项目 ID"
                  disabled={blocked}
                  value={projectId}
                  onChange={(e) => setProjectId(e.target.value)}
                />
              </label>
            )}
            <label>
              已保存计划修订{' '}
              <input
                aria-label="已保存计划修订"
                disabled={blocked}
                value={revision}
                onChange={(e) => setRevision(e.target.value)}
                inputMode="numeric"
              />
            </label>
            <p>
              共享提案会展示来源摘要；确认会共享完整已保存计划与来源摘录。创建、发布均需要单独确认，不会更改当前计划或
              PowerPoint。
            </p>
            <button
              type="button"
              disabled={
                blocked || !validRevision || (!team && !/^[A-Za-z0-9_-]{1,80}$/.test(projectId))
              }
              onClick={() =>
                run(() =>
                  team
                    ? controller.publish(Number(revision))
                    : controller.create(projectId, Number(revision)),
                )
              }
            >
              {team ? '创建新版本发布提案' : '创建共享提案'}
            </button>
          </div>
        )}
        {team && (
          <div>
            <p>团队 ID：{team.teamId}</p>
            <p>已发布计划：{team.publishedPlan.plan.title}</p>
            <p>已发布计划修订：{team.publishedPlan.revision}</p>
            <p>团队账本修订：{team.revision}</p>
            <details>
              <summary>
                已发布来源（完整列表，共 {team.publishedPlan.plan.sources.length} 条）
              </summary>
              {team.publishedPlan.plan.sources.map((source) => (
                <div key={source.id}>
                  <p>
                    来源 ID：{source.id}；标题：{source.title}
                  </p>
                  <p>来源地址：{source.uri}</p>
                  {source.locator && <p>位置：{source.locator}</p>}
                </div>
              ))}
            </details>
            <details>
              <summary>团队成员</summary>
              {team.members.length ? (
                team.members.map((m) => (
                  <div key={m.subject}>
                    <p>
                      成员账号 ID：{m.subject}；角色：{roles[m.role]}
                    </p>
                    {owner && (
                      <button
                        type="button"
                        disabled={blocked}
                        onClick={() => run(() => controller.revokeMember(m.subject))}
                      >
                        创建撤销权限提案
                      </button>
                    )}
                  </div>
                ))
              ) : (
                <p>尚未授权成员。</p>
              )}
            </details>
            {owner && (
              <div aria-label="成员权限提案">
                <label>
                  成员账号主体{' '}
                  <input
                    aria-label="成员账号主体"
                    disabled={blocked}
                    value={member}
                    onChange={(e) => setMember(e.target.value)}
                  />
                </label>
                <label>
                  授予成员权限{' '}
                  <select
                    aria-label="授予成员权限"
                    disabled={blocked}
                    value={memberRole}
                    onChange={(e) => setMemberRole(e.target.value as 'reviewer' | 'viewer')}
                  >
                    <option value="reviewer">审阅成员</option>
                    <option value="viewer">只读成员</option>
                  </select>
                </label>
                <button
                  type="button"
                  disabled={
                    blocked || !/^[a-f0-9]{64}$/.test(member) || member === team.ownerSubject
                  }
                  onClick={() => run(() => controller.setMember(member, memberRole))}
                >
                  创建成员权限提案
                </button>
              </div>
            )}
            {canComment && (
              <div>
                <label>
                  评论目标类型{' '}
                  <select
                    aria-label="评论目标类型"
                    disabled={blocked}
                    value={kind}
                    onChange={(e) => {
                      setKind(e.target.value as typeof kind)
                      setTarget('')
                    }}
                  >
                    {Object.entries(kinds).map(([value, label]) => (
                      <option key={value} value={value}>
                        {label}
                      </option>
                    ))}
                  </select>
                </label>
                <label>
                  已发布评论目标{' '}
                  <select
                    aria-label="已发布评论目标"
                    disabled={blocked}
                    value={targetId}
                    onChange={(e) => setTarget(e.target.value)}
                  >
                    {targets.map((item) => (
                      <option key={item.id} value={item.id}>
                        {item.id}
                        {'title' in item ? `：${item.title}` : ''}
                      </option>
                    ))}
                  </select>
                </label>
                <label>
                  评论内容{' '}
                  <textarea
                    aria-label="评论内容"
                    disabled={blocked}
                    maxLength={2000}
                    value={text}
                    onChange={(e) => setText(e.target.value)}
                  />
                </label>
                <button
                  type="button"
                  disabled={blocked || !targetId || !text.trim()}
                  onClick={() => run(() => controller.addComment(kind, targetId, text))}
                >
                  创建评论提案
                </button>
              </div>
            )}
            <details>
              <summary>审阅评论（{team.comments.length} 条）</summary>
              {team.comments.length ? (
                team.comments.map((comment) => (
                  <div key={comment.id}>
                    <p>
                      {comment.planRevision !== team.publishedPlan.revision
                        ? `历史意见（计划修订 ${comment.planRevision}）`
                        : `当前意见（计划修订 ${comment.planRevision}）`}
                    </p>
                    <p>
                      {kinds[comment.targetKind]} ID：{comment.targetId}；评论 ID：{comment.id}
                    </p>
                    <p>{comment.text}</p>
                    <p>
                      作者账号 ID：{comment.authorSubject}；状态：
                      {comment.state === 'open' ? '待处理' : '已解决（不代表内容获认证）'}
                    </p>
                    {canComment &&
                      comment.state === 'open' &&
                      (owner || comment.authorSubject === snapshot.identity?.actorSubject) && (
                        <button
                          type="button"
                          disabled={blocked}
                          onClick={() => run(() => controller.resolveComment(comment.id))}
                        >
                          创建解决提案
                        </button>
                      )}
                  </div>
                ))
              ) : (
                <p>尚无审阅评论。</p>
              )}
            </details>
          </div>
        )}
      </details>
    </section>
  )
}
