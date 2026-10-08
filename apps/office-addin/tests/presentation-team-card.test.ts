// @vitest-environment jsdom
import React, { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, expect, it, vi } from 'vitest'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { PresentationTeamCard } from '../src/agent/presentation-team-card.js'
import type {
  PresentationTeamController,
  PresentationTeamSnapshot,
} from '../src/agent/presentation-team-controller.js'
const roots: ReturnType<typeof createRoot>[] = []
afterEach(async () => {
  for (const root of roots.splice(0)) await act(async () => root.unmount())
})
function snapshot(role: 'owner' | 'reviewer' | 'viewer' = 'owner'): PresentationTeamSnapshot {
  const owner = 'a'.repeat(64),
    actor = role === 'owner' ? owner : 'b'.repeat(64),
    plan = benchmarkPlan()
  return {
    available: true,
    phase: 'idle',
    identity: { version: 1, actorSubject: actor, pcSubject: owner },
    role,
    team: {
      version: 1,
      teamId: 'team_' + 'c'.repeat(64),
      documentId: 'doc',
      projectId: plan.projectId,
      ownerSubject: owner,
      revision: 3,
      createdAt: '2026-09-29T00:00:00.000Z',
      updatedAt: '2026-09-29T00:00:00.000Z',
      publishedPlan: { revision: 2, plan },
      members: [{ subject: 'b'.repeat(64), role: 'reviewer' }],
      comments: [
        {
          id: 'old-comment',
          targetKind: 'slide',
          targetId: plan.slides[0]!.id,
          text: 'Historical review',
          authorSubject: actor,
          planRevision: 1,
          state: 'open',
          createdAt: '2026-09-29T00:00:00.000Z',
          updatedAt: '2026-09-29T00:00:00.000Z',
        },
      ],
    },
  }
}
async function mount(
  initial: PresentationTeamSnapshot,
  disabled = false,
  account?: import('../src/agent/team-connection.js').OfficeTeamConnection,
) {
  let state = initial
  const listeners = new Set<() => void>()
  const controller: PresentationTeamController = {
    snapshot: () => state,
    subscribe: (l) => {
      listeners.add(l)
      return () => listeners.delete(l)
    },
    refresh: vi.fn(async () => {}),
    create: vi.fn(async () => 'proposal'),
    publish: vi.fn(async () => 'proposal'),
    setMember: vi.fn(async () => 'proposal'),
    revokeMember: vi.fn(async () => 'proposal'),
    addComment: vi.fn(async () => 'proposal'),
    resolveComment: vi.fn(async () => 'proposal'),
    clear: vi.fn(),
  }
  const node = document.createElement('div')
  const root = createRoot(node)
  roots.push(root)
  await act(async () =>
    root.render(React.createElement(PresentationTeamCard, { controller, disabled, account })),
  )
  return {
    node,
    controller,
    update: async (next: PresentationTeamSnapshot) => {
      state = next
      await act(async () => listeners.forEach((l) => l()))
    },
  }
}
async function input(node: HTMLElement, label: string, value: string) {
  const field = Array.from(node.querySelectorAll('input,textarea,select')).find(
    (e) => e.getAttribute('aria-label') === label,
  ) as HTMLInputElement
  await act(async () => {
    const setter = Object.getOwnPropertyDescriptor(
      field instanceof HTMLTextAreaElement
        ? HTMLTextAreaElement.prototype
        : field instanceof HTMLSelectElement
          ? HTMLSelectElement.prototype
          : HTMLInputElement.prototype,
      'value',
    )!.set!
    setter.call(field, value)
    field.dispatchEvent(new Event('input', { bubbles: true }))
    field.dispatchEvent(new Event('change', { bubbles: true }))
  })
}
async function click(node: HTMLElement, text: string) {
  const button = Array.from(node.querySelectorAll('button')).find((b) => b.textContent === text)!
  await act(async () => button.click())
}
it('shows unavailable authentication without requesting tokens or inventing login', async () => {
  const f = await mount({ available: false, phase: 'idle' })
  expect(f.node.textContent).toContain('团队协作暂不可用')
  expect(f.node.textContent).toContain('请先连接团队账号')
  expect(f.controller.refresh).not.toHaveBeenCalled()
})
it('owner reads explicit team and creates only a visible sharing proposal', async () => {
  const f = await mount({
    available: true,
    phase: 'idle',
    identity: { version: 1, actorSubject: 'a'.repeat(64), pcSubject: 'a'.repeat(64) },
  })
  await input(f.node, '团队 ID', 'team_' + 'c'.repeat(64))
  await click(f.node, '读取团队')
  expect(f.controller.refresh).toHaveBeenCalledWith('team_' + 'c'.repeat(64))
  await input(f.node, '共享项目 ID', 'project')
  await input(f.node, '已保存计划修订', '7')
  await click(f.node, '创建共享提案')
  expect(f.controller.create).toHaveBeenCalledWith('project', 7)
  expect(f.node.textContent).toContain('完整已保存计划与来源摘录')
})
it('reviewer comments on published targets and history stays on its original revision', async () => {
  const s = snapshot('reviewer'),
    f = await mount(s)
  expect(f.node.textContent).toContain('历史意见（计划修订 1）')
  expect(f.node.textContent).toContain('已发布计划修订：2')
  expect(f.node.textContent).toContain(s.team!.publishedPlan.plan.sources[0]!.id)
  await input(f.node, '评论内容', 'Review source')
  await click(f.node, '创建评论提案')
  expect(f.controller.addComment).toHaveBeenCalledWith(
    'slide',
    s.team!.publishedPlan.plan.slides[0]!.id,
    'Review source',
  )
  await click(f.node, '创建解决提案')
  expect(f.controller.resolveComment).toHaveBeenCalledWith('old-comment')
  expect(f.node.textContent).not.toContain('成员权限提案')
})
it('viewer is read only and revoked access clears published content while retaining a safe error', async () => {
  const f = await mount(snapshot('viewer'))
  expect(f.node.textContent).toContain('只读成员')
  expect(f.node.textContent).not.toContain('创建评论提案')
  expect(f.node.textContent).not.toContain('创建解决提案')
  await f.update({ available: true, phase: 'idle', error: 'access_denied' })
  expect(f.node.textContent).not.toContain('Historical review')
  expect(f.node.textContent).not.toContain('已发布计划修订：2')
  expect(f.node.textContent).toContain('没有团队访问权限')
})
it('busy and external disabled states prevent all mutations', async () => {
  const f = await mount({ ...snapshot(), phase: 'awaiting_confirmation' })
  expect(Array.from(f.node.querySelectorAll('button')).every((b) => b.disabled)).toBe(true)
  await click(f.node, '创建解决提案')
  expect(f.controller.resolveComment).not.toHaveBeenCalled()
  const g = await mount(snapshot(), true)
  expect(Array.from(g.node.querySelectorAll('button')).every((b) => b.disabled)).toBe(true)
})
it('owner publishes an exact saved revision and proposes member grants and revocations', async () => {
  const f = await mount(snapshot())
  await input(f.node, '已保存计划修订', '4')
  await click(f.node, '创建新版本发布提案')
  expect(f.controller.publish).toHaveBeenCalledWith(4)
  await input(f.node, '成员账号主体', 'd'.repeat(64))
  await input(f.node, '授予成员权限', 'viewer')
  await click(f.node, '创建成员权限提案')
  expect(f.controller.setMember).toHaveBeenCalledWith('d'.repeat(64), 'viewer')
  await click(f.node, '创建撤销权限提案')
  expect(f.controller.revokeMember).toHaveBeenCalledWith('b'.repeat(64))
})
it('a different published team resets a draft comment and unknown error codes stay private', async () => {
  const f = await mount(snapshot('reviewer'))
  await input(f.node, '评论内容', 'Do not move this to another project')
  const next = snapshot('reviewer')
  next.team = { ...next.team!, teamId: 'team_' + 'e'.repeat(64) }
  next.error = 'Bearer private-token'
  await f.update(next)
  expect(f.node.textContent).not.toContain('Bearer private-token')
  expect((f.node.querySelector('textarea') as HTMLTextAreaElement).value).toBe('')
  await click(f.node, '创建评论提案')
  expect(f.controller.addComment).not.toHaveBeenCalled()
})
it('retains an explicitly selected public team ID after content clears and refreshes with that exact ID', async () => {
  const s = snapshot(),
    f = await mount(s)
  expect((f.node.querySelector('[aria-label="团队 ID"]') as HTMLInputElement).value).toBe(
    s.team!.teamId,
  )
  await f.update({ available: true, phase: 'idle', identity: s.identity, error: 'cancelled' })
  expect((f.node.querySelector('[aria-label="团队 ID"]') as HTMLInputElement).value).toBe(
    s.team!.teamId,
  )
  expect(f.node.textContent).not.toContain('Historical review')
  expect(f.node.textContent).not.toContain('已发布计划修订：2')
  await click(f.node, '刷新身份与权限')
  expect(f.controller.refresh).toHaveBeenCalledWith(s.team!.teamId)
})
it('hides stale identity and all shared content whenever authenticated availability becomes false', async () => {
  const stale = { ...snapshot(), available: false },
    f = await mount(stale)
  expect(f.node.textContent).toContain('团队协作暂不可用')
  for (const text of [
    '当前已验证账号',
    '项目所有者',
    '合成基准',
    'Historical review',
    '已发布计划修订',
    '已验证作者',
    '成员账号 ID',
    '创建共享提案',
    '创建评论提案',
  ])
    expect(f.node.textContent).not.toContain(text)
  expect(f.node.textContent).not.toContain(stale.identity!.actorSubject)
  expect(f.node.textContent).not.toContain(stale.team!.publishedPlan.plan.sources[0]!.uri)
  expect(f.node.querySelector('textarea')).toBeNull()
})
it('optional authenticated account UI signs in, shows pairing and permits signout without treating display ID as author', async () => {
  const listeners = new Set<() => void>()
  let state: import('../src/agent/team-connection.js').OfficeTeamConnectionSnapshot = {
    phase: 'signed_out',
  }
  const account: import('../src/agent/team-connection.js').OfficeTeamConnection = {
    snapshot: () => state,
    subscribe: (l) => {
      listeners.add(l)
      return () => listeners.delete(l)
    },
    signIn: vi.fn(async () => {}),
    connect: vi.fn(async () => {}),
    signOut: vi.fn(async () => {}),
    request: vi.fn(),
    dispose: vi.fn(),
    available: () => state.phase === 'ready',
  }
  const f = await mount({ available: false, phase: 'idle' }, false, account)
  await click(f.node, '登录团队账号')
  expect(account.signIn).toHaveBeenCalledOnce()
  state = {
    phase: 'pairing',
    account: { loggedIn: true, email: 'reader@example.com', userId: 'display-not-author' },
    relayStatus: 'pending',
    verificationCode: '123456',
  }
  await act(async () => listeners.forEach((l) => l()))
  expect(f.node.textContent).toContain('123456')
  expect(f.node.textContent).toContain('reader@example.com')
  expect(f.node.textContent).not.toContain('当前已验证账号 ID：display-not-author')
  await click(f.node, '退出团队账号')
  expect(account.signOut).toHaveBeenCalledOnce()
})
it('without a configured account provider there is no invented login button', async () => {
  const f = await mount({ available: false, phase: 'idle' })
  expect(f.node.textContent).not.toContain('登录团队账号')
  expect(f.node.textContent).not.toContain('退出团队账号')
})
it('an authenticated ready connection permits explicit permission refresh before the team snapshot is available', async () => {
  const accountState: import('../src/agent/team-connection.js').OfficeTeamConnectionSnapshot = {
    phase: 'ready',
    account: { loggedIn: true },
  }
  const account: import('../src/agent/team-connection.js').OfficeTeamConnection = {
    snapshot: () => accountState,
    subscribe: () => () => {},
    signIn: vi.fn(),
    connect: vi.fn(),
    signOut: vi.fn(),
    request: vi.fn(),
    dispose: vi.fn(),
    available: () => true,
  }
  const f = await mount({ available: false, phase: 'idle' }, false, account)
  await click(f.node, '刷新身份与权限')
  expect(f.controller.refresh).toHaveBeenCalledWith(undefined)
})
it('a signed-out account immediately hides a stale authorized team snapshot', async () => {
  const accountState: import('../src/agent/team-connection.js').OfficeTeamConnectionSnapshot = {
    phase: 'signed_out',
  }
  const account: import('../src/agent/team-connection.js').OfficeTeamConnection = {
    snapshot: () => accountState,
    subscribe: () => () => {},
    signIn: vi.fn(),
    connect: vi.fn(),
    signOut: vi.fn(),
    request: vi.fn(),
    dispose: vi.fn(),
    available: () => false,
  }
  const f = await mount(snapshot(), false, account)
  expect(f.node.textContent).not.toContain('Historical review')
  expect(f.node.textContent).not.toContain('当前已验证账号 ID')
  expect(f.node.textContent).not.toContain('创建新版本发布提案')
})
