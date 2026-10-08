// @vitest-environment jsdom
import React, { act } from 'react'
import { createRoot } from 'react-dom/client'
import { expect, it, vi } from 'vitest'
import { AgentWorkspace, type OfficeWorkspaceUi } from '../src/App.js'
import type { OfficeAgentSession, OfficeAgentSnapshot } from '../src/agent/use-office-agent.js'
import type { PresentationProjectController } from '../src/skills/powerpoint/presentation-project.js'

it('connects the completion action to Stop while applying, preserving the conversation and project', async () => {
  const snapshot: OfficeAgentSnapshot = {
    assistantText: '准备完成',
    activity: '',
    busy: false,
    applying: true,
    status: 'working',
    retryable: false,
    timeline: [{ id: 'user-1', kind: 'user', text: '修改当前页' }],
  }
  const session = {
    snapshot: () => snapshot,
    subscribe: () => () => {},
    stop: vi.fn(),
    newTask: vi.fn(),
  } as unknown as OfficeAgentSession
  const projectSnapshot = { phase: 'loading' as const }
  const project = {
    snapshot: () => projectSnapshot,
    subscribe: () => () => {},
    refresh: vi.fn(async () => {}),
    cancel: vi.fn(),
    clear: vi.fn(),
  } as unknown as PresentationProjectController
  const ui: OfficeWorkspaceUi = {
    attachments: () => [],
    skills: () => [],
    skillPackagesEnabled: false,
    upload: vi.fn(),
    clear: vi.fn(),
    project,
  }
  const container = document.createElement('div')
  const root = createRoot(container)
  try {
    await act(async () =>
      root.render(
        React.createElement(AgentWorkspace, {
          session,
          ui,
          host: 'powerpoint',
          disconnect: vi.fn(),
        }),
      ),
    )
    const menu = container.querySelector<HTMLDetailsElement>('[aria-label="完成操作"]')!
    expect(menu.querySelector('summary')?.textContent).toBe('完成')
    const end = Array.from(menu.querySelectorAll('button')).find(
      (button) => button.textContent === '保留成果并结束前台',
    )!
    expect(end.disabled).toBe(false)
    await act(async () => end.click())
    expect(session.stop).toHaveBeenCalledOnce()
    expect(project.cancel).toHaveBeenCalledOnce()
    expect(session.newTask).not.toHaveBeenCalled()
    expect(project.clear).not.toHaveBeenCalled()
    expect(ui.clear).not.toHaveBeenCalled()
    expect(container.textContent).toContain('修改当前页')
  } finally {
    await act(async () => root.unmount())
  }
})
