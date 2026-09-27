// @vitest-environment jsdom
import React, { act } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it, vi } from 'vitest'
import { AgentWorkspace, createOfficeWorkspaceUi, type OfficeWorkspaceUi } from '../src/App.js'
import type { OfficeAgentSession, OfficeAgentSnapshot } from '../src/agent/use-office-agent.js'
import type { OfficeHostRuntime } from '../src/agent/host-runtime.js'
import type { PresentationChangesController } from '../src/agent/presentation-changes.js'
import type { PresentationProjectController } from '../src/skills/powerpoint/presentation-project.js'

it('exposes the optional project controller in the UI facade', () => {
  const project = {} as PresentationProjectController
  expect(
    createOfficeWorkspaceUi({ presentation: project } as unknown as OfficeHostRuntime).project,
  ).toBe(project)
})

describe('workspace project integration', () => {
  it('refreshes after work, observes restored downloads, and guards sends during recovery', async () => {
    let snapshot: OfficeAgentSnapshot = {
      assistantText: '',
      activity: '',
      busy: false,
      applying: false,
      status: 'idle',
      retryable: false,
      timeline: [],
    }
    const listeners = new Set<() => void>()
    const session: OfficeAgentSession = {
      snapshot: () => snapshot,
      subscribe: (fn) => {
        listeners.add(fn)
        return () => {
          listeners.delete(fn)
        }
      },
      send: vi.fn(),
      stop: vi.fn(),
      confirm: vi.fn(),
      reject: vi.fn(),
      newTask: vi.fn(),
      retry: vi.fn(),
      logout: vi.fn(),
      authenticationLost: vi.fn(),
      dispose: vi.fn(),
    }
    let projectSnapshot: ReturnType<PresentationProjectController['snapshot']> = { phase: 'idle' }
    const projectListeners = new Set<() => void>()
    const project: PresentationProjectController = {
      snapshot: () => projectSnapshot,
      subscribe: (fn) => {
        projectListeners.add(fn)
        return () => {
          projectListeners.delete(fn)
        }
      },
      readDeliveryReport: vi.fn(async () => {}),
      exportDeliveryReport: vi.fn(async () => {}),
      recordIssueAction: vi.fn(async () => {}),
      refresh: vi.fn(async () => {}),
      restore: vi.fn(async () => {}),
      resume: vi.fn(async () => {}),
      runProduction: vi.fn(async () => {}),
      selectProduction: vi.fn(async () => {}),
      startProductionJob: vi.fn(async () => {}),
      pauseProductionJob: vi.fn(async () => {}),
      resumeProductionJob: vi.fn(async () => {}),
      cancelProductionJob: vi.fn(async () => {}),
      downloadProductionPage: vi.fn(async () => {}),
      prepareProduction: vi.fn(async () => {}),
      cancel: vi.fn(),
      clear: vi.fn(),
      prepareReconnect: vi.fn(),
    }
    const changes: PresentationChangesController = {
      snapshot: () => ({
        phase: 'idle',
        projectId: 'project',
        requestId: 'request',
        entries: [
          {
            id: 'text:change',
            kind: 'text',
            pageId: 'page1',
            state: 'applied',
            before: 'original text',
            after: 'revised text',
            actions: ['undo'],
          },
        ],
      }),
      subscribe: () => () => {},
      refresh: vi.fn(async () => {}),
      run: vi.fn(async () => {}),
      clear: vi.fn(),
    }
    expect(createOfficeWorkspaceUi({ changes } as unknown as OfficeHostRuntime).changes).toBe(
      changes,
    )
    let files: string[] = []
    const ui: OfficeWorkspaceUi = {
      project,
      changes,
      attachments: () => files,
      downloadFile: vi.fn(),
      skills: () => [],
      skillPackagesEnabled: false,
      upload: vi.fn(),
      clear: vi.fn(),
    }
    const container = document.createElement('div')
    const root = createRoot(container)
    const update = async (changes: Partial<OfficeAgentSnapshot>) => {
      snapshot = { ...snapshot, ...changes }
      await act(async () => listeners.forEach((fn) => fn()))
    }
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
      expect(container.querySelector('.composer-shell [aria-label="演示文稿项目"]')).not.toBeNull()
      expect(project.refresh).toHaveBeenCalledTimes(1)
      expect(
        container.querySelector('.composer-shell [aria-label="修改保存点工作台"]'),
      ).not.toBeNull()
      const undo = container.querySelector<HTMLButtonElement>('[aria-label="撤销 page1"]')!
      await act(async () => undo.click())
      expect(changes.run).toHaveBeenCalledWith('text:change', 'undo')
      expect(session.confirm).not.toHaveBeenCalled()
      await update({ busy: true })
      expect(undo.disabled).toBe(true)
      await act(async () => undo.click())
      expect(changes.run).toHaveBeenCalledTimes(1)
      await update({ timeline: [{ id: 'a', kind: 'assistant', text: 'Draft', streaming: true }] })
      expect(project.refresh).toHaveBeenCalledTimes(1)
      await update({ busy: false })
      expect(project.refresh).toHaveBeenCalledTimes(2)
      // A starter prompt fills the composer, then recovery disables both click and Enter sends.
      await update({ timeline: [] })
      await act(async () =>
        Array.from(container.querySelectorAll('button'))
          .find((button) => button.textContent === '起草一份项目汇报')!
          .click(),
      )
      await update({ error: 'network_error', errorMessage: '连接中断', retryable: true })
      projectSnapshot = { phase: 'restoring' }
      await act(async () => projectListeners.forEach((fn) => fn()))
      expect(undo.disabled).toBe(true)
      expect(undo.closest('section')?.textContent).toContain('original text')
      expect(undo.closest('section')?.textContent).toContain('revised text')
      expect(
        container.querySelector<HTMLButtonElement>('[aria-label="Send message"]')!.disabled,
      ).toBe(true)
      await act(async () =>
        container
          .querySelector('textarea')!
          .dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })),
      )
      expect(session.send).not.toHaveBeenCalled()
      const retry = Array.from(container.querySelectorAll('button')).find(
        (button) => button.textContent === 'Retry',
      )!
      expect(retry.disabled).toBe(true)
      await act(async () => retry.click())
      expect(session.retry).not.toHaveBeenCalled()
      files = ['/home/user/generated/recovered.pptx']
      projectSnapshot = { phase: 'idle' }
      await act(async () => projectListeners.forEach((fn) => fn()))
      expect(container.textContent).toContain('下载 PPTX')
      await act(async () =>
        container.querySelector<HTMLButtonElement>('[aria-label="Send message"]')!.click(),
      )
      expect(session.send).toHaveBeenCalledWith('起草一份项目汇报')
    } finally {
      await act(async () => root.unmount())
    }
  })
})
