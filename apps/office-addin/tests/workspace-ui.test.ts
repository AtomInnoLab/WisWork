// @vitest-environment jsdom

import React from 'react'
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it, vi } from 'vitest'
import {
  AgentWorkspace,
  composerKeyAction,
  createOfficeWorkspaceUi,
  focusWorkspacePanel,
  type OfficeWorkspaceUi,
  type WorkspacePanelName,
} from '../src/App.js'
import type { OfficeAgentSession, OfficeAgentSnapshot } from '../src/agent/use-office-agent.js'
import type { OfficeHostRuntime } from '../src/agent/host-runtime.js'

const proposal = {
  id: 'proposal-1',
  operation: 'replace' as const,
  before: 'old',
  value: 'new',
  fingerprint: 'fp',
}

function workspaceMarkup(overrides: Partial<OfficeAgentSnapshot> = {}, panel?: WorkspacePanelName) {
  const snapshot: OfficeAgentSnapshot = {
    assistantText: 'Draft ready',
    activity: '',
    busy: false,
    applying: false,
    status: 'done',
    retryable: false,
    proposal,
    timeline: Object.freeze([
      Object.freeze({ id: 'u1', kind: 'user' as const, text: 'Rewrite this' }),
      Object.freeze({ id: 'a1', kind: 'assistant' as const, text: 'Draft ready' }),
      Object.freeze({
        id: 't1',
        kind: 'tool' as const,
        callId: 'call-1',
        name: 'replace',
        summary: 'Prepared change',
        state: 'complete' as const,
      }),
      Object.freeze({
        id: 'p1',
        kind: 'proposal' as const,
        proposal,
        state: 'pending' as const,
      }),
    ]),
    ...overrides,
  }
  const session: OfficeAgentSession = {
    snapshot: () => snapshot,
    subscribe: () => () => undefined,
    send: vi.fn(),
    stop: vi.fn(),
    confirm: vi.fn(),
    reject: vi.fn(),
    newTask: vi.fn(),
    retry: vi.fn(),
    logout: vi.fn(),
    authenticationLost: vi.fn(),
  }
  const ui: OfficeWorkspaceUi = Object.freeze({
    attachments: () => Object.freeze(['/home/user/source.docx']),
    skills: () => Object.freeze(['Editorial review']),
    skillPackagesEnabled: true,
    upload: vi.fn(),
    clear: vi.fn(),
  })
  return renderToStaticMarkup(
    React.createElement(AgentWorkspace, {
      session,
      ui,
      disconnect: vi.fn(),
      host: 'word',
      initialPanel: panel,
    }),
  )
}

describe('Office Agent workspace UI', () => {
  it('renders an accessible full workspace with causal timeline and inline proposal actions', () => {
    const html = workspaceMarkup()
    expect(html).toContain('aria-label="Agent conversation"')
    expect(html).toContain('Rewrite this')
    expect(html).toContain('Prepared change')
    expect(html.indexOf('Prepared change')).toBeLessThan(html.indexOf('Approval required'))
    expect(html).toContain('Confirm change')
    expect(html).toContain('New task')
    expect(html).toContain('aria-label="Message WisWork Agent"')
  })

  it('matches the WisWork writing-first empty state without legacy selection or session-file chrome', () => {
    const html = workspaceMarkup({
      assistantText: '',
      status: 'idle',
      proposal: undefined,
      timeline: Object.freeze([]),
    })
    expect(html).toContain('让 AI 帮你从零起草')
    expect(html).toContain('描述主题、要点或粘贴参考素材')
    expect(html).toContain('帮我写一份项目周报')
    expect(html).toContain('写一篇产品发布公告')
    expect(html).toContain('列一个活动策划提纲')
    expect(html).toContain('描述修改、写作要求，或直接提问')
    expect(html).toContain('更改需确认')
    expect(html).not.toContain('Work with your selection')
    expect(html).not.toContain('Session files')
    expect(html).not.toContain('Agent is ready')
    expect(html).not.toContain('class="app-header"')
  })

  it('exposes bounded attachment and skill management panels without permanent vertical chrome', () => {
    const files = workspaceMarkup({}, 'attachments')
    expect(files).toContain('role="dialog"')
    expect(files).toContain('Session attachments')
    expect(files).toContain('source.docx')
    expect(files).not.toContain('Editorial review')

    const skills = workspaceMarkup({}, 'skills')
    expect(skills).toContain('Installed skills')
    expect(skills).toContain('Editorial review')
    expect(skills).toContain('Install skill package')
    expect(skills).toContain('accept=".zip,application/zip"')
    expect(skills).toContain('>Remove</button>')
  })

  it('renders distinct accessible working, applying, stop, and retry states', () => {
    const working = workspaceMarkup({ busy: true, status: 'working' })
    expect(working).toContain('Agent is working')
    expect(working).toContain('>Stop<')
    expect(working).toContain('aria-busy="true"')

    const applying = workspaceMarkup({ applying: true })
    expect(applying).toContain('Applying approved change')
    expect(applying).toContain('Applying…')

    expect(applying).toMatch(/aria-label="Attachments"[^>]*disabled/)
    expect(applying).toMatch(/<button type="button" disabled="">管理技能<\/button>/)
    const applyingPanel = workspaceMarkup({ applying: true }, 'attachments')
    expect(applyingPanel).toMatch(/class="upload-button"[^>]*aria-disabled="true"/)
    expect(applyingPanel).toMatch(/id="session-upload"[^>]*disabled/)

    const failed = workspaceMarkup({
      status: 'error',
      error: 'office_write_failed',
      errorMessage: 'The approved change could not be applied.',
      retryable: false,
    })
    expect(failed).toContain('The approved change could not be applied.')
    expect(failed).not.toContain('>Retry<')
    expect(failed).toContain('role="alert"')

    const retryable = workspaceMarkup({
      status: 'error',
      error: 'network_error',
      errorMessage: 'The connection was interrupted. Check WisWork PC and try again.',
      retryable: true,
    })
    expect(retryable).toContain('>Retry<')
  })

  it('uses a frozen UI-only facade instead of exposing the Office runtime', () => {
    const runtime = {
      vfs: { list: () => ['/home/user/a.txt'] },
      skills: { list: () => [{ name: 'Review' }] },
      skillPackagesEnabled: true,
      uploadFile: vi.fn(),
      installSkill: vi.fn(),
      clearSession: vi.fn(),
    } as unknown as OfficeHostRuntime
    const ui = createOfficeWorkspaceUi(runtime)
    expect(Object.isFrozen(ui)).toBe(true)
    expect(Object.isFrozen(ui.attachments())).toBe(true)
    expect(Object.isFrozen(ui.skills())).toBe(true)
    expect(ui).not.toHaveProperty('vfs')
    expect(ui).not.toHaveProperty('runtime')
    expect(ui).not.toHaveProperty('proposals')
  })

  it('focuses the panel heading on open and restores the opener on close', () => {
    const heading = { focus: vi.fn() }
    const opener = { focus: vi.fn() }
    const restore = focusWorkspacePanel(heading, opener)
    expect(heading.focus).toHaveBeenCalledOnce()
    restore()
    expect(opener.focus).toHaveBeenCalledOnce()
  })

  it('moves real DOM focus into an opened panel and restores it after Escape', async () => {
    const snapshot: OfficeAgentSnapshot = {
      assistantText: '',
      activity: '',
      busy: false,
      applying: false,
      status: 'idle',
      retryable: false,
      timeline: Object.freeze([]),
    }
    const session: OfficeAgentSession = {
      snapshot: () => snapshot,
      subscribe: () => () => undefined,
      send: vi.fn(),
      stop: vi.fn(),
      confirm: vi.fn(),
      reject: vi.fn(),
      newTask: vi.fn(),
      retry: vi.fn(),
      logout: vi.fn(),
      authenticationLost: vi.fn(),
    }
    const ui: OfficeWorkspaceUi = Object.freeze({
      attachments: () => Object.freeze([]),
      skills: () => Object.freeze([]),
      skillPackagesEnabled: true,
      upload: vi.fn(),
      clear: vi.fn(),
    })
    const container = document.createElement('div')
    document.body.append(container)
    const root = createRoot(container)
    await act(async () => {
      root.render(
        React.createElement(AgentWorkspace, {
          session,
          ui,
          disconnect: vi.fn(),
          host: 'word',
        }),
      )
    })
    const opener = container.querySelector<HTMLButtonElement>('[aria-label="Attachments"]')!
    await act(async () => opener.click())
    const dialog = container.querySelector<HTMLElement>('[role="dialog"]')!
    expect(document.activeElement).toBe(dialog.querySelector('h2'))

    await act(async () => {
      dialog.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    })
    expect(container.querySelector('[role="dialog"]')).toBeNull()
    expect(document.activeElement).toBe(opener)
    await act(async () => root.unmount())
    container.remove()
  })

  it('sends on Enter, preserves multiline input on Shift+Enter, and ignores composition', () => {
    expect(composerKeyAction({ key: 'Enter', shiftKey: false, isComposing: false })).toBe('send')
    expect(composerKeyAction({ key: 'Enter', shiftKey: true, isComposing: false })).toBe('newline')
    expect(composerKeyAction({ key: 'Enter', shiftKey: false, isComposing: true })).toBe('newline')
    expect(composerKeyAction({ key: 'Escape', shiftKey: false, isComposing: false })).toBe('none')
  })
})
