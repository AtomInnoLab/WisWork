// @vitest-environment jsdom

import React from 'react'
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it, vi } from 'vitest'
import {
  AgentWorkspace,
  DiagnosticCopyButton,
  LegacyAgentWorkspace,
  composerKeyAction,
  createOfficeWorkspaceUi,
  focusWorkspacePanel,
  isTimelineNearBottom,
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

function workspaceMarkup(
  overrides: Partial<OfficeAgentSnapshot> = {},
  panel?: WorkspacePanelName,
  host: 'word' | 'excel' | 'powerpoint' | 'unknown' = 'word',
  connectionNotice?: string,
) {
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
    dispose: vi.fn(),
  }
  const ui: OfficeWorkspaceUi = Object.freeze({
    attachments: () => Object.freeze(['/home/user/source.docx']),
    skills: () => Object.freeze(['Editorial review']),
    skillPackagesEnabled: true,
    upload: vi.fn(),
    copyDiagnostics: vi.fn(),
    clear: vi.fn(),
  })
  return renderToStaticMarkup(
    React.createElement(AgentWorkspace, {
      session,
      ui,
      disconnect: vi.fn(),
      host,
      initialPanel: panel,
      connectionNotice,
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
    expect(html).toContain('新对话')
    expect(html).toContain('AI Word')
    expect(html).not.toContain('WisWork Agent</span>')
    expect(html).not.toContain('Microsoft Word is connected')
    expect(html).not.toContain('Edits always require your approval')
    expect(html).toContain('aria-label="Message WisWork Agent"')
    expect(html).not.toContain('<pre')
  })

  it('announces when the current relay session could not be remembered', () => {
    const notice =
      'Connected, but this Office installation was not remembered. Pair again after reconnecting.'
    const html = workspaceMarkup({}, undefined, 'word', notice)
    expect(html).toContain('role="status"')
    expect(html).toContain(notice)
  })

  it('renders Markdown only for assistant timeline messages', () => {
    const html = workspaceMarkup({
      timeline: Object.freeze([
        Object.freeze({
          id: 'assistant-markdown',
          kind: 'assistant' as const,
          text: '# Result\n\n- **safe** `code`',
          streaming: true,
        }),
        Object.freeze({ id: 'user-plain', kind: 'user' as const, text: '**user stays plain**' }),
        Object.freeze({
          id: 'error-plain',
          kind: 'error' as const,
          text: '*error stays plain*',
        }),
      ]),
    })
    expect(html).toContain('<div class="ai-md">')
    expect(html).toContain('<p class="ai-md-h">Result</p>')
    expect(html).toContain('<ul><li><strong>safe</strong> <code>code</code></li></ul>')
    expect(html).toContain('<span class="streaming-cursor" aria-label="Response streaming"></span>')
    expect(html).toContain('<p>**user stays plain**</p>')
    expect(html).toContain('<p>*error stays plain*</p>')
    expect(html).not.toContain('<strong>user stays plain</strong>')
    expect(html).not.toContain('<em>error stays plain</em>')
  })

  it('keeps approval actionable while the agent loop is suspended', () => {
    const waiting = workspaceMarkup({ busy: true, status: 'working' })
    expect(waiting).toMatch(/<button type="button" class="secondary">Reject<\/button>/)
    expect(waiting).toMatch(/<button type="button">Confirm change<\/button>/)

    const applying = workspaceMarkup({ busy: true, applying: true, status: 'working' })
    expect(applying).toMatch(/<button type="button" class="secondary" disabled="">Reject<\/button>/)
    expect(applying).toMatch(/<button type="button" class="quiet">新对话<\/button>/)
    expect(applying).toContain('<button type="button">退出登录</button>')
    expect(applying).toMatch(/<button type="button" class="stop-button">Stop<\/button>/)
  })

  it('uses the corresponding compact PC editor identity for every Office host', () => {
    expect(workspaceMarkup({}, undefined, 'word')).toContain('AI Word')
    expect(workspaceMarkup({}, undefined, 'excel')).toContain('AI Sheets')
    expect(workspaceMarkup({}, undefined, 'powerpoint')).toContain('AI Slides')
    expect(workspaceMarkup({}, undefined, 'unknown')).toContain('WisWork AI')
  })

  it('uses the desktop Slides conversation hierarchy for PowerPoint', () => {
    const html = workspaceMarkup({}, undefined, 'powerpoint')
    expect(html).toContain('class="agent-workspace presentation-agent')
    expect(html).toContain('class="ai-msg ai-msg-user"')
    expect(html).toContain('class="ai-msg ai-msg-assistant"')
    expect(html).toContain('class="ai-work-group"')
    expect(html).toContain('已完成 · 1 个步骤')
    expect(html).not.toContain('class="tool-event')
    expect(html).not.toContain('message-role')
    expect(html).not.toContain('class="agent-status"')
  })

  it('uses the desktop Slides generation empty state for PowerPoint', () => {
    const html = workspaceMarkup(
      {
        assistantText: '',
        status: 'idle',
        proposal: undefined,
        timeline: Object.freeze([]),
      },
      undefined,
      'powerpoint',
    )
    expect(html).toContain('让 AI 为你生成演示文稿')
    expect(html).toContain('描述主题、场合和大致页数')
    expect(html).toContain('起草一份项目汇报')
    expect(html).toContain('class="ai-starter"')
    expect(html).not.toContain('让 AI 帮你从零起草')
  })

  it('keeps the rollback workspace compact without the legacy explanatory masthead', () => {
    const snapshot = {
      assistantText: '',
      activity: '',
      busy: false,
      applying: false,
      status: 'idle' as const,
      retryable: false,
      timeline: Object.freeze([{ id: 'u1', kind: 'user' as const, text: 'Hello' }]),
    }
    const session = {
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
      dispose: vi.fn(),
    }
    const html = renderToStaticMarkup(
      React.createElement(LegacyAgentWorkspace, {
        session,
        ui: Object.freeze({
          attachments: () => Object.freeze([]),
          skills: () => Object.freeze([]),
          skillPackagesEnabled: false,
          upload: vi.fn(),
          clear: vi.fn(),
        }),
        disconnect: vi.fn(),
        host: 'word',
      }),
    )
    expect(html).toContain('AI Word')
    expect(html).not.toContain('Microsoft Word is connected')
    expect(html).not.toContain('Edits always require your approval')
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

    const stale = workspaceMarkup({
      status: 'error',
      error: 'proposal_stale',
      errorMessage: '文档内容已发生变化，刚才的修改未应用。',
      retryable: true,
      proposal: undefined,
    })
    expect(stale).toContain('>重新生成<')
    expect(stale).toContain('id="instruction"')
    expect(stale).not.toContain('proposal_stale')

    const retryable = workspaceMarkup({
      status: 'error',
      error: 'network_error',
      errorMessage: 'The connection was interrupted. Check WisWork PC and try again.',
      retryable: true,
    })
    expect(retryable).toContain('>Retry<')
  })

  it('keeps internal tool names out of the user-facing activity row', () => {
    const html = workspaceMarkup()
    expect(html).toContain('Prepared change')
    expect(html).not.toContain('>replace<')
  })

  it('shows a readable preview instead of an internal before-only snapshot', () => {
    const structured = {
      id: 'structured-1',
      operation: 'duplicate_slide',
      title: 'Duplicate slide',
      impact: { host: 'powerpoint', targets: ['slide-1'], count: 1 },
      preview: { slideIndex: 3, slideId: 'slide-1' },
      fingerprint: 'fp',
      before: { fingerprint: 'internal-hash', slideId: 'slide-1' },
    }
    const html = workspaceMarkup({
      proposal: structured,
      timeline: Object.freeze([
        Object.freeze({
          id: 'p-structured',
          kind: 'proposal' as const,
          proposal: structured,
          state: 'pending' as const,
        }),
      ]),
    })
    expect(html).toContain('Slide index: 3')
    expect(html).toContain('Slide ID: Slide 1')
    expect(html).not.toContain('internal-hash')
    expect(html).not.toContain('(described by preview)')
  })

  it('shows the complete proposed draft when writing into an empty document', () => {
    const emptyDraft = {
      id: 'empty-draft',
      operation: 'write_document',
      title: 'Write document',
      impact: { host: 'word', targets: ['document:replace'], count: 1 },
      preview: { mode: 'replace' },
      fingerprint: 'fp',
      before: '',
      after: '这是完整草稿。',
    }
    const html = workspaceMarkup({
      proposal: emptyDraft,
      timeline: Object.freeze([
        Object.freeze({
          id: 'p-empty',
          kind: 'proposal' as const,
          proposal: emptyDraft,
          state: 'pending' as const,
        }),
      ]),
    })
    expect(html).toContain('(empty document)')
    expect(html).toContain('这是完整草稿。')
    expect(html).not.toContain('Mode: replace')
  })

  it('uses a frozen UI-only facade instead of exposing the Office runtime', async () => {
    const runtime = {
      vfs: { list: () => ['/home/user/a.txt'] },
      skills: { list: () => [{ name: 'Review' }] },
      skillPackagesEnabled: true,
      uploadFile: vi.fn(),
      installSkill: vi.fn(),
      clearSession: vi.fn(),
    } as unknown as OfficeHostRuntime
    const writeText = vi.fn(async () => undefined)
    const diagnostics = { exportJson: vi.fn(() => '{"version":1}') }
    const ui = createOfficeWorkspaceUi(runtime, diagnostics, { writeText })
    expect(Object.isFrozen(ui)).toBe(true)
    expect(Object.isFrozen(ui.attachments())).toBe(true)
    expect(Object.isFrozen(ui.skills())).toBe(true)
    expect(ui).not.toHaveProperty('vfs')
    expect(ui).not.toHaveProperty('runtime')
    expect(ui).not.toHaveProperty('proposals')
    await expect(ui.copyDiagnostics!()).resolves.toBeUndefined()
    expect(writeText).toHaveBeenCalledWith('{"version":1}')
  })

  it('offers a direct copy-diagnostics action without exposing diagnostic state', () => {
    const html = workspaceMarkup()
    expect(html).toContain('复制诊断信息')
    expect(html).not.toContain('trace_id')
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
      dispose: vi.fn(),
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

  it('shows sighted users whether copying the bounded diagnostic export succeeded', async () => {
    const snapshot: OfficeAgentSnapshot = {
      assistantText: 'Done',
      activity: '',
      busy: false,
      applying: false,
      status: 'done',
      retryable: false,
      timeline: Object.freeze([
        Object.freeze({ id: 'a1', kind: 'assistant' as const, text: 'Done' }),
      ]),
    }
    const session = {
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
      dispose: vi.fn(),
    } satisfies OfficeAgentSession
    const copyDiagnostics = vi.fn(async () => undefined)
    const ui: OfficeWorkspaceUi = Object.freeze({
      attachments: () => Object.freeze([]),
      skills: () => Object.freeze([]),
      skillPackagesEnabled: true,
      upload: vi.fn(),
      copyDiagnostics,
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
    const copy = Array.from(container.querySelectorAll('button')).find(
      (button) => button.textContent === '复制诊断信息',
    )!
    await act(async () => copy.click())
    expect(copyDiagnostics).toHaveBeenCalledOnce()
    expect(container.querySelector('.diagnostic-status')?.textContent).toBe('诊断信息已复制')
    await act(async () => root.unmount())
    container.remove()
  })

  it('offers the same diagnostic copy feedback on a disconnected status screen', async () => {
    const copyDiagnostics = vi.fn(async () => undefined)
    const container = document.createElement('div')
    document.body.append(container)
    const root = createRoot(container)
    await act(async () => {
      root.render(React.createElement(DiagnosticCopyButton, { copyDiagnostics }))
    })
    const button = container.querySelector('button')!
    await act(async () => button.click())
    expect(copyDiagnostics).toHaveBeenCalledOnce()
    expect(container.querySelector('[role="status"]')?.textContent).toBe('诊断信息已复制')
    await act(async () => root.unmount())
    container.remove()
  })

  it('sends on Enter, preserves multiline input on Shift+Enter, and ignores composition', () => {
    expect(composerKeyAction({ key: 'Enter', shiftKey: false, isComposing: false })).toBe('send')
    expect(composerKeyAction({ key: 'Enter', shiftKey: true, isComposing: false })).toBe('newline')
    expect(composerKeyAction({ key: 'Enter', shiftKey: false, isComposing: true })).toBe('newline')
    expect(composerKeyAction({ key: 'Escape', shiftKey: false, isComposing: false })).toBe('none')
  })

  it('keeps automatic scrolling only while the reader remains near the latest turn', () => {
    expect(isTimelineNearBottom({ scrollHeight: 900, scrollTop: 580, clientHeight: 300 })).toBe(
      true,
    )
    expect(isTimelineNearBottom({ scrollHeight: 900, scrollTop: 300, clientHeight: 300 })).toBe(
      false,
    )
  })
})
