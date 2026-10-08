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
  presentationDesignLifecycle,
  type OfficeWorkspaceUi,
  type WorkspacePanelName,
} from '../src/App.js'
import {
  createOfficeAgentSession,
  type OfficeAgentSession,
  type OfficeAgentSnapshot,
} from '../src/agent/use-office-agent.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import type { OfficeToolActivity } from '../src/agent/transport.js'
import type { OfficeHostRuntime } from '../src/agent/host-runtime.js'
import { createOfficeDiagnostics } from '../src/diagnostics/office-diagnostics.js'
import type { PresentationQaAttempt } from '../src/skills/powerpoint/presentation-qa-attempts.js'

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
  runtimeMode: 'standard' | 'enhanced' = 'standard',
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
      runtimeMode,
    }),
  )
}

describe('Office Agent workspace UI', () => {
  it('does not revive an older unreceived result when a new presentation request starts', () => {
    const timeline: OfficeAgentSnapshot['timeline'] = [
      { id: 'old-user', kind: 'user', text: 'Old request' },
      {
        id: 'old-tool',
        kind: 'tool',
        callId: 'old-call',
        name: 'run_presentation_production',
        summary: 'Running',
        state: 'running',
      },
      { id: 'new-user', kind: 'user', text: 'New request' },
      {
        id: 'new-tool',
        kind: 'tool',
        callId: 'new-call',
        name: 'run_presentation_production',
        summary: 'Running',
        state: 'running',
      },
    ]
    const container = document.createElement('div')
    container.innerHTML = workspaceMarkup({ timeline, busy: true }, undefined, 'powerpoint')
    const stages = container.querySelectorAll('.stage-event')
    expect(stages[0]?.getAttribute('aria-busy')).toBe('false')
    expect(stages[0]?.textContent).toContain('尚未收到最终回执')
    expect(stages[1]?.getAttribute('aria-busy')).toBe('true')
    expect(stages[1]?.textContent).toContain('当前操作仍在进行')
  })

  it('folds presentation stages while keeping confirmation outside the stage details', () => {
    const timeline: OfficeAgentSnapshot['timeline'] = [
      { id: 'u1', kind: 'user', text: '制作演示文稿' },
      {
        id: 't1',
        kind: 'tool',
        callId: 'c1',
        name: 'read_presentation_plan',
        summary: 'Ready',
        state: 'complete',
      },
      {
        id: 't2',
        kind: 'tool',
        callId: 'c2',
        name: 'save_presentation_plan',
        summary: 'Failed',
        state: 'error',
      },
      { id: 'p1', kind: 'proposal', proposal, state: 'pending' },
      {
        id: 't3',
        kind: 'tool',
        callId: 'c3',
        name: 'read_presentation_plan',
        summary: 'Reading',
        state: 'running',
      },
    ]
    const container = document.createElement('div')
    container.innerHTML = workspaceMarkup({ timeline }, undefined, 'powerpoint')
    const stages = container.querySelectorAll('.stage-event')
    expect(stages).toHaveLength(2)
    expect(stages[0]?.querySelector('details')?.open).toBe(false)
    expect(stages[0]?.querySelector('summary')?.textContent).toContain('2 项')
    expect(stages[0]?.querySelector('[role="alert"]')?.textContent).toContain('恢复记录')
    const approval = Array.from(container.querySelectorAll('button')).find(
      (button) => button.textContent === 'Confirm change',
    )!
    expect(approval.disabled).toBe(false)
    expect(approval.closest('.stage-event')).toBeNull()
    expect(workspaceMarkup({ timeline }, undefined, 'word')).not.toContain('stage-event')
  })

  it('requires a user confirmation before deleting a PC attachment', async () => {
    const id = 'a'.repeat(64)
    const list = vi.fn().mockResolvedValue([
      {
        attachmentId: id,
        name: 'source.pdf',
        sizeBytes: 1,
        sha256: id,
        receivedBytes: 1,
        status: 'ready',
        kind: 'text',
        totalChars: 1,
      },
    ])
    const remove = vi.fn().mockResolvedValue(undefined)
    const importUrl = vi.fn().mockResolvedValue(undefined)
    const upload = vi.fn().mockResolvedValue(undefined)
    const confirm = vi.spyOn(window, 'confirm').mockReturnValue(false)
    const snapshot = {
      assistantText: '',
      activity: '',
      busy: false,
      applying: false,
      status: 'done',
      retryable: false,
      timeline: Object.freeze([]),
    } as OfficeAgentSnapshot
    const session = {
      snapshot: () => snapshot,
      subscribe: () => () => undefined,
    } as unknown as OfficeAgentSession
    const ui = {
      attachments: () => [],
      skills: () => [],
      skillPackagesEnabled: true,
      upload,
      clear: vi.fn(),
      durableAttachmentsAvailable: () => true,
      durableImagesAvailable: () => true,
      remoteImagesAvailable: () => true,
      listDurableAttachments: list,
      deleteDurableAttachment: remove,
      importPresentationImageUrl: importUrl,
    } as OfficeWorkspaceUi
    const container = document.createElement('div')
    document.body.append(container)
    const root = createRoot(container)
    await act(async () =>
      root.render(
        React.createElement(AgentWorkspace, {
          session,
          ui,
          disconnect: vi.fn(),
          host: 'powerpoint',
        }),
      ),
    )
    await act(async () =>
      container.querySelector<HTMLButtonElement>('[aria-label="Attachments"]')!.click(),
    )
    const button = Array.from(container.querySelectorAll('button')).find(
      (item) => item.textContent === '删除 PC 副本',
    )!
    await act(async () => button.click())
    expect(remove).not.toHaveBeenCalled()
    confirm.mockReturnValue(true)
    await act(async () => button.click())
    expect(remove).toHaveBeenCalledWith(id)
    const input = container.querySelector<HTMLTextAreaElement>('#presentation-image-url')!
    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!.call(
        input,
        'https://example.com/image.png',
      )
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })
    await act(async () =>
      input
        .closest('form')!
        .dispatchEvent(new Event('submit', { bubbles: true, cancelable: true })),
    )
    expect(importUrl).toHaveBeenCalledWith('https://example.com/image.png')
    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!.call(
        input,
        'https://example.com/failed.png\nhttps://example.org/backup.webp',
      )
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })
    await act(async () =>
      input
        .closest('form')!
        .dispatchEvent(new Event('submit', { bubbles: true, cancelable: true })),
    )
    expect(importUrl).toHaveBeenCalledWith([
      'https://example.com/failed.png',
      'https://example.org/backup.webp',
    ])
    const pasted = container.querySelector<HTMLTextAreaElement>('#presentation-pasted-source')!
    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!.call(
        pasted,
        '用户提供的原文资料',
      )
      pasted.dispatchEvent(new Event('input', { bubbles: true }))
    })
    list.mockRejectedValueOnce(new Error('refresh unavailable'))
    await act(async () =>
      pasted
        .closest('form')!
        .dispatchEvent(new Event('submit', { bubbles: true, cancelable: true })),
    )
    expect(upload).toHaveBeenCalledWith(
      expect.objectContaining({ name: expect.stringMatching(/^粘贴资料-.*\.txt$/) }),
    )
    expect(container.textContent).toContain('粘贴资料已保存到当前文档')
    await act(async () => root.unmount())
    container.remove()
    confirm.mockRestore()
  })
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

  it('disables approval until lock review succeeds and displays explicit locked-page coverage', () => {
    const base = {
      id: 'locked-proposal',
      operation: 'edit_existing_presentation_text',
      title: 'Edit locked page',
      preview: {},
      impact: { host: 'powerpoint', targets: ['host'], count: 1 },
      fingerprint: 'fp',
    }
    for (const state of ['checking', 'unavailable', 'ready'] as const) {
      const lockReview =
        state === 'ready'
          ? {
              state,
              token: 'secret-token',
              pages: [
                { projectId: 'p', pageId: 'page', title: '财务结论', slideIds: ['host', 'copy'] },
              ],
            }
          : { state }
      const value = { ...base, lockReview }
      const html = workspaceMarkup(
        {
          proposal: value,
          timeline: [{ id: 'lock-event', kind: 'proposal', state: 'pending', proposal: value }],
        },
        undefined,
        'powerpoint',
      )
      const dom = new DOMParser().parseFromString(html, 'text/html')
      const button = Array.from(dom.querySelectorAll('button')).find((item) =>
        ['Confirm change', '确认本次覆盖锁页'].includes(item.textContent!),
      )!
      expect(button.disabled).toBe(state !== 'ready')
      expect(html).not.toContain('secret-token')
      if (state === 'ready') {
        expect(html).toContain('财务结论')
        expect(html).toContain('2 个宿主副本')
        expect(html).toContain('页面保持锁定')
      }
    }
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

  it('shows the desktop Slides working indicator and descriptive search step', () => {
    const html = workspaceMarkup(
      {
        busy: true,
        status: 'working',
        activity: '图片搜索',
        timeline: Object.freeze([
          {
            id: 'search-1',
            kind: 'tool' as const,
            callId: 'call-search-1',
            name: 'image_search',
            summary: '图片搜索',
            state: 'running' as const,
          },
        ]),
      },
      undefined,
      'powerpoint',
    )
    expect(html).toContain('图片搜索')
    expect(html).not.toContain('aria-label="继续处理中"')
  })

  it('keeps completed PowerPoint tool output available as expandable detail', () => {
    const html = workspaceMarkup(
      {
        timeline: Object.freeze([
          {
            id: 'search-1',
            kind: 'tool' as const,
            callId: 'call-search-1',
            name: 'image_search',
            summary: '图片搜索完成',
            state: 'complete' as const,
            output: '{"images":[{"title":"LLM"}]}',
          },
        ]),
      },
      undefined,
      'powerpoint',
    )
    expect(html).toContain('class="ai-step-title clickable"')
    expect(html).toContain('aria-expanded="false"')
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

  it('shows uninterrupted ordinary editing for PowerPoint while keeping elevated review implicit', () => {
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
    expect(html).toContain('自动应用常规更改')
    expect(html).not.toContain('更改需确认')
  })

  it('keeps the PowerPoint composer structure stable and opens the native file picker directly', async () => {
    const snapshot: OfficeAgentSnapshot = {
      assistantText: '',
      activity: '',
      busy: false,
      applying: false,
      status: 'idle',
      retryable: false,
      timeline: Object.freeze([]),
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
          host: 'powerpoint',
        }),
      )
    })

    const picker = container.querySelector<HTMLInputElement>('#composer-attachment-upload')!
    const click = vi.spyOn(picker, 'click')
    await act(async () =>
      container.querySelector<HTMLButtonElement>('[aria-label="Add attachments"]')!.click(),
    )

    expect(click).toHaveBeenCalledOnce()
    expect(container.querySelector('.composer-input-box')).not.toBeNull()
    expect(container.querySelector('[role="dialog"]')).toBeNull()
    await act(async () => root.unmount())
    container.remove()
  })

  it('does not replace model-authored progress with a generic continuing placeholder', () => {
    const html = workspaceMarkup(
      {
        busy: true,
        status: 'working',
        timeline: Object.freeze([
          Object.freeze({ id: 'u1', kind: 'user' as const, text: '制作 PPT' }),
          Object.freeze({ id: 'a1', kind: 'assistant' as const, text: '先整理页面结构。' }),
          Object.freeze({
            id: 't1',
            kind: 'tool' as const,
            callId: 'tool-1',
            name: 'plan_deck',
            summary: '规划演示文稿完成',
            state: 'complete' as const,
          }),
        ]),
      },
      undefined,
      'powerpoint',
    )
    expect(html).toContain('先整理页面结构。')
    expect(html).not.toContain('继续处理中')
  })

  it.each([
    ['draft', 1, 'DESIGN.md · 已创建'],
    ['ready', 1, 'DESIGN.md · 已锁定'],
    ['draft', 2, 'DESIGN.md · 已修订'],
    ['verified', 2, 'DESIGN.md · 已验证'],
  ])('renders the %s revision lifecycle in the timeline', (status, revision, label) => {
    const output = JSON.stringify({
      status,
      revision,
      designMd: `# DESIGN.md\n\nStatus: ${status}\nRevision: ${revision}`,
    })
    const html = workspaceMarkup(
      {
        timeline: Object.freeze([
          Object.freeze({ id: 'u1', kind: 'user' as const, text: '制作 PPT' }),
          Object.freeze({
            id: 't1',
            kind: 'tool' as const,
            callId: 'tool-1',
            name: 'plan_deck',
            summary: '规划演示文稿完成',
            state: 'complete' as const,
            output,
          }),
        ]),
      },
      undefined,
      'powerpoint',
    )
    expect(html).toContain(label)
  })

  it('opens only draft and ready design lifecycle states as editable', () => {
    expect(presentationDesignLifecycle('{"status":"draft","revision":2}')?.editable).toBe(true)
    expect(presentationDesignLifecycle('{"status":"ready","revision":2}')?.editable).toBe(true)
    expect(presentationDesignLifecycle('{"status":"producing","revision":2}')?.editable).toBe(false)
    expect(presentationDesignLifecycle('{"status":"verified","revision":2}')?.editable).toBe(false)
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
    expect(applying).toContain('>Stop<')
    expect(applying).not.toContain('class="send-button"')

    expect(applying).toMatch(/aria-label="Add attachments"[^>]*disabled/)
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
    expect(html).toContain('空白内容')
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
    const diagnostics = {
      exportJson: vi.fn((options?: { includeLocalContext?: boolean }) =>
        options?.includeLocalContext ? '{"version":1,"context":true}' : '{"version":1}',
      ),
    }
    const ui = createOfficeWorkspaceUi(runtime, diagnostics, { writeText })
    expect(Object.isFrozen(ui)).toBe(true)
    expect(Object.isFrozen(ui.attachments())).toBe(true)
    expect(Object.isFrozen(ui.skills())).toBe(true)
    expect(ui).not.toHaveProperty('vfs')
    expect(ui).not.toHaveProperty('runtime')
    expect(ui).not.toHaveProperty('proposals')
    await expect(ui.copyDiagnostics!()).resolves.toBeUndefined()
    expect(writeText).toHaveBeenCalledWith('{"version":1}')
    expect(diagnostics.exportJson).toHaveBeenLastCalledWith()
    await expect(ui.copyDiagnosticsWithContext!()).resolves.toBeUndefined()
    expect(writeText).toHaveBeenCalledWith('{"version":1,"context":true}')
    expect(diagnostics.exportJson).toHaveBeenLastCalledWith({ includeLocalContext: true })
  })

  it('reads local screenshot attempts only for an explicit contextual diagnostic copy', async () => {
    const closed: PresentationQaAttempt = {
      version: 1,
      id: '12345678-1234-4234-8234-123456789abc',
      source: 'production',
      documentId: 'doc',
      projectId: 'project',
      requestId: 'run',
      pageId: 'page',
      hostSlideId: 'new-slide',
      artifactDigest: 'a'.repeat(64),
      startedAt: '2026-09-29T00:00:00.000Z',
      status: 'closed',
      finishedAt: '2026-09-29T00:01:00.000Z',
      errorCode: 'explicitly_closed',
    }
    let fail = false
    const attempts = vi.fn(() => {
      if (fail) throw Error('secret SDK failure')
      return [closed]
    })
    const runtime = {
      vfs: { list: () => [] },
      skills: { list: () => [] },
      qa: { attempts },
    } as unknown as OfficeHostRuntime
    const send = vi.fn()
    const diagnostics = createOfficeDiagnostics({
      host: 'powerpoint',
      build: 'test',
      localDocumentId: 'doc',
      remoteEnabled: true,
      send,
    })
    diagnostics.startTrace()
    diagnostics.setTool('run_presentation_production', { project_id: 'project', request_id: 'run' })
    diagnostics.record({ phase: 'tool', errorCode: 'office_write_failed' })
    const sentBefore = send.mock.calls.length
    const writeText = vi.fn(async (_value: string) => {})
    const ui = createOfficeWorkspaceUi(runtime, diagnostics, { writeText })
    expect(attempts).not.toHaveBeenCalled()
    await ui.copyDiagnostics!()
    expect(attempts).not.toHaveBeenCalled()
    expect(writeText.mock.calls.at(-1)![0]).not.toContain('local_presentation_qa_attempts')
    await ui.copyDiagnosticsWithContext!()
    expect(attempts).toHaveBeenCalledOnce()
    const available = JSON.parse(writeText.mock.calls.at(-1)![0])
    expect(available.local_presentation_qa_attempts).toMatchObject({
      scope: 'retained_visible_presentation_task',
      status: 'available',
      attempts: [closed],
      record_count: 1,
      unresolved_count: 0,
    })
    fail = true
    await ui.copyDiagnosticsWithContext!()
    const unavailable = JSON.parse(writeText.mock.calls.at(-1)![0])
    expect(unavailable.local_presentation_qa_attempts).toMatchObject({ status: 'unavailable' })
    expect(JSON.stringify(unavailable)).not.toContain('secret SDK failure')
    expect(JSON.stringify(unavailable.local_presentation_qa_attempts)).not.toContain(closed.id)
    expect(send).toHaveBeenCalledTimes(sentBefore)
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
          host: 'powerpoint',
        }),
      )
    })
    const opener = Array.from(container.querySelectorAll<HTMLButtonElement>('button')).find(
      (button) => button.textContent === '管理技能',
    )!
    await act(async () => opener.click())
    const dialog = container.querySelector<HTMLElement>('[role="dialog"]')!
    expect(document.activeElement).toBe(dialog.querySelector('h2'))

    await act(async () => {
      dialog.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    })
    expect(container.querySelector('[role="dialog"]')).toBeNull()
    expect(document.activeElement).toBe(container.querySelector('[aria-label="Session menu"]'))
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
    const copyDiagnosticsWithContext = vi.fn(async () => undefined)
    const ui: OfficeWorkspaceUi = Object.freeze({
      attachments: () => Object.freeze([]),
      skills: () => Object.freeze([]),
      skillPackagesEnabled: true,
      upload: vi.fn(),
      copyDiagnostics,
      copyDiagnosticsWithContext,
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
    expect(container.querySelector('.diagnostic-status')?.textContent).toBe(
      '诊断信息已复制；已移除文档、会话、项目和页面 ID',
    )
    const withContext = Array.from(container.querySelectorAll('button')).find(
      (button) => button.textContent === '复制含定位 ID 的诊断',
    )!
    await act(async () => withContext.click())
    expect(copyDiagnosticsWithContext).toHaveBeenCalledOnce()
    expect(container.querySelector('.diagnostic-status')?.textContent).toBe(
      '已复制含定位 ID 的本机诊断；请检查后分享',
    )
    await act(async () => root.unmount())
    container.remove()
  })

  it('offers the same diagnostic copy feedback on a disconnected status screen', async () => {
    const copyDiagnostics = vi.fn(async () => undefined)
    const copyDiagnosticsWithContext = vi.fn(async () => undefined)
    const container = document.createElement('div')
    document.body.append(container)
    const root = createRoot(container)
    await act(async () => {
      root.render(
        React.createElement(DiagnosticCopyButton, {
          copyDiagnostics,
          copyDiagnosticsWithContext,
        }),
      )
    })
    const button = container.querySelector('button')!
    await act(async () => button.click())
    expect(copyDiagnostics).toHaveBeenCalledOnce()
    expect(container.querySelector('[role="status"]')?.textContent).toBe(
      '诊断信息已复制；已移除文档、会话、项目和页面 ID',
    )
    const withContext = Array.from(container.querySelectorAll('button')).find(
      (item) => item.textContent === '复制含定位 ID 的诊断',
    )!
    await act(async () => withContext.click())
    expect(copyDiagnosticsWithContext).toHaveBeenCalledOnce()
    expect(container.querySelector('[role="status"]')?.textContent).toBe(
      '已复制含定位 ID 的本机诊断；请检查后分享',
    )
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
it.each(['session-upload', 'composer-attachment-upload'])(
  'serializes %s and drops a late result after new task',
  async (pickerId) => {
    const snapshot: OfficeAgentSnapshot = {
      assistantText: 'Ready',
      activity: '',
      busy: false,
      applying: false,
      status: 'done',
      retryable: true,
      error: 'tool_failed',
      errorMessage: 'Retry',
      timeline: [{ id: 'a', kind: 'assistant', text: 'Ready' }],
    }
    const session = {
      snapshot: () => snapshot,
      subscribe: () => () => {},
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
    let finish!: () => void
    const upload = vi.fn(
      () =>
        new Promise<void>((resolve) => {
          finish = resolve
        }),
    )
    const ui: OfficeWorkspaceUi = {
      attachments: () => [],
      skills: () => [],
      skillPackagesEnabled: true,
      durableAttachmentsAvailable: () => true,
      upload,
      clear: vi.fn(),
    }
    const container = document.createElement('div')
    document.body.append(container)
    const root = createRoot(container)
    await act(async () =>
      root.render(
        React.createElement(AgentWorkspace, {
          session,
          ui,
          disconnect: vi.fn(),
          host: 'powerpoint',
          initialPanel: 'attachments',
        }),
      ),
    )
    const input = container.querySelector<HTMLInputElement>(`#${pickerId}`)!
    Object.defineProperty(input, 'files', {
      configurable: true,
      value: [new File(['hello'], 'source.txt')],
    })
    await act(async () => input.dispatchEvent(new Event('change', { bubbles: true })))
    expect(upload).toHaveBeenCalledOnce()
    expect(input.value).toBe('')
    expect(input.disabled).toBe(true)
    expect(container.textContent).toContain('正在上传')
    expect(container.textContent).toContain('退出登录不会删除')
    const retry = Array.from(container.querySelectorAll('button')).find(
      (b) => b.textContent === 'Retry',
    )!
    expect(retry.disabled).toBe(true)
    const newTask = Array.from(container.querySelectorAll('button')).find(
      (b) => b.textContent === '新对话',
    )!
    await act(async () => newTask.click())
    await act(async () => finish())
    expect(container.textContent).not.toContain('source.txt 已保存')
    await act(async () => root.unmount())
    container.remove()
  },
)

it('refreshes persisted acquisition history after webpage failure and explicit retry without losing durable files', async () => {
  const snapshot = {
    assistantText: '',
    activity: '',
    busy: false,
    applying: false,
    status: 'done',
    retryable: false,
    timeline: [],
  } as OfficeAgentSnapshot
  const session = {
    snapshot: () => snapshot,
    subscribe: () => () => undefined,
  } as unknown as OfficeAgentSession
  const id = 'a'.repeat(64)
  const read = vi.fn(async () => ({
    version: 1 as const,
    scope: 'remote_material_acquisition' as const,
    documentId: 'doc1',
    revision: 0,
    totalAttempts: 0,
    records: [],
  }))
  const acquire = vi
    .fn()
    .mockRejectedValueOnce(new Error('presentation_remote_webpage_unavailable'))
    .mockResolvedValueOnce(undefined)
  const ui: OfficeWorkspaceUi = {
    attachments: () => [],
    skills: () => [],
    skillPackagesEnabled: true,
    upload: vi.fn(),
    clear: vi.fn(),
    durableAttachmentsAvailable: () => true,
    webpagesAvailable: () => true,
    readPresentationAcquisitionHistory: read,
    importPresentationWebpageUrl: acquire,
    listDurableAttachments: async () => [
      {
        attachmentId: id,
        sha256: id,
        name: 'kept.txt',
        sizeBytes: 1,
        receivedBytes: 1,
        status: 'ready',
        kind: 'text',
        totalChars: 1,
      },
    ],
  }
  const container = document.createElement('div')
  document.body.append(container)
  const root = createRoot(container)
  try {
    await act(async () =>
      root.render(
        React.createElement(AgentWorkspace, {
          session,
          ui,
          disconnect: vi.fn(),
          host: 'powerpoint',
          initialPanel: 'attachments',
        }),
      ),
    )
    expect(read).toHaveBeenCalledTimes(1)
    const input = container.querySelector<HTMLInputElement>('#presentation-webpage-url')!
    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(
        input,
        'https://example.com/page',
      )
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })
    await act(async () =>
      input
        .closest('form')!
        .dispatchEvent(new Event('submit', { bubbles: true, cancelable: true })),
    )
    expect(acquire).toHaveBeenCalledTimes(1)
    expect(read).toHaveBeenCalledTimes(2)
    expect(container.textContent).toContain('kept.txt')
    await act(async () =>
      input
        .closest('form')!
        .dispatchEvent(new Event('submit', { bubbles: true, cancelable: true })),
    )
    expect(acquire).toHaveBeenCalledTimes(2)
    expect(read).toHaveBeenCalledTimes(3)
    expect(container.textContent).toContain('kept.txt')
  } finally {
    await act(async () => root.unmount())
    container.remove()
  }
})

it('renders independent research when no presentation project or production exists', async () => {
  const { createPresentationResearchController } =
    await import('../src/agent/presentation-research.js')
  const { researchRecord, researchSummary } = await import('./presentation-research-fixture.js')
  const research = createPresentationResearchController({
    available: () => true,
    documentId: async () => 'doc',
    lastProject: () => 'research',
    request: async (body) =>
      new Response(
        JSON.stringify(
          (body as { operation: string }).operation === 'research_list'
            ? researchSummary()
            : { version: 1, available: true },
        ),
      ),
    executeTool: async (call) => ({
      output: JSON.stringify(
        call.name === 'list_research_ledgers' ? researchSummary() : researchRecord(),
      ),
      mutated: false,
      summary: 'read',
    }),
  })
  const snapshot = {
    assistantText: '',
    activity: '',
    busy: false,
    applying: false,
    status: 'done',
    retryable: false,
    timeline: [],
  } as OfficeAgentSnapshot
  const session = {
    snapshot: () => snapshot,
    subscribe: () => () => undefined,
  } as unknown as OfficeAgentSession
  const ui: OfficeWorkspaceUi = {
    attachments: () => [],
    skills: () => [],
    skillPackagesEnabled: true,
    upload: vi.fn(),
    clear: vi.fn(),
    research,
  }
  const container = document.createElement('div')
  const root = createRoot(container)
  try {
    await act(async () =>
      root.render(
        React.createElement(AgentWorkspace, {
          session,
          ui,
          disconnect: vi.fn(),
          host: 'powerpoint',
        }),
      ),
    )
    expect(container.textContent).toContain('资料研究账本')
    expect(container.textContent).not.toContain('演示文稿项目')
    await act(async () => research.read('ledger1'))
    expect(container.textContent).toContain('销售增长')
    expect(container.textContent).toContain('销售下降')
  } finally {
    await act(async () => root.unmount())
  }
})

it('shows the actual image step and expandable safe PC failure detail', async () => {
  let observe: ((event: OfficeToolActivity) => void) | undefined
  const session = createOfficeAgentSession({
    transport: {
      stream: () => ({ cancel: vi.fn() }),
      setToolActivityHandler: (next) => {
        observe = next
      },
    },
    skill: {
      id: 'test',
      systemPrompt: '',
      tools: [{ name: 'insert_web_image', description: 'image', inputSchema: { type: 'object' } }],
      executeTool: vi.fn(),
    },
    proposals: createStructuredProposalController(),
  })
  session.send('Insert image')
  await Promise.resolve()
  const base = { callId: 'call_image123', toolName: 'insert_web_image', startedAt: Date.now() }
  observe!({ ...base, state: 'running' })
  observe!({ ...base, state: 'error', summary: 'image_fetch_unavailable' })
  const container = document.createElement('div')
  document.body.append(container)
  const root = createRoot(container)
  await act(async () =>
    root.render(
      React.createElement(AgentWorkspace, {
        session,
        ui: {
          attachments: () => [],
          skills: () => [],
          skillPackagesEnabled: true,
          upload: vi.fn(),
          clear: vi.fn(),
        },
        disconnect: vi.fn(),
        host: 'powerpoint',
      }),
    ),
  )
  expect(container.textContent).toContain('插入网络图片未完成')
  expect(container.textContent).not.toContain('准备修改')
  expect(container.textContent).not.toContain('Office tool failed')
  await act(async () =>
    container.querySelector<HTMLButtonElement>('.ai-work-group-summary')!.click(),
  )
  await act(async () => container.querySelector<HTMLButtonElement>('.ai-step-title')!.click())
  expect(container.querySelector('.ai-step-detail')?.textContent).toBe(
    '图片暂时无法获取（image_fetch_unavailable）',
  )
  await act(async () => root.unmount())
  session.dispose()
  container.remove()
})

it('shows the runtime selected by WisWork PC without exposing a second mode switch', () => {
  const standard = workspaceMarkup({}, undefined, 'powerpoint', undefined, 'standard')
  const enhanced = workspaceMarkup({}, undefined, 'powerpoint', undefined, 'enhanced')

  expect(standard).toContain('标准模式')
  expect(enhanced).toContain('增强模式')
  expect(enhanced).toContain('由 WisWork PC 管理')
  expect(enhanced).not.toMatch(/切换到|启用增强|mode-switch/)
})

it('shows PowerPoint clarification as one model-authored question at a time', () => {
  const html = workspaceMarkup(
    {
      questionnaire: Object.freeze([
        Object.freeze({ id: 'audience', label: '这份 PPT 面向谁？', options: ['客户', '团队'] }),
        Object.freeze({ id: 'style', label: '希望什么风格？', options: ['简洁', '杂志感'] }),
      ]),
    },
    undefined,
    'powerpoint',
  )

  expect(html).toContain('这份 PPT 面向谁？')
  expect(html).not.toContain('希望什么风格？')
  expect(html).toContain('下一题')
})

it('collapses an applied PowerPoint proposal into a concise verified result', () => {
  const structured = {
    id: 'ppt-applied',
    operation: 'edit_slide',
    title: '设置 LLM 介绍 PPT 封面',
    impact: {
      host: 'powerpoint',
      targets: ['256#3943334991', 'slide-1'],
      count: 4,
    },
    preview: { slideIndex: 0 },
    fingerprint: 'private',
    before: '',
    after: '',
  }
  const html = workspaceMarkup(
    {
      proposal: undefined,
      timeline: Object.freeze([
        Object.freeze({
          id: 'result',
          kind: 'proposal' as const,
          proposal: structured,
          state: 'applied' as const,
        }),
      ]),
    },
    undefined,
    'powerpoint',
  )

  expect(html).toContain('设置 LLM 介绍 PPT 封面')
  expect(html).toContain('已更新 4 项')
  expect(html).not.toMatch(
    /CHANGE APPLIED|Review exact impact|Before|After|empty document|Version/i,
  )
  expect(html).not.toContain('256#3943334991')
})

it('announces when the current relay session could not be remembered', () => {
  const notice =
    'Connected, but this Office installation was not remembered. Pair again after reconnecting.'
  const html = workspaceMarkup({}, undefined, 'word', notice)
  expect(html).toContain('role="status"')
  expect(html).toContain(notice)
})
