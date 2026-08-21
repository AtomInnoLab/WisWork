import React from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it, vi } from 'vitest'
import { AgentWorkspace, composerKeyAction, type WorkspacePanelName } from '../src/App.js'
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
  const runtime = {
    vfs: { list: () => ['/home/user/source.docx'] },
    skills: { list: () => [{ name: 'Editorial review' }] },
  } as unknown as OfficeHostRuntime
  return renderToStaticMarkup(
    React.createElement(AgentWorkspace, {
      session,
      runtime,
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

  it('exposes bounded attachment and skill management panels without permanent vertical chrome', () => {
    const files = workspaceMarkup({}, 'attachments')
    expect(files).toContain('role="dialog"')
    expect(files).toContain('Session attachments')
    expect(files).toContain('source.docx')
    expect(files).not.toContain('Editorial review')

    const skills = workspaceMarkup({}, 'skills')
    expect(skills).toContain('Installed skills')
    expect(skills).toContain('Editorial review')
  })

  it('renders distinct accessible working, applying, stop, and retry states', () => {
    const working = workspaceMarkup({ busy: true, status: 'working' })
    expect(working).toContain('Agent is working')
    expect(working).toContain('>Stop<')
    expect(working).toContain('aria-busy="true"')

    const applying = workspaceMarkup({ applying: true })
    expect(applying).toContain('Applying approved change')
    expect(applying).toContain('Applying…')

    const failed = workspaceMarkup({ status: 'error', error: 'office_write_failed' })
    expect(failed).toContain('>Retry<')
    expect(failed).toContain('role="alert"')
  })

  it('sends on Enter, preserves multiline input on Shift+Enter, and ignores composition', () => {
    expect(composerKeyAction({ key: 'Enter', shiftKey: false, isComposing: false })).toBe('send')
    expect(composerKeyAction({ key: 'Enter', shiftKey: true, isComposing: false })).toBe('newline')
    expect(composerKeyAction({ key: 'Enter', shiftKey: false, isComposing: true })).toBe('newline')
    expect(composerKeyAction({ key: 'Escape', shiftKey: false, isComposing: false })).toBe('none')
  })
})
