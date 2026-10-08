import React from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { expect, it, vi } from 'vitest'
import { AgentWorkspace, createOfficeWorkspaceUi, type OfficeWorkspaceUi } from '../src/App.js'
import type { OfficeAgentSession, OfficeAgentSnapshot } from '../src/agent/use-office-agent.js'
import type { OfficeHostRuntime } from '../src/agent/host-runtime.js'
import type { PresentationTeamController } from '../src/agent/presentation-team-controller.js'

it('binds the actual runtime team controller into workspace UI', () => {
  const team = { snapshot: vi.fn() } as unknown as PresentationTeamController
  const runtime = {
    team,
    vfs: { list: () => [] },
    skills: { names: () => [] },
  } as unknown as OfficeHostRuntime
  expect(createOfficeWorkspaceUi(runtime).team).toBe(team)
})
it('renders the team workbench in PowerPoint and keeps host-specific data out of other hosts', () => {
  const state: OfficeAgentSnapshot = {
    assistantText: '',
    activity: '',
    busy: false,
    applying: false,
    status: 'idle',
    retryable: false,
    timeline: [],
  }
  const session: OfficeAgentSession = {
    snapshot: () => state,
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
  }
  const snapshot = { available: false, phase: 'idle' as const }
  const team: PresentationTeamController = {
    snapshot: () => snapshot,
    subscribe: () => () => {},
    refresh: vi.fn(),
    clear: vi.fn(),
    create: vi.fn(),
    publish: vi.fn(),
    setMember: vi.fn(),
    revokeMember: vi.fn(),
    addComment: vi.fn(),
    resolveComment: vi.fn(),
  }
  const ui: OfficeWorkspaceUi = {
    team,
    attachments: () => [],
    skills: () => [],
    skillPackagesEnabled: false,
    upload: vi.fn(),
    clear: vi.fn(),
  }
  const render = (host: 'powerpoint' | 'word') =>
    renderToStaticMarkup(
      React.createElement(AgentWorkspace, { session, ui, disconnect: vi.fn(), host }),
    )
  expect(render('powerpoint')).toContain('团队')
  expect(render('word')).not.toContain('团队协作')
})
