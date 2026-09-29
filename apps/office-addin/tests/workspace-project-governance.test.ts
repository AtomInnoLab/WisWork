import React from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { afterEach, expect, it, vi } from 'vitest'
import { AgentWorkspace, createOfficeWorkspaceUi, type OfficeWorkspaceUi } from '../src/App.js'
import { createPresentationProjectGovernanceController } from '../src/agent/presentation-project-governance'
import type { OfficeHostRuntime } from '../src/agent/host-runtime'
import type { OfficeAgentSession, OfficeAgentSnapshot } from '../src/agent/use-office-agent'
afterEach(() => vi.unstubAllEnvs())
it('binds actual runtime governance and defaults the product UI off', () => {
  const governance = createPresentationProjectGovernanceController({
    available: () => false,
    documentId: async () => 'doc',
    currentProjectId: () => 'p',
    request: vi.fn(),
    readAttempt: () => undefined,
    writeAttempt: () => {},
  })
  const runtime = {
    governance,
    vfs: { list: () => [] },
    skills: { names: () => [] },
  } as unknown as OfficeHostRuntime
  expect(createOfficeWorkspaceUi(runtime).governance).toBe(governance)
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
  const ui: OfficeWorkspaceUi = {
    governance,
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
  vi.stubEnv('VITE_WISWORK_PPT_PROJECT_GOVERNANCE_ENABLED', '0')
  expect(render('powerpoint')).not.toContain('本机项目资料治理')
  vi.stubEnv('VITE_WISWORK_PPT_PROJECT_GOVERNANCE_ENABLED', 'true')
  expect(render('powerpoint')).not.toContain('本机项目资料治理')
  vi.stubEnv('VITE_WISWORK_PPT_PROJECT_GOVERNANCE_ENABLED', '1')
  expect(render('powerpoint')).toContain('当前连接不支持本机项目治理')
  expect(render('word')).not.toContain('本机项目资料治理')
})
