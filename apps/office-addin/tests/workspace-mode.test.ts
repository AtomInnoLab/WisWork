import { describe, expect, it } from 'vitest'
import { deploymentConfig, officeWorkspaceMode } from '../build-config.js'
import { AgentWorkspace, LegacyAgentWorkspace, workspaceComponentForMode } from '../src/App.js'

describe('Office workspace rollback flag', () => {
  it('defaults to the new workspace and accepts only the exact independent legacy flag', () => {
    expect(officeWorkspaceMode({})).toBe('workspace')
    expect(officeWorkspaceMode({ VITE_WISWORK_OFFICE_WORKSPACE: '1' })).toBe('workspace')
    expect(officeWorkspaceMode({ VITE_WISWORK_OFFICE_WORKSPACE: '0' })).toBe('legacy')
    expect(() => officeWorkspaceMode({ VITE_WISWORK_OFFICE_WORKSPACE: 'false' })).toThrow(
      'invalid_office_workspace_mode',
    )
    expect(
      deploymentConfig({
        VITE_WISWORK_ADDIN_ORIGIN: 'https://office.example',
        VITE_WISWORK_OFFICE_WORKSPACE: 'false',
      }),
    ).toBeUndefined()
    expect(workspaceComponentForMode(officeWorkspaceMode({}))).toBe(AgentWorkspace)
    expect(
      workspaceComponentForMode(officeWorkspaceMode({ VITE_WISWORK_OFFICE_WORKSPACE: '0' })),
    ).toBe(LegacyAgentWorkspace)
  })
})
