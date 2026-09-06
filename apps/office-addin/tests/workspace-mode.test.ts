import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { deploymentConfig, officeWorkspaceMode } from '../build-config.js'
import { officePresentationVerificationFlags } from '../src/agent/presentation-flags.js'
import {
  AgentWorkspace,
  LegacyAgentWorkspace,
  officeRuntimeModeForTaskpane,
  officePresentationText,
  workspaceComponentForMode,
} from '../src/App.js'
import { LANGS, htmlLang, presentationVerificationStrings } from '@wiswork/i18n'

describe('Office workspace rollback flag', () => {
  it('derives the Taskpane mode only from the current PC session statement', () => {
    const enhanced = {
      version: 1 as const,
      runtime_mode: 'enhanced' as const,
      runtime_instance: 'runtime_123456789',
      component_version: '0.147.0' as const,
      host: 'office-powerpoint' as const,
      raw_office: false,
      expires_at: 20_000,
      policy_generation: 1,
      session_generation: 2,
    }
    expect(
      officeRuntimeModeForTaskpane('powerpoint', { status: 'connected', enhanced }, 10_000),
    ).toBe('enhanced')
    expect(officeRuntimeModeForTaskpane('word', { status: 'connected', enhanced }, 10_000)).toBe(
      'standard',
    )
    expect(
      officeRuntimeModeForTaskpane('powerpoint', { status: 'connected', enhanced }, 20_000),
    ).toBe('standard')
    expect(
      officeRuntimeModeForTaskpane('powerpoint', { status: 'offline', enhanced }, 10_000),
    ).toBe('standard')
  })

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

describe('Office presentation verification rollout flags', () => {
  const defaults = {
    planning: true,
    verifiedCompletion: true,
    visualReview: true,
    autoCorrection: true,
  }
  const flagCases = [
    ['PLANNING', 'planning'],
    ['VERIFIED_COMPLETION', 'verifiedCompletion'],
    ['VISUAL_REVIEW', 'visualReview'],
    ['AUTO_CORRECTION', 'autoCorrection'],
  ] as const

  it('keeps the Node-evaluated Vite build config free of workspace runtime imports', () => {
    const source = readFileSync(resolve(import.meta.dirname, '../build-config.ts'), 'utf8')
    expect(source).not.toMatch(/from\s+['"]@wiswork\//)
  })

  it('uses the production locale adapter for every presentation state in all locales', () => {
    for (const lang of LANGS)
      for (const key of Object.keys(presentationVerificationStrings.zh))
        expect(
          officePresentationText(
            htmlLang(lang),
            key as keyof typeof presentationVerificationStrings.zh,
          ),
        ).toBe(
          presentationVerificationStrings[lang][
            key as keyof typeof presentationVerificationStrings.zh
          ],
        )
  })
  it('enables the bounded verification and correction loop by default', () => {
    expect(officePresentationVerificationFlags({})).toEqual(defaults)
  })

  it.each(flagCases)('rolls back only %s with the exact zero flag', (name, property) => {
    const key = `VITE_WISWORK_PRESENTATION_${name}`
    expect(officePresentationVerificationFlags({ [key]: '0' })).toEqual({
      ...defaults,
      [property]: false,
    })
    for (const value of [undefined, '', '1'])
      expect(officePresentationVerificationFlags({ [key]: value })).toEqual(defaults)
  })

  it.each(flagCases)('rejects non-exact values for %s', (name) => {
    for (const value of ['false', 'true', '2', ' 0', '1 '])
      expect(() =>
        officePresentationVerificationFlags({ [`VITE_WISWORK_PRESENTATION_${name}`]: value }),
      ).toThrow('invalid_presentation_verification_flags')
  })

  it('keeps correction independent from the other verification switches', () => {
    expect(
      officePresentationVerificationFlags({
        VITE_WISWORK_PRESENTATION_PLANNING: '0',
        VITE_WISWORK_PRESENTATION_VERIFIED_COMPLETION: '0',
        VITE_WISWORK_PRESENTATION_VISUAL_REVIEW: '0',
        VITE_WISWORK_PRESENTATION_AUTO_CORRECTION: '1',
      }),
    ).toEqual({
      planning: false,
      verifiedCompletion: false,
      visualReview: false,
      autoCorrection: true,
    })
  })

  it('ignores server-only and workspace rollback flags', () => {
    expect(
      officePresentationVerificationFlags({
        WISWORK_PRESENTATION_PLANNING: '0',
        WISWORK_PRESENTATION_VERIFIED_COMPLETION: '0',
        WISWORK_PRESENTATION_VISUAL_REVIEW: '0',
        WISWORK_PRESENTATION_AUTO_CORRECTION: '0',
        VITE_WISWORK_OFFICE_WORKSPACE: '0',
      }),
    ).toEqual(defaults)
  })
})
