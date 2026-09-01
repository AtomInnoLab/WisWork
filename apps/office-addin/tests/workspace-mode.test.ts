import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { deploymentConfig, officeWorkspaceMode } from '../build-config.js'
import { officePresentationVerificationFlags } from '../src/agent/presentation-flags.js'
import {
  AgentWorkspace,
  LegacyAgentWorkspace,
  officePresentationText,
  workspaceComponentForMode,
} from '../src/App.js'
import { LANGS, htmlLang, presentationVerificationStrings } from '@wiswork/i18n'

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

describe('Office presentation verification rollout flags', () => {
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
  it('uses safe defaults and independent exact rollback switches', () => {
    expect(officePresentationVerificationFlags({})).toEqual({
      planning: true,
      verifiedCompletion: true,
      visualReview: true,
      autoCorrection: false,
    })
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
})
