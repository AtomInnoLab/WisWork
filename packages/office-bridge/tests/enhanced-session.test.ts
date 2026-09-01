import { describe, expect, it } from 'vitest'

import {
  parseOfficeEnhancedSessionStatement,
  rawOfficeCapabilities,
  selectOfficeAgentRuntime,
} from '../src/index'

const statement = {
  version: 1,
  runtime_mode: 'enhanced',
  runtime_instance: 'runtime_0123456789abcdef',
  component_version: '0.147.0',
  host: 'office-word',
  raw_office: false,
  expires_at: 2_000,
  policy_generation: 7,
  session_generation: 3,
} as const

describe('Office Enhanced session statement', () => {
  it('derives host-specific raw sub-capabilities without advertising Excel OOXML', () => {
    expect(rawOfficeCapabilities({ ...statement, raw_office: true })).toEqual({
      rawJs: true,
      rawOoxml: true,
    })
    expect(rawOfficeCapabilities({ ...statement, host: 'office-excel', raw_office: true })).toEqual(
      { rawJs: true, rawOoxml: false },
    )
  })
  it('strictly parses a bounded non-callable statement', () => {
    expect(parseOfficeEnhancedSessionStatement(statement)).toEqual(statement)
    expect(JSON.stringify(statement)).not.toMatch(/token|secret|path|process|credential/i)
  })

  it('rejects unknown fields, malformed identities, unsupported versions, and excess bounds', () => {
    for (const candidate of [
      { ...statement, token: 'nope' },
      { ...statement, version: 2 },
      { ...statement, runtime_instance: '../escape' },
      { ...statement, component_version: 'latest' },
      { ...statement, policy_generation: -1 },
      { ...statement, expires_at: Number.MAX_SAFE_INTEGER + 1 },
    ]) {
      expect(() => parseOfficeEnhancedSessionStatement(candidate)).toThrow(
        'invalid_office_enhanced_statement',
      )
    }
  })

  it('selects Enhanced only for an exact active host/session/generation/version statement', () => {
    expect(
      selectOfficeAgentRuntime(statement, {
        host: 'office-word',
        now: 1_000,
        componentVersion: '0.147.0',
        sessionGeneration: 3,
        minimumPolicyGeneration: 7,
      }),
    ).toEqual({
      mode: 'enhanced',
      rawOfficeAllowed: false,
      runtimeInstance: statement.runtime_instance,
    })

    for (const context of [
      {
        host: 'office-excel',
        now: 1_000,
        componentVersion: '0.147.0',
        sessionGeneration: 3,
        minimumPolicyGeneration: 7,
      },
      {
        host: 'office-word',
        now: 2_000,
        componentVersion: '0.147.0',
        sessionGeneration: 3,
        minimumPolicyGeneration: 7,
      },
      {
        host: 'office-word',
        now: 1_000,
        componentVersion: '0.148.0',
        sessionGeneration: 3,
        minimumPolicyGeneration: 7,
      },
      {
        host: 'office-word',
        now: 1_000,
        componentVersion: '0.147.0',
        sessionGeneration: 4,
        minimumPolicyGeneration: 7,
      },
      {
        host: 'office-word',
        now: 1_000,
        componentVersion: '0.147.0',
        sessionGeneration: 3,
        minimumPolicyGeneration: 8,
      },
    ] as const) {
      expect(selectOfficeAgentRuntime(statement, context)).toEqual({
        mode: 'standard',
        rawOfficeAllowed: false,
      })
    }
  })

  it('keeps Standard default and raw Office as an independent typed gate', () => {
    expect(
      selectOfficeAgentRuntime(undefined, {
        host: 'office-word',
        now: 1_000,
        componentVersion: '0.147.0',
        sessionGeneration: 3,
        minimumPolicyGeneration: 0,
      }),
    ).toEqual({ mode: 'standard', rawOfficeAllowed: false })
    expect(
      selectOfficeAgentRuntime(
        { ...statement, raw_office: true },
        {
          host: 'office-word',
          now: 1_000,
          componentVersion: '0.147.0',
          sessionGeneration: 3,
          minimumPolicyGeneration: 7,
        },
      ),
    ).toMatchObject({ mode: 'enhanced', rawOfficeAllowed: true })
  })
})
