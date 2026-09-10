import { describe, expect, it, vi } from 'vitest'
import { createPowerPointSkill } from '../src/skills/powerpoint/powerpoint-skill.js'
import { createOfficeAgentSession } from '../src/agent/use-office-agent.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import { createOfficeDiagnostics } from '../src/diagnostics/office-diagnostics.js'
import { createElevatedOfficeSkill } from '../src/skills/shared/elevated-office-program.js'

describe('PowerPoint remote error diagnostics', () => {
  it('preserves safe host identifiers through skill and remote handler without leaking content', async () => {
    const error = Object.assign(new Error('private slide title https://secret.example'), {
      name: 'RichApi.Error',
      code: 'InvalidArgument',
      debugInfo: { errorLocation: 'Shape.textFrame' },
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({
      adapter: { getPresentationState: vi.fn().mockRejectedValue(error) } as any,
      proposals,
    })
    const diagnostics = createOfficeDiagnostics({ host: 'powerpoint', build: 'test' })
    let handler: any
    const session = createOfficeAgentSession({
      skill,
      proposals,
      transport: { stream: vi.fn() },
      diagnostics,
      remoteTools: { setToolHandler: (value) => (handler = value) },
    })
    try {
      const result = await handler({
        turnId: 'turn_12345678',
        callId: 'call_read123',
        generation: 1,
        toolName: 'get_presentation_state',
        input: {},
        signal: new AbortController().signal,
      })
      expect(result).toEqual({ output: 'office_read_failed', isError: true })
      expect(diagnostics.snapshot().events).toEqual([
        expect.objectContaining({
          error_code: 'office_read_failed',
          office_error_code: 'InvalidArgument',
          office_error_location: 'Shape.textFrame',
        }),
      ])
      for (const value of [
        diagnostics.exportJson(),
        JSON.stringify(session.snapshot()),
        JSON.stringify(result),
      ]) {
        expect(value).not.toContain('private slide title')
        expect(value).not.toContain('secret.example')
      }
    } finally {
      session.dispose()
    }
  })

  it('distinguishes rejected raw input from agent failure before touching the host', async () => {
    const proposals = createStructuredProposalController()
    const adapter = {
      host: 'powerpoint',
      captureAuthority: () => ({
        activeMode: 'enhanced',
        signedIn: true,
        paired: true,
        hostEnabled: true,
        rawOfficeEnabled: true,
        rawOfficeJsEnabled: true,
        rawOfficeOoxmlEnabled: true,
        documentId: 'doc_AAAAAAAAAAAAAAAA',
        sessionId: 'ses_AAAAAAAAAAAAAAAA',
        generation: 1,
        revision: 'rev_AAAAAAAAAAAAAAAA',
      }),
      snapshot: vi.fn(),
    }
    const skill = createElevatedOfficeSkill({
      host: 'powerpoint',
      adapter: adapter as any,
      proposals,
    })
    const diagnostics = createOfficeDiagnostics({ host: 'powerpoint', build: 'test' })
    let handler: any
    const session = createOfficeAgentSession({
      skill,
      proposals,
      diagnostics,
      transport: { stream: vi.fn() },
      remoteTools: { setToolHandler: (value) => (handler = value) },
    })
    try {
      const result = await handler({
        turnId: 'turn_12345678',
        callId: 'call_raw123',
        generation: 1,
        toolName: 'propose_raw_office_edit',
        input: {
          program: {
            version: 1,
            kind: 'office_js_ast',
            operations: [{ call: 'shape.setText', args: { private: 'private slide content' } }],
          },
        },
        signal: new AbortController().signal,
      })
      expect(result).toEqual({ output: 'raw_office_program_invalid', isError: true })
      expect(adapter.snapshot).not.toHaveBeenCalled()
      expect(diagnostics.snapshot().events).toEqual([
        expect.objectContaining({
          error_code: 'invalid_tool_input',
          office_error_code: 'RawOfficeProgramInvalid',
        }),
      ])
      expect(diagnostics.exportJson()).not.toContain('private slide content')
      expect(proposals.pending()).toBeUndefined()
    } finally {
      session.dispose()
    }
  })
})
