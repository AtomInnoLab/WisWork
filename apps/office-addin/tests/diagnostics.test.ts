import { describe, expect, it, vi } from 'vitest'
import {
  createOfficeDiagnostics,
  officeDiagnosticEnvironment,
} from '../src/diagnostics/office-diagnostics.js'

describe('Office safe diagnostics', () => {
  it('normalizes Office platform and exposes only the active known requirement set', () => {
    const isSetSupported = vi.fn(
      (name: string, version: string) => name === 'WordApi' && version === '1.3',
    )
    expect(
      officeDiagnosticEnvironment('word', {
        Office: { context: { platform: 'Mac', requirements: { isSetSupported } } },
      }),
    ).toEqual({ platform: 'mac', requirementSets: { WordApi: true } })
    expect(officeDiagnosticEnvironment('excel', {})).toEqual({
      platform: 'unknown',
      requirementSets: { ExcelApi: false },
    })
  })
  it('retains a bounded local ring and exports no document or raw error content', () => {
    let id = 0
    const diagnostics = createOfficeDiagnostics({
      host: 'word',
      platform: 'mac',
      build: 'build-123',
      now: () => 1_000 + id,
      randomUUID: () => `00000000-0000-4000-8000-${String(++id).padStart(12, '0')}`,
      requirementSets: { WordApi: true },
    })
    diagnostics.startTrace()
    diagnostics.setTool('write_document')
    for (let index = 0; index < 205; index += 1)
      diagnostics.record({
        phase: 'write',
        error: Object.assign(new Error('document secret'), {
          code: 'InvalidArgument',
          debugInfo: {
            errorLocation: 'Body.insertText',
            statement: 'insert super secret document text',
          },
        }),
        errorCode: 'office_write_failed',
        durationMs: index,
        prohibitedPrompt: 'never retain me',
      } as never)

    const snapshot = diagnostics.snapshot()
    expect(snapshot.events).toHaveLength(200)
    expect(snapshot.events[0]?.duration_ms).toBe(5)
    expect(snapshot.events.at(-1)).toMatchObject({
      host: 'word',
      platform: 'mac',
      build: 'build-123',
      tool: 'write_document',
      phase: 'write',
      outcome: 'failed',
      error_code: 'office_write_failed',
      office_error_code: 'InvalidArgument',
      office_error_name: 'Error',
      office_error_location: 'Body.insertText',
      requirement_sets: { WordApi: true },
    })
    const exported = diagnostics.exportJson()
    expect(new TextEncoder().encode(exported).byteLength).toBeLessThanOrEqual(256 * 1024)
    expect(exported).not.toContain('document secret')
    expect(exported).not.toContain('super secret')
    expect(exported).not.toContain('never retain me')
  })

  it('isolates asynchronous upload failures from staged Office diagnostics', async () => {
    const sent: unknown[] = []
    const diagnostics = createOfficeDiagnostics({
      host: 'word',
      platform: 'unknown',
      build: 'dev',
      remoteEnabled: true,
      send: async (event) => {
        sent.push(event)
        throw new Error('relay secret failure')
      },
      randomUUID: () => '00000000-0000-4000-8000-000000000001',
      now: () => 10,
    })
    diagnostics.startTrace()
    diagnostics.setTool('write_document')
    const stagedError = Object.assign(new Error('document contains payroll'), {
      code: 'InvalidArgument',
      verificationStage: 'content',
      debugInfo: { errorLocation: 'Body.insertOoxml' },
    })
    expect(() =>
      diagnostics.record({
        phase: 'verify',
        errorCode: 'office_verify_failed',
        error: stagedError,
      }),
    ).not.toThrow()
    await vi.waitFor(() =>
      expect(diagnostics.snapshot().events.at(-1)?.error_code).toBe('diagnostic_upload_failed'),
    )
    expect(sent).toHaveLength(1)
    expect(JSON.stringify(sent)).not.toContain('payroll')
    expect(diagnostics.snapshot().events[0]).toMatchObject({
      verification_stage: 'content',
      office_error_code: 'InvalidArgument',
      office_error_name: 'Error',
      office_error_location: 'Body.insertOoxml',
    })
    expect(diagnostics.snapshot().events.at(-1)).toMatchObject({
      phase: 'transport',
      error_code: 'diagnostic_upload_failed',
    })
    expect(diagnostics.snapshot().events.at(-1)).not.toHaveProperty('verification_stage')
    expect(diagnostics.snapshot().events.at(-1)).not.toHaveProperty('office_error_code')
    expect(diagnostics.snapshot().events.at(-1)).not.toHaveProperty('office_error_name')
    expect(diagnostics.snapshot().events.at(-1)).not.toHaveProperty('office_error_location')
  })

  it('isolates synchronous upload failures from staged Office diagnostics', () => {
    const diagnostics = createOfficeDiagnostics({
      host: 'word',
      build: 'dev',
      remoteEnabled: true,
      send: () => {
        throw new Error('relay secret failure')
      },
      randomUUID: () => '00000000-0000-4000-8000-000000000001',
      now: () => 10,
    })
    const stagedError = Object.assign(new Error('secret document content'), {
      code: 'InvalidArgument',
      verificationStage: 'boundary',
      debugInfo: { errorLocation: 'Body.insertOoxml' },
    })

    expect(() =>
      diagnostics.record({
        phase: 'verify',
        errorCode: 'office_verify_failed',
        error: stagedError,
      }),
    ).not.toThrow()

    expect(diagnostics.snapshot().events[0]).toMatchObject({
      verification_stage: 'boundary',
      office_error_code: 'InvalidArgument',
      office_error_name: 'Error',
      office_error_location: 'Body.insertOoxml',
    })
    expect(diagnostics.snapshot().events.at(-1)).toMatchObject({
      phase: 'transport',
      error_code: 'diagnostic_upload_failed',
    })
    expect(diagnostics.snapshot().events.at(-1)).not.toHaveProperty('verification_stage')
    expect(diagnostics.snapshot().events.at(-1)).not.toHaveProperty('office_error_code')
    expect(diagnostics.snapshot().events.at(-1)).not.toHaveProperty('office_error_name')
    expect(diagnostics.snapshot().events.at(-1)).not.toHaveProperty('office_error_location')
  })

  it('clears traces and records stable unsupported and cancellation outcomes', () => {
    const diagnostics = createOfficeDiagnostics({ host: 'powerpoint', build: 'dev' })
    diagnostics.startTrace()
    diagnostics.record({ phase: 'tool', errorCode: 'office_api_unsupported' })
    diagnostics.record({ phase: 'transport', errorCode: 'cancelled' })
    expect(diagnostics.snapshot().events.map((event) => event.outcome)).toEqual([
      'unsupported',
      'cancelled',
    ])
    diagnostics.clear()
    expect(diagnostics.snapshot().events).toEqual([])
    expect(diagnostics.snapshot().trace_id).toBeUndefined()
  })

  it('preserves invalid tool input locally for actionable model-contract diagnosis', () => {
    const diagnostics = createOfficeDiagnostics({ host: 'powerpoint', build: 'dev' })
    diagnostics.setTool('edit_slide_xml')
    diagnostics.record({ phase: 'tool', errorCode: 'invalid_tool_input' })
    expect(diagnostics.snapshot().events[0]).toMatchObject({
      tool: 'edit_slide_xml',
      error_code: 'invalid_tool_input',
    })
  })

  it('ignores hostile Office error accessors without affecting the failure record', () => {
    const diagnostics = createOfficeDiagnostics({ host: 'word', build: 'dev' })
    const hostile = Object.defineProperty({}, 'debugInfo', {
      get() {
        throw new Error('secret getter')
      },
    })
    expect(() =>
      diagnostics.record({
        phase: 'write',
        errorCode: 'office_write_failed',
        error: hostile,
      }),
    ).not.toThrow()
    expect(diagnostics.snapshot().events[0]).not.toHaveProperty('office_error_location')
  })

  it('extracts only allowlisted identifiers from a shallow wrapped Office error', () => {
    const diagnostics = createOfficeDiagnostics({ host: 'powerpoint', build: 'dev' })
    const cause = Object.assign(new Error('secret document content'), {
      code: 'InvalidArgument',
      debugInfo: { errorLocation: 'ShapeCollection.add' },
    })
    const wrapped = new Error('office_write_failed', { cause })
    diagnostics.record({ phase: 'write', errorCode: 'office_write_failed', error: wrapped })

    expect(diagnostics.snapshot().events[0]).toMatchObject({
      office_error_code: 'InvalidArgument',
      office_error_name: 'Error',
      office_error_location: 'ShapeCollection.add',
    })
    expect(diagnostics.exportJson()).not.toContain('secret document content')
  })

  it('retains an allowlisted verification stage from a shallow error cause', () => {
    const diagnostics = createOfficeDiagnostics({ host: 'word', build: 'dev' })
    const staged = Object.assign(new Error('secret document content'), {
      verificationStage: 'body_shape',
    })
    const wrapped = new Error('office_verify_failed', {
      cause: new Error('verification failed', { cause: staged }),
    })
    diagnostics.record({ phase: 'verify', errorCode: 'office_verify_failed', error: wrapped })

    expect(diagnostics.snapshot().events[0]).toMatchObject({
      error_code: 'office_verify_failed',
      verification_stage: 'body_shape',
    })
    const exported = diagnostics.exportJson()
    expect(exported).toContain('"verification_stage": "body_shape"')
    expect(exported).not.toContain('secret document content')
  })

  it('discards invalid and hostile verification stages without exposing content', () => {
    const diagnostics = createOfficeDiagnostics({ host: 'word', build: 'dev' })
    diagnostics.record({
      phase: 'verify',
      errorCode: 'office_verify_failed',
      error: { verificationStage: 'paragraph contains secret document text' },
    })
    const hostile = Object.defineProperty({}, 'verificationStage', {
      get() {
        throw new Error('secret getter content')
      },
    })
    expect(() =>
      diagnostics.record({
        phase: 'verify',
        errorCode: 'office_verify_failed',
        error: hostile,
      }),
    ).not.toThrow()

    expect(diagnostics.snapshot().events).toHaveLength(2)
    expect(diagnostics.snapshot().events[0]).not.toHaveProperty('verification_stage')
    expect(diagnostics.snapshot().events[1]).not.toHaveProperty('verification_stage')
    const exported = diagnostics.exportJson()
    expect(exported).not.toContain('paragraph contains secret document text')
    expect(exported).not.toContain('secret getter content')
  })

  it('retains only allowlisted content-free Word recovery stages', () => {
    const diagnostics = createOfficeDiagnostics({ host: 'word', build: 'dev' })
    diagnostics.record({
      phase: 'recovery',
      errorCode: 'office_recovery_failed:word_content',
    })
    diagnostics.record({
      phase: 'recovery',
      errorCode: 'office_recovery_failed:word_secret-document-text',
    })

    expect(diagnostics.snapshot().events.map((event) => event.error_code)).toEqual([
      'office_recovery_failed:word_content',
      'office_write_failed',
    ])
  })
})
