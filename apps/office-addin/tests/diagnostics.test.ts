import { describe, expect, it, vi } from 'vitest'
import { readFileSync } from 'node:fs'
import {
  createOfficeDiagnostics,
  officeDiagnosticEnvironment,
} from '../src/diagnostics/office-diagnostics.js'

describe('Office safe diagnostics', () => {
  it('keeps the frozen P0-06 contract probe out of local and remote diagnostics', () => {
    const contract = readFileSync(
      new URL(
        '../../../docs/product/ppt-benchmark-materials/PPT-P0-06/synthetic-nda.txt',
        import.meta.url,
      ),
      'utf8',
    )
    const probe = /^测试追踪码：([^\r\n]+)$/m.exec(contract)?.[1]
    expect(probe).toMatch(/^PRIVATE-SIM-P0-06-/)
    const sent: unknown[] = []
    const diagnostics = createOfficeDiagnostics({
      host: 'powerpoint',
      build: 'test',
      remoteEnabled: true,
      send: (event) => {
        sent.push(event)
      },
    })
    diagnostics.startTrace()
    diagnostics.setTool('read_presentation_attachment', {
      page_id: probe,
      extra: contract,
    } as never)
    diagnostics.record({
      phase: 'tool',
      errorCode: 'office_read_failed',
      error: Object.assign(new Error(contract), {
        debugInfo: { statement: contract, errorLocation: 'Body.insertText' },
      }),
      prohibitedPrompt: contract,
    } as never)
    expect(diagnostics.snapshot().events[0]?.presentation_context?.page_id).toBe(probe)
    expect(diagnostics.exportJson()).not.toContain(probe)
    expect(diagnostics.exportJson({ includeLocalContext: true })).toContain(probe)
    expect(JSON.stringify(sent)).not.toContain(probe)
    expect(sent).toHaveLength(1)
  })
  it('keeps validated presentation identifiers in local export and strips them from remote events', () => {
    const sent: unknown[] = []
    const diagnostics = createOfficeDiagnostics({
      host: 'powerpoint',
      build: 'build-123',
      localDocumentId: 'document-1',
      remoteEnabled: true,
      send: (event) => {
        sent.push(event)
      },
      randomUUID: () => '00000000-0000-4000-8000-000000000001',
    })
    diagnostics.startTrace()
    diagnostics.setTool('run_presentation_production', {
      document_id: 'spoofed-document',
      project_id: 'project-1',
      request_id: 'run-1',
      page_id: 'page-3',
      tool_call_id: 'call-1',
      extra: 'private brief',
    } as never)
    expect(diagnostics.record({ phase: 'tool', errorCode: 'office_write_failed' })).toMatchObject({
      presentation_stage: 'production',
      presentation_context: {
        document_id: 'document-1',
        project_id: 'project-1',
        request_id: 'run-1',
        page_id: 'page-3',
        tool_call_id: 'call-1',
      },
    })
    expect(diagnostics.exportJson()).not.toContain('presentation_context')
    expect(diagnostics.exportJson({ includeLocalContext: true })).toContain('"page_id": "page-3"')
    expect(JSON.stringify(sent)).not.toContain('presentation_context')
    expect(JSON.stringify(sent)).not.toContain('document-1')
    expect(JSON.stringify(sent)).not.toContain('private brief')
    expect(sent[0]).toMatchObject({ presentation_stage: 'production' })
    diagnostics.setTool('read_document', { page_id: 'secret page title' })
    expect(diagnostics.record({ phase: 'tool', errorCode: 'office_read_failed' })).toMatchObject({
      presentation_context: { document_id: 'document-1' },
    })
    diagnostics.startTrace()
    expect(diagnostics.record({ phase: 'run', errorCode: 'agent_run_completed' })).toMatchObject({
      presentation_context: { document_id: 'document-1' },
    })
    expect(diagnostics.snapshot().events.at(-1)).not.toHaveProperty('presentation_stage')
  })
  it('records a bounded run completion without document content', () => {
    const sent: unknown[] = []
    const diagnostics = createOfficeDiagnostics({
      host: 'powerpoint',
      build: 'build-123',
      remoteEnabled: true,
      send: (event) => {
        sent.push(event)
      },
      randomUUID: () => '00000000-0000-4000-8000-000000000001',
    })
    diagnostics.startTrace()
    diagnostics.setTool('agent_run')
    const event = diagnostics.record({
      phase: 'run',
      errorCode: 'agent_run_completed',
      durationMs: 1234,
    })
    expect(event).toMatchObject({
      tool: 'agent_run',
      phase: 'run',
      outcome: 'passed',
      error_code: 'agent_run_completed',
      duration_ms: 1234,
    })
    expect(sent).toHaveLength(1)
    expect(JSON.stringify(sent)).not.toContain('document')
  })
  it('binds each local event to the current Relay session and clears it after revocation', () => {
    const sent: unknown[] = []
    let sessionId: string | undefined = 'relay-1'
    const diagnostics = createOfficeDiagnostics({
      host: 'powerpoint',
      build: 'test',
      localDocumentId: 'document-1',
      localSessionId: () => sessionId,
      remoteEnabled: true,
      send: (event) => {
        sent.push(event)
      },
    })
    diagnostics.startTrace()
    diagnostics.setTool('run_presentation_production', { session_id: 'spoofed-session' })
    expect(diagnostics.record({ phase: 'tool', errorCode: 'office_write_failed' })).toMatchObject({
      presentation_context: { document_id: 'document-1', session_id: 'relay-1' },
    })
    sessionId = undefined
    expect(diagnostics.record({ phase: 'transport', errorCode: 'network_error' })).toMatchObject({
      presentation_context: { document_id: 'document-1' },
    })
    expect(diagnostics.snapshot().events.at(-1)?.presentation_context).not.toHaveProperty(
      'session_id',
    )
    expect(JSON.stringify(sent)).not.toContain('relay-1')
  })
  it('samples remote events per trace while retaining every local event', () => {
    const sent: string[] = []
    let sequence = 0
    const diagnostics = createOfficeDiagnostics({
      host: 'powerpoint',
      build: 'build-123',
      remoteEnabled: true,
      remoteSamplePercent: 25,
      send: (event) => {
        sent.push(event.trace_id)
      },
      randomUUID: () => `00000000-0000-4000-8000-${String(++sequence).padStart(12, '0')}`,
    })
    for (let index = 0; index < 100; index += 1) {
      diagnostics.startTrace()
      diagnostics.record({ phase: 'write', errorCode: 'office_write_failed' })
      diagnostics.record({ phase: 'verify', errorCode: 'office_verify_failed' })
    }
    expect(diagnostics.snapshot().events).toHaveLength(200)
    expect(sent.length).toBeGreaterThan(0)
    expect(sent.length).toBeLessThan(200)
    expect(sent.length % 2).toBe(0)
    for (let index = 0; index < sent.length; index += 2) expect(sent[index]).toBe(sent[index + 1])
    const sampledCount = sent.length
    const off = createOfficeDiagnostics({
      host: 'powerpoint',
      build: 'build-123',
      remoteEnabled: true,
      remoteSamplePercent: 0,
      send: (event) => {
        sent.push(event.trace_id)
      },
      randomUUID: () => '00000000-0000-4000-8000-000000000001',
    })
    off.record({ phase: 'write', errorCode: 'office_write_failed' })
    expect(off.snapshot().events).toHaveLength(1)
    expect(sent).toHaveLength(sampledCount)
    expect(() =>
      createOfficeDiagnostics({ host: 'powerpoint', build: 'x', remoteSamplePercent: 101 }),
    ).toThrow('invalid_office_diagnostic_sample_percent')
  })
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

  it('ignores an out-of-order upload rejection from an older trace', async () => {
    const rejectors: Array<(error: Error) => void> = []
    let id = 0
    const diagnostics = createOfficeDiagnostics({
      host: 'word',
      build: 'dev',
      remoteEnabled: true,
      send: () =>
        new Promise<void>((_resolve, reject) => {
          rejectors.push(reject)
        }),
      randomUUID: () => `00000000-0000-4000-8000-${String(++id).padStart(12, '0')}`,
      now: () => 10,
    })
    diagnostics.startTrace()
    diagnostics.record({ phase: 'write', errorCode: 'office_write_failed' })
    const currentTrace = diagnostics.startTrace()
    diagnostics.record({ phase: 'verify', errorCode: 'office_verify_failed' })

    rejectors[0]?.(new Error('old trace rejection'))
    await Promise.resolve()
    await Promise.resolve()
    expect(
      diagnostics
        .snapshot()
        .events.filter((event) => event.error_code === 'diagnostic_upload_failed'),
    ).toEqual([])

    rejectors[1]?.(new Error('current trace rejection'))
    await vi.waitFor(() =>
      expect(
        diagnostics
          .snapshot()
          .events.filter((event) => event.error_code === 'diagnostic_upload_failed'),
      ).toHaveLength(1),
    )
    expect(diagnostics.snapshot().events.at(-1)).toMatchObject({
      trace_id: currentTrace,
      phase: 'transport',
      error_code: 'diagnostic_upload_failed',
    })
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

  it.each(['debugInfo', 'code', 'name', 'verificationStage'] as const)(
    'ignores a hostile outer %s accessor while retaining safe nested metadata',
    (property) => {
      const diagnostics = createOfficeDiagnostics({ host: 'word', build: 'dev' })
      const officeError = Object.assign(new Error('secret document content'), {
        name: 'RichApi.Error',
        code: 'InvalidArgument',
        debugInfo: { errorLocation: 'Body.insertOoxml' },
      })
      const staged = Object.assign(
        new Error('word_write_verification_failed', { cause: officeError }),
        { verificationStage: 'boundary' },
      )
      const hostile = Object.defineProperty({ cause: staged }, property, {
        get() {
          throw new Error('secret hostile getter')
        },
      })

      expect(() =>
        diagnostics.record({
          phase: 'verify',
          errorCode: 'office_verify_failed',
          error: hostile,
        }),
      ).not.toThrow()
      expect(diagnostics.snapshot().events[0]).toMatchObject({
        verification_stage: 'boundary',
        office_error_code: 'InvalidArgument',
        office_error_name: 'RichApi.Error',
        office_error_location: 'Body.insertOoxml',
      })
      const exported = diagnostics.exportJson()
      expect(exported).not.toContain('secret document content')
      expect(exported).not.toContain('secret hostile getter')
    },
  )

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

  it('preserves a specialized Office error name through a staged verification cause', () => {
    const diagnostics = createOfficeDiagnostics({ host: 'word', build: 'dev' })
    const officeError = Object.assign(new Error('secret document content'), {
      name: 'RichApi.Error',
      code: 'InvalidArgument',
      debugInfo: { errorLocation: 'Body.insertOoxml' },
    })
    const staged = Object.assign(
      new Error('word_write_verification_failed', { cause: officeError }),
      {
        verificationStage: 'content',
      },
    )
    const wrapped = new Error('office_verify_failed', { cause: staged })

    diagnostics.record({ phase: 'verify', errorCode: 'office_verify_failed', error: wrapped })

    expect(diagnostics.snapshot().events[0]).toMatchObject({
      verification_stage: 'content',
      office_error_code: 'InvalidArgument',
      office_error_name: 'RichApi.Error',
      office_error_location: 'Body.insertOoxml',
    })
    expect(diagnostics.exportJson()).not.toContain('secret document content')
  })

  it('does not replace an outer specialized Office name with a deeper generic name', () => {
    const diagnostics = createOfficeDiagnostics({ host: 'word', build: 'dev' })
    const specialized = Object.assign(
      new Error('office failure', { cause: new Error('wrapped') }),
      {
        name: 'RichApi.Error',
      },
    )

    diagnostics.record({ phase: 'write', errorCode: 'office_write_failed', error: specialized })

    expect(diagnostics.snapshot().events[0]?.office_error_name).toBe('RichApi.Error')
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

const screenshotAttempt = (index = 0) => ({
  version: 1 as const,
  id: `12345678-1234-4234-8234-${index.toString(16).padStart(12, '0')}`,
  documentId: 'document',
  projectId: 'project',
  requestId: 'request',
  artifactDigest: 'a'.repeat(64),
  pageId: 'page',
  hostSlideId: 'host',
  startedAt: '2026-09-29T00:00:00.000Z',
  status: 'started' as const,
})
describe('explicit local persisted screenshot attempt export', () => {
  it('never invokes the provider for default copy or remote events; clear retains durable current window', () => {
    let attempts: unknown[] = [screenshotAttempt()]
    const provider = vi.fn(() => attempts as never)
    const sent: unknown[] = []
    const d = createOfficeDiagnostics({
      host: 'powerpoint',
      build: 'test',
      remoteEnabled: true,
      send: (e) => {
        sent.push(e)
      },
    })
    d.record({ phase: 'tool', errorCode: 'office_read_failed' })
    expect(d.exportJson({ screenshotAttempts: provider })).not.toContain(
      'local_presentation_qa_attempts',
    )
    expect(provider).not.toHaveBeenCalled()
    expect(JSON.stringify(sent)).not.toContain('attempts')
    const section = () =>
      JSON.parse(d.exportJson({ includeLocalContext: true, screenshotAttempts: provider }))
        .local_presentation_qa_attempts
    expect(section()).toEqual({
      scope: 'retained_visible_presentation_task',
      status: 'available',
      attempts,
      record_count: 1,
      unresolved_count: 1,
    })
    attempts = [
      {
        ...screenshotAttempt(),
        status: 'closed',
        finishedAt: screenshotAttempt().startedAt,
        errorCode: 'explicitly_closed',
      },
    ]
    d.clear()
    expect(section()).toMatchObject({ attempts, record_count: 1, unresolved_count: 0 })
    expect(d.snapshot().events).toEqual([])
    attempts = []
    expect(section()).toMatchObject({ attempts: [], record_count: 0, unresolved_count: 0 })
  })
  it.each([
    () => {
      throw Error('secret exception')
    },
    () => null,
    () => Array(1),
    () => [screenshotAttempt(), screenshotAttempt()],
    () => Array.from({ length: 65 }, (_, i) => screenshotAttempt(i)),
    () => [{ ...screenshotAttempt(), png: 'secret' }],
    ...['documentId', 'projectId', 'requestId', 'artifactDigest', 'source'].map((key) => () => [
      screenshotAttempt(),
      {
        ...screenshotAttempt(1),
        [key]:
          key === 'source' ? 'production' : key === 'artifactDigest' ? 'b'.repeat(64) : 'other',
      },
    ]),
    () =>
      Array.from({ length: 64 }, (_, i) => ({
        ...screenshotAttempt(i),
        documentId: '界'.repeat(2048),
      })),
  ])('reports invalid/throwing provider %# as unavailable without raw reasons', (provider) => {
    const d = createOfficeDiagnostics({ host: 'powerpoint', build: 'test' })
    const result = JSON.parse(
      d.exportJson({ includeLocalContext: true, screenshotAttempts: provider as never }),
    )
    expect(result.local_presentation_qa_attempts).toEqual({
      scope: 'retained_visible_presentation_task',
      status: 'unavailable',
    })
    expect(JSON.stringify(result)).not.toContain('secret')
  })
  it('retains a full valid UTF8 attempt snapshot and newest events by explicitly omitting oldest events only', () => {
    let sequence = 0
    const d = createOfficeDiagnostics({
      host: 'powerpoint',
      build: 'test',
      randomUUID: () => 'e' + sequence++,
    })
    d.setTool('t'.repeat(128), {
      project_id: 'p'.repeat(128),
      request_id: 'r'.repeat(128),
      page_id: 'a'.repeat(128),
      tool_call_id: 'c'.repeat(128),
    })
    for (let i = 0; i < 200; i++)
      d.record({ phase: 'tool', errorCode: 'office_read_failed', durationMs: i })
    const attempts = Array.from({ length: 32 }, (_, i) => ({
      ...screenshotAttempt(i),
      documentId: '界'.repeat(1000),
    }))
    const exportOptions = { includeLocalContext: true, screenshotAttempts: () => attempts }
    const raw = d.exportJson(exportOptions),
      result = JSON.parse(raw)
    expect(new TextEncoder().encode(raw).byteLength).toBeLessThanOrEqual(256 * 1024)
    expect(result.local_presentation_qa_attempts.attempts).toEqual(attempts)
    expect(result.omitted_event_count).toBeGreaterThan(0)
    expect(result.events).toEqual(d.snapshot().events.slice(result.omitted_event_count))
    expect(result.events.at(-1).duration_ms).toBe(199)
    expect(d.snapshot().events).toHaveLength(200)
  })
})

it('keeps legacy oversized local exports throwing without a provider and never changes remote or volatile events', () => {
  const d = createOfficeDiagnostics({
    host: 'powerpoint',
    build: 'b'.repeat(64),
    localDocumentId: 'd'.repeat(128),
    localSessionId: () => 's'.repeat(128),
  })
  d.setTool('t'.repeat(128), {
    project_id: 'p'.repeat(128),
    request_id: 'r'.repeat(128),
    page_id: 'a'.repeat(128),
    tool_call_id: 'c'.repeat(128),
  })
  for (let i = 0; i < 200; i++)
    d.record({
      phase: 'tool',
      errorCode: 'office_read_failed',
      error: {
        name: 'n'.repeat(128),
        code: 'e'.repeat(128),
        debugInfo: { errorLocation: 'l'.repeat(128) },
      },
    })
  expect(() => d.exportJson({ includeLocalContext: true })).toThrow('diagnostic_export_too_large')
  const result = JSON.parse(
    d.exportJson({ includeLocalContext: true, screenshotAttempts: () => [screenshotAttempt()] }),
  )
  expect(result.events.length + result.omitted_event_count).toBe(200)
  expect(result.events).toEqual(d.snapshot().events.slice(result.omitted_event_count))
  expect(d.snapshot().events).toHaveLength(200)
})
