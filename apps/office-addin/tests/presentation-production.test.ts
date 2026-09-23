import { expect, it, vi } from 'vitest'
import {
  createPresentationProductionSkill,
  parsePresentationProductionStatus,
} from '../src/skills/powerpoint/presentation-production.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'
const summary = {
  projectId: 'p',
  requestId: 'r',
  planRevision: 1,
  status: 'partial',
  compiledCount: 1,
  total: 2,
  pages: [
    { id: 'one', title: 'One', state: 'compiled', attempt: 1 },
    { id: 'two', title: 'Two', state: 'failed', attempt: 1, error: 'compile_failed' },
  ],
}
const deck = {
  version: 1,
  id: 'p',
  title: 'Project',
  style: { fontFace: 'Arial', background: 'FFFFFF', textColor: '000000', accentColor: '223344' },
  assets: [],
  claims: [],
  slides: [
    {
      id: 'one',
      title: 'One',
      elements: [{ id: 't', kind: 'text', text: 'Title', x: 1, y: 1, w: 3, h: 1 }],
    },
  ],
}
function fixture() {
  const vfs = new InMemoryVfs(),
    documentId = vi.fn(async () => 'doc'),
    request = vi.fn(
      async (_body: unknown, _signal?: AbortSignal) => new Response(JSON.stringify(summary)),
    ),
    rememberProject = vi.fn(async (_id: string) => {})
  const skill = createPresentationProductionSkill({
    vfs,
    documentId,
    request,
    rememberProject,
    available: () => true,
    lastProject: () => 'p',
  })
  return { vfs, documentId, request, rememberProject, skill }
}
const run = {
  id: 'c',
  name: 'run_presentation_production',
  input: { project_id: 'p', request_id: 'r' },
}
it('runs partial production and keeps compiled distinct from host delivery', async () => {
  const f = fixture()
  expect(await f.skill.executeTool(run)).toMatchObject({
    mutated: false,
    output: JSON.stringify(summary),
  })
  expect(f.request).toHaveBeenCalledWith(
    { operation: 'production_run', documentId: 'doc', projectId: 'p', requestId: 'r' },
    undefined,
  )
  expect(f.vfs.list('/home/user')).toEqual([])
  expect(f.rememberProject).toHaveBeenCalledWith('p')
})
it('rejects forged summaries and unknown fields', () => {
  for (const value of [
    { ...summary, compiledCount: 2 },
    { ...summary, status: 'compiled' },
    { ...summary, extra: 1 },
    { ...summary, pages: [summary.pages[0], summary.pages[0]] },
    { ...summary, pages: [summary.pages[0], { ...summary.pages[1], error: 'secret' }] },
  ])
    expect(() => parsePresentationProductionStatus(value)).toThrow()
})
it('rejects identity mismatch, unknown input, and old PC without claiming support', async () => {
  const f = fixture()
  f.request.mockResolvedValue(new Response(JSON.stringify({ ...summary, requestId: 'other' })))
  expect(await f.skill.executeTool(run)).toMatchObject({ isError: true })
  expect(f.rememberProject).not.toHaveBeenCalled()
  expect(await f.skill.executeTool({ ...run, input: { ...run.input, extra: 1 } })).toMatchObject({
    isError: true,
    output: 'invalid_tool_input',
  })
  f.request.mockResolvedValue(new Response(JSON.stringify({ error: 'invalid_request' })))
  expect(await f.skill.executeTool(run)).toMatchObject({
    isError: true,
    summary: expect.stringContaining('升级'),
  })
})
it('validates begin revision and response pages against frozen input', async () => {
  const f = fixture()
  expect(
    await f.skill.executeTool({
      id: 'b',
      name: 'start_presentation_production',
      input: { request_id: 'r', deck },
    }),
  ).toMatchObject({ isError: true })
  expect(f.request).not.toHaveBeenCalled()
  expect(
    await f.skill.executeTool({
      id: 'b',
      name: 'start_presentation_production',
      input: { request_id: 'r', deck, plan_revision: 1 },
    }),
  ).toMatchObject({ isError: true })
})
it('publishes no late response or VFS artifact across clear/document change/abort', async () => {
  for (const mode of ['clear', 'document', 'abort']) {
    const f = fixture(),
      ac = new AbortController()
    f.request.mockImplementation(async () => {
      if (mode === 'clear') f.skill.clear()
      if (mode === 'document') f.documentId.mockResolvedValue('other')
      if (mode === 'abort') ac.abort()
      return new Response(JSON.stringify(summary))
    })
    expect(await f.skill.executeTool(run, ac.signal)).toMatchObject({ isError: true })
    expect(f.vfs.list('/home/user')).toEqual([])
    expect(f.rememberProject).not.toHaveBeenCalled()
  }
})
it('downloads an isolated single-page artifact and report without binary model output', async () => {
  const f = fixture()
  f.request.mockResolvedValue(
    new Response(
      JSON.stringify({
        projectId: 'p',
        requestId: 'r',
        pageId: 'one',
        planRevision: 1,
        status: 'compiled',
        pptxBase64: 'UEsDBAAAAAA=',
        sourceSlideId: '256#',
        report: { deckId: 'p', slideCount: 1 },
      }),
    ),
  )
  const result = await f.skill.executeTool({
    id: 'a',
    name: 'read_presentation_page_artifact',
    input: { project_id: 'p', request_id: 'r', page_id: 'one' },
  })
  expect(result.isError).not.toBe(true)
  expect(result.output).not.toContain('UEsD')
  expect(JSON.parse(result.output)).toMatchObject({
    sourceSlideId: '256#',
    report: { slideCount: 1 },
  })
  expect(f.vfs.list('/home/user')).toHaveLength(2)
})
it('starts and reads production with strict operations and saved revision', async () => {
  const f = fixture(),
    one = {
      ...summary,
      status: 'pending',
      compiledCount: 0,
      total: 1,
      pages: [{ id: 'one', title: 'One', state: 'pending', attempt: 0 }],
    }
  f.request.mockImplementation(async () => new Response(JSON.stringify(one)))
  expect(
    (
      await f.skill.executeTool({
        id: 'b',
        name: 'start_presentation_production',
        input: { request_id: 'r', deck, plan_revision: 1 },
      })
    ).isError,
  ).not.toBe(true)
  expect(f.request).toHaveBeenCalledWith(
    expect.objectContaining({ operation: 'production_begin', planRevision: 1, deck }),
    undefined,
  )
  expect(
    (
      await f.skill.executeTool({
        id: 's',
        name: 'read_presentation_production',
        input: { project_id: 'p' },
      })
    ).isError,
  ).not.toBe(true)
  expect(f.request).toHaveBeenLastCalledWith(
    { operation: 'production_status', projectId: 'p', documentId: 'doc' },
    undefined,
  )
})
it('rejects oversized response, forged page IDs and noncanonical artifact bytes', async () => {
  const f = fixture(),
    call = {
      id: 'a',
      name: 'read_presentation_page_artifact',
      input: { project_id: 'p', request_id: 'r', page_id: 'one' },
    }
  const artifact = {
    projectId: 'p',
    requestId: 'r',
    pageId: 'one',
    planRevision: 1,
    status: 'compiled',
    pptxBase64: 'UEsDBAAAAAA=',
    sourceSlideId: '256#',
    report: { deckId: 'p', slideCount: 1 },
  }
  for (const bad of [
    { ...artifact, pageId: 'two' },
    { ...artifact, sourceSlideId: '0256#' },
    { ...artifact, pptxBase64: 'UEsDBAAAAAB=' },
    { ...artifact, report: { ...artifact.report, note: 'x'.repeat(49 * 1024) } },
  ]) {
    f.request.mockResolvedValue(new Response(JSON.stringify(bad)))
    expect(await f.skill.executeTool(call)).toMatchObject({ isError: true })
    expect(f.vfs.list('/home/user')).toEqual([])
  }
  f.request.mockResolvedValue(new Response(' '.repeat(64 * 1024 + 1)))
  expect(await f.skill.executeTool(run)).toMatchObject({
    isError: true,
    output: 'presentation_response_invalid',
  })
})
it('prevents publication after project remembering changes document', async () => {
  const f = fixture()
  f.rememberProject.mockImplementation(async () => {
    f.documentId.mockResolvedValue('other')
  })
  expect(await f.skill.executeTool(run)).toMatchObject({
    isError: true,
    output: 'presentation_document_changed',
  })
  expect(f.vfs.list('/home/user')).toEqual([])
})
it('freezes request identity across async dispatch', async () => {
  const f = fixture(),
    call = { ...run, input: { ...run.input } }
  f.request.mockImplementation(async () => {
    call.input.request_id = 'changed'
    return new Response(JSON.stringify(summary))
  })
  expect((await f.skill.executeTool(call)).isError).not.toBe(true)
})
it('requires negotiated asset support before sending reference assets', async () => {
  const f = fixture()
  const result = await f.skill.executeTool({
    id: 'b',
    name: 'start_presentation_production',
    input: {
      request_id: 'r',
      plan_revision: 1,
      deck: { ...deck, assets: [{ id: 'asset', attachmentId: 'a'.repeat(64) }] },
    },
  })
  expect(result).toMatchObject({ isError: true, output: 'presentation_assets_unavailable' })
  expect(f.request).not.toHaveBeenCalled()
})
it('does not publish artifact paths when the VFS batch fails', async () => {
  const f = fixture()
  vi.spyOn(f.vfs, 'writeBatch').mockImplementation(() => {
    throw new Error('vfs_limit')
  })
  f.request.mockResolvedValue(
    new Response(
      JSON.stringify({
        projectId: 'p',
        requestId: 'r',
        pageId: 'one',
        planRevision: 1,
        status: 'compiled',
        pptxBase64: 'UEsDBAAAAAA=',
        sourceSlideId: '256#',
        report: { deckId: 'p', slideCount: 1 },
      }),
    ),
  )
  const result = await f.skill.executeTool({
    id: 'a',
    name: 'read_presentation_page_artifact',
    input: { project_id: 'p', request_id: 'r', page_id: 'one' },
  })
  expect(result.isError).toBe(true)
  expect(result.output).not.toContain('/home/user')
  expect(f.vfs.list('/home/user')).toEqual([])
})
const prepare = {
  id: 'prepare',
  name: 'prepare_presentation_production_import',
  input: { project_id: 'p', request_id: 'r' },
}
const compiled = {
  ...summary,
  status: 'compiled',
  compiledCount: 2,
  pages: summary.pages.map((p) => ({ id: p.id, title: p.title, state: 'compiled', attempt: 1 })),
}
function pageArtifact(pageId: string, extra = {}) {
  return {
    projectId: 'p',
    requestId: 'r',
    pageId,
    planRevision: 1,
    status: 'compiled',
    pptxBase64: 'UEsDBAAAAAA=',
    sourceSlideId: '256#',
    report: { deckId: 'p', slideCount: 1 },
    ...extra,
  }
}
function prepareFixture() {
  const f = fixture()
  f.request.mockImplementation(async (body) => {
    const b = body as { operation: string; pageId?: string }
    return new Response(
      JSON.stringify(b.operation === 'production_status' ? compiled : pageArtifact(b.pageId!)),
    )
  })
  return f
}
it('prepares ordered independent pages atomically without VFS publication and replaces cache identity on retry', async () => {
  const f = prepareFixture()
  expect(f.skill.artifact()).toBeUndefined()
  expect((await f.skill.executeTool(prepare)).isError).not.toBe(true)
  const first = f.skill.artifact('p')!
  expect(first).toMatchObject({
    documentId: 'doc',
    projectId: 'p',
    requestId: 'r',
    slideCount: 2,
    pptxBase64: '',
    planRevision: 1,
    pagePptxBase64: ['UEsDBAAAAAA=', 'UEsDBAAAAAA='],
    pages: [
      { id: 'one', title: 'One', sourceSlideId: '256#' },
      { id: 'two', title: 'Two', sourceSlideId: '256#' },
    ],
  })
  expect(f.request.mock.calls.map(([b]) => (b as { operation: string }).operation)).toEqual([
    'production_status',
    'production_page',
    'production_page',
  ])
  expect(f.vfs.list('/home/user')).toEqual([])
  await f.skill.executeTool(prepare)
  expect(f.skill.artifact()).not.toBe(first)
  f.skill.clear()
  expect(f.skill.artifact()).toBeUndefined()
})
it('keeps previous cache if partial production or a later page has wrong revision or identity', async () => {
  const f = prepareFixture()
  await f.skill.executeTool(prepare)
  const first = f.skill.artifact()
  for (const bad of [summary, pageArtifact('two', { planRevision: 2 }), pageArtifact('wrong')]) {
    f.request.mockImplementation(async (body) => {
      const b = body as { operation: string; pageId?: string }
      return new Response(
        JSON.stringify(
          b.operation === 'production_status'
            ? 'pages' in bad
              ? bad
              : compiled
            : b.pageId === 'two'
              ? bad
              : pageArtifact(b.pageId!),
        ),
      )
    })
    expect(await f.skill.executeTool(prepare)).toMatchObject({ isError: true })
    expect(f.skill.artifact()).toBe(first)
  }
})
it('never publishes a prepared artifact when cancelled or document changes between pages or during remember', async () => {
  for (const mode of ['clear', 'abort', 'document', 'remember']) {
    const f = prepareFixture(),
      ac = new AbortController()
    f.request.mockImplementation(async (body) => {
      const b = body as { operation: string; pageId?: string }
      if (b.pageId === 'two') {
        if (mode === 'clear') f.skill.clear()
        if (mode === 'abort') ac.abort()
        if (mode === 'document') f.documentId.mockResolvedValue('other')
      }
      return new Response(
        JSON.stringify(b.operation === 'production_status' ? compiled : pageArtifact(b.pageId!)),
      )
    })
    if (mode === 'remember')
      f.rememberProject.mockImplementation(async () => {
        f.documentId.mockResolvedValue('other')
      })
    expect(await f.skill.executeTool(prepare, ac.signal)).toMatchObject({ isError: true })
    expect(f.skill.artifact()).toBeUndefined()
  }
})
it('enforces cumulative decoded preparation budget and requires explicit request identity', async () => {
  const f = prepareFixture(),
    bytes = Buffer.alloc(6 * 1024 * 1024)
  bytes.write('PK\x03\x04')
  f.request.mockImplementation(async (body) => {
    const b = body as { operation: string; pageId?: string }
    return new Response(
      JSON.stringify(
        b.operation === 'production_status'
          ? compiled
          : pageArtifact(b.pageId!, { pptxBase64: bytes.toString('base64') }),
      ),
    )
  })
  expect(await f.skill.executeTool(prepare)).toMatchObject({
    isError: true,
    output: 'presentation_output_too_large',
  })
  expect(f.skill.artifact()).toBeUndefined()
  expect(await f.skill.executeTool({ ...prepare, input: { project_id: 'p' } })).toMatchObject({
    isError: true,
    output: 'invalid_tool_input',
  })
})
it('keeps only the last prepared project and ordinary downloads never replace its import cache', async () => {
  const f = prepareFixture()
  await f.skill.executeTool(prepare)
  const previous = f.skill.artifact()!
  f.request.mockImplementation(async (body) => {
    const b = body as { operation: string; projectId: string; pageId?: string }
    return new Response(
      JSON.stringify(
        b.operation === 'production_status'
          ? { ...compiled, projectId: b.projectId }
          : pageArtifact(b.pageId!, {
              projectId: b.projectId,
              report: { deckId: b.projectId, slideCount: 1 },
            }),
      ),
    )
  })
  expect(
    (
      await f.skill.executeTool({
        id: 'download',
        name: 'read_presentation_page_artifact',
        input: { project_id: 'q', request_id: 'r', page_id: 'one' },
      })
    ).isError,
  ).not.toBe(true)
  expect(f.skill.artifact()).toBe(previous)
  expect(
    (await f.skill.executeTool({ ...prepare, input: { project_id: 'q', request_id: 'r' } }))
      .isError,
  ).not.toBe(true)
  expect(f.skill.artifact('p')).toBeUndefined()
  expect(f.skill.artifact('q')?.projectId).toBe('q')
  expect(Object.isFrozen(f.skill.artifact())).toBe(true)
  expect(Object.isFrozen(f.skill.artifact()?.pagePptxBase64)).toBe(true)
})
it('does not replace a valid cache when final project persistence fails', async () => {
  const f = prepareFixture()
  await f.skill.executeTool(prepare)
  const previous = f.skill.artifact()
  f.rememberProject.mockRejectedValue(new Error('save_failed'))
  expect(await f.skill.executeTool(prepare)).toMatchObject({ isError: true })
  expect(f.skill.artifact()).toBe(previous)
})
it('accepts exactly 10MiB across two pages and retains duplicate numeric source IDs', async () => {
  const f = prepareFixture(),
    bytes = Buffer.alloc(5 * 1024 * 1024)
  bytes.write('PK\x03\x04')
  f.request.mockImplementation(async (body) => {
    const b = body as { operation: string; pageId?: string }
    return new Response(
      JSON.stringify(
        b.operation === 'production_status'
          ? compiled
          : pageArtifact(b.pageId!, { pptxBase64: bytes.toString('base64') }),
      ),
    )
  })
  expect((await f.skill.executeTool(prepare)).isError).not.toBe(true)
  expect(
    f.skill.artifact()?.pagePptxBase64?.reduce((n, s) => n + Buffer.from(s, 'base64').length, 0),
  ).toBe(10 * 1024 * 1024)
})
it('does not let an older pending prepare overwrite a newer request for the same project', async () => {
  const f = prepareFixture()
  let release!: () => void, entered!: () => void
  const waiting = new Promise<void>((resolve) => {
      release = resolve
    }),
    started = new Promise<void>((resolve) => {
      entered = resolve
    })
  f.request.mockImplementation(async (body) => {
    const b = body as { operation: string; requestId: string; pageId?: string }
    if (b.operation === 'production_page' && b.requestId === 'old' && b.pageId === 'one') {
      entered()
      await waiting
    }
    return new Response(
      JSON.stringify(
        b.operation === 'production_status'
          ? { ...compiled, requestId: b.requestId }
          : pageArtifact(b.pageId!, { requestId: b.requestId }),
      ),
    )
  })
  const old = f.skill.executeTool({ ...prepare, input: { project_id: 'p', request_id: 'old' } })
  await started
  expect(
    (await f.skill.executeTool({ ...prepare, input: { project_id: 'p', request_id: 'new' } }))
      .isError,
  ).not.toBe(true)
  const latest = f.skill.artifact()!
  expect(latest.requestId).toBe('new')
  release()
  expect(await old).toMatchObject({ isError: true, output: 'cancelled' })
  expect(f.skill.artifact()).toBe(latest)
  expect(f.rememberProject).toHaveBeenCalledTimes(1)
})
it('does not invalidate an active prepare for a malformed later invocation', async () => {
  const f = prepareFixture()
  let release!: () => void, entered!: () => void
  const waiting = new Promise<void>((resolve) => {
      release = resolve
    }),
    started = new Promise<void>((resolve) => {
      entered = resolve
    })
  f.request.mockImplementation(async (body) => {
    const b = body as { operation: string; pageId?: string }
    if (b.pageId === 'one') {
      entered()
      await waiting
    }
    return new Response(
      JSON.stringify(b.operation === 'production_status' ? compiled : pageArtifact(b.pageId!)),
    )
  })
  const first = f.skill.executeTool(prepare)
  await started
  expect(await f.skill.executeTool({ ...prepare, input: { project_id: 'p' } })).toMatchObject({
    isError: true,
    output: 'invalid_tool_input',
  })
  release()
  expect((await first).isError).not.toBe(true)
  expect(f.skill.artifact()?.requestId).toBe('r')
})
