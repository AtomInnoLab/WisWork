import { createHash } from 'node:crypto'
import type { PresentationImportRecord } from '../src/skills/powerpoint/presentation-delivery.js'
import { presentationArtifactContent } from '../src/skills/powerpoint/presentation-page-delivery.js'
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
  const readReceipt = vi.fn((_key: string): PresentationImportRecord | undefined => undefined)
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
    readReceipt: (key) => readReceipt(key),
  })
  return { vfs, documentId, request, rememberProject, skill, readReceipt }
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
const rebuild = {
  id: 'rebuild',
  name: 'rebuild_presentation_page',
  input: {
    project_id: 'p',
    parent_request_id: 'parent',
    request_id: 'r',
    page_id: 'one',
    slide: deck.slides[0],
  },
}
const derived = {
  ...summary,
  revision: { parentRequestId: 'parent', pageId: 'one', parentInputDigest: 'a'.repeat(64) },
}
it('creates a derived page task with validated parent identity and blocks its bulk import preparation', async () => {
  const f = fixture()
  f.request.mockImplementation(async () => new Response(JSON.stringify(derived)))
  expect((await f.skill.executeTool(rebuild)).isError).not.toBe(true)
  expect(f.request).toHaveBeenCalledWith(
    expect.objectContaining({
      operation: 'production_rebuild_page',
      parentRequestId: 'parent',
      requestId: 'r',
      pageId: 'one',
      slide: deck.slides[0],
    }),
    undefined,
  )
  f.request.mockImplementation(
    async () => new Response(JSON.stringify({ ...compiled, revision: derived.revision })),
  )
  expect(await f.skill.executeTool(prepare)).toMatchObject({
    isError: true,
    output: 'presentation_page_replacement_required',
  })
  expect(f.skill.artifact()).toBeUndefined()
})
it('rejects malformed derived metadata, wrong parent, missing revision and invalid rebuild input', async () => {
  for (const revision of [
    { ...derived.revision, parentInputDigest: 'bad' },
    { ...derived.revision, pageId: 'missing' },
    { ...derived.revision, parentRequestId: 'r' },
    { ...derived.revision, extra: 1 },
  ])
    expect(() => parsePresentationProductionStatus({ ...summary, revision })).toThrow()
  const f = fixture()
  for (const value of [
    summary,
    { ...derived, revision: { ...derived.revision, parentRequestId: 'other' } },
    { ...derived, revision: { ...derived.revision, pageId: 'two' } },
  ]) {
    f.request.mockImplementation(async () => new Response(JSON.stringify(value)))
    expect(await f.skill.executeTool(rebuild)).toMatchObject({ isError: true })
  }
  f.request.mockClear()
  for (const input of [
    { ...rebuild.input, parent_request_id: 'r' },
    { ...rebuild.input, slide: { ...deck.slides[0], id: 'two' } },
    { ...rebuild.input, extra: 1 },
  ])
    expect(await f.skill.executeTool({ ...rebuild, input })).toMatchObject({
      isError: true,
      output: 'invalid_tool_input',
    })
  expect(f.request).not.toHaveBeenCalled()
})
it('uses the shared slide schema and supports notes while guarding rebuild size, legacy PC and late cancellation', async () => {
  const f = fixture(),
    tool = f.skill.tools.find((t) => t.name === 'rebuild_presentation_page')!
  const { PRESENTATION_DECK_SCHEMA } = await import('@wiswork/pptx-engine/presentation')
  expect((tool.inputSchema as { properties: { slide: unknown } }).properties.slide).toBe(
    (PRESENTATION_DECK_SCHEMA as unknown as { properties: { slides: { items: unknown } } })
      .properties.slides.items,
  )
  f.request.mockImplementation(async () => new Response(JSON.stringify(derived)))
  expect(
    (
      await f.skill.executeTool({
        ...rebuild,
        input: { ...rebuild.input, slide: { ...deck.slides[0], notes: 'Source notes' } },
      })
    ).isError,
  ).not.toBe(true)
  expect(
    await f.skill.executeTool({
      ...rebuild,
      input: { ...rebuild.input, slide: { ...deck.slides[0], notes: 'x'.repeat(256 * 1024) } },
    }),
  ).toMatchObject({ isError: true, output: 'presentation_request_too_large' })
  f.request.mockImplementation(
    async () => new Response(JSON.stringify({ error: 'invalid_request' })),
  )
  expect(await f.skill.executeTool(rebuild)).toMatchObject({
    isError: true,
    output: 'presentation_upgrade_required',
  })
  f.request.mockImplementation(async () => {
    f.skill.clear()
    return new Response(JSON.stringify(derived))
  })
  expect(await f.skill.executeTool(rebuild)).toMatchObject({ isError: true, output: 'cancelled' })
})
it('blocks derived preparation without erasing a prior valid non-derived cache', async () => {
  const f = prepareFixture()
  await f.skill.executeTool(prepare)
  const previous = f.skill.artifact()
  f.request.mockImplementation(
    async () => new Response(JSON.stringify({ ...compiled, revision: derived.revision })),
  )
  expect(await f.skill.executeTool(prepare)).toMatchObject({
    isError: true,
    output: 'presentation_page_replacement_required',
  })
  expect(f.skill.artifact()).toBe(previous)
})

it('loads a committed derived bundle only with its exact complete receipt', async () => {
  const f = prepareFixture()
  await f.skill.executeTool(prepare)
  const artifact = f.skill.artifact()!
  const receipt: PresentationImportRecord = {
    state: 'complete',
    documentId: 'doc',
    slideIds: ['host1', 'host2'],
    checkpoint: {
      version: 2,
      artifactDigest: createHash('sha256')
        .update(presentationArtifactContent(artifact))
        .digest('hex'),
      sourceSlideIds: ['256#', '256#'],
      pageIds: ['one', 'two'],
      baselineSlideIds: [],
      completed: [
        { sourceSlideId: '256#', slideId: 'host1' },
        { sourceSlideId: '256#', slideId: 'host2' },
      ],
    },
  }
  f.readReceipt.mockReturnValue(receipt)
  f.request.mockImplementation(async (body) => {
    const b = body as { operation: string; pageId: string }
    return Response.json(
      b.operation === 'production_status'
        ? { ...compiled, revision: derived.revision }
        : pageArtifact(b.pageId),
    )
  })
  expect((await f.skill.executeTool(prepare)).isError).not.toBe(true)
  expect(f.skill.artifact()).not.toBe(artifact)
  const prepared = f.skill.artifact()
  for (const mutation of ['digest', 'document', 'pages', 'sources', 'incomplete']) {
    const bad = structuredClone(receipt)
    if (mutation === 'digest') bad.checkpoint!.artifactDigest = '0'.repeat(64)
    if (mutation === 'document') bad.documentId = 'other'
    if (mutation === 'pages') (bad.checkpoint as { pageIds: string[] }).pageIds.reverse()
    if (mutation === 'sources') bad.checkpoint!.sourceSlideIds[0] = '257#'
    if (mutation === 'incomplete') {
      bad.state = 'pending'
      delete bad.slideIds
      bad.checkpoint!.completed = []
    }
    f.readReceipt.mockReturnValue(bad)
    expect(await f.skill.executeTool(prepare)).toMatchObject({ isError: true })
    expect(f.skill.artifact()).toBe(prepared)
  }
})
it('rejects superseded and racing receipts without replacing the prepared cache', async () => {
  const f = prepareFixture()
  await f.skill.executeTool(prepare)
  const cached = f.skill.artifact()
  f.readReceipt.mockImplementation(() => {
    throw new Error('presentation_import_superseded')
  })
  expect(await f.skill.executeTool(prepare)).toMatchObject({
    isError: true,
    output: 'presentation_import_superseded',
  })
  expect(f.skill.artifact()).toBe(cached)
  f.readReceipt.mockReturnValue(undefined)
  f.rememberProject.mockImplementation(async () => {
    f.readReceipt.mockImplementation(() => {
      throw new Error('presentation_import_superseded')
    })
  })
  expect(await f.skill.executeTool(prepare)).toMatchObject({ isError: true })
  expect(f.skill.artifact()).toBe(cached)
})

const contentCall = {
  id: 'content',
  name: 'check_presentation_page_content',
  input: { project_id: 'p', request_id: 'r', page_id: 'one' },
}
const contentResponse = () => ({
  projectId: 'p',
  requestId: 'r',
  planRevision: 1,
  inputDigest: 'a'.repeat(64),
  planDigest: 'b'.repeat(64),
  report: {
    version: 1,
    pageId: 'one',
    claimIds: ['c1'],
    findings: [{ code: 'claim_text_not_found', claimId: 'c1' }],
    checks: {
      content: 'needs_review',
      sources: 'not_verified',
      calculations: 'not_verified',
      timeliness: 'not_verified',
      host: 'not_checked',
    },
  },
})
it('checks frozen page content without publishing artifacts, remembering a project or marking QA passed', async () => {
  const f = fixture()
  f.request.mockResolvedValue(new Response(JSON.stringify(contentResponse())))
  expect(f.skill.tools.map((t) => t.name)).toContain(contentCall.name)
  const result = await f.skill.executeTool(contentCall)
  expect(result.isError, result.output).not.toBe(true)
  expect(JSON.parse(result.output)).toMatchObject(contentResponse())
  expect(JSON.parse(result.output).recommendations[0]).toMatchObject({
    code: 'claim_text_not_found',
    action: expect.any(String),
  })
  expect(f.request).toHaveBeenCalledWith(
    {
      operation: 'production_content_check',
      documentId: 'doc',
      projectId: 'p',
      requestId: 'r',
      pageId: 'one',
    },
    undefined,
  )
  expect(f.rememberProject).not.toHaveBeenCalled()
  expect(f.vfs.list('/home/user')).toEqual([])
  expect(f.skill.artifact()).toBeUndefined()
  expect(f.readReceipt).not.toHaveBeenCalled()
})
it('rejects mismatched or forged content reports including fabricated pass and oversized responses', async () => {
  const f = fixture()
  const values = [
    { ...contentResponse(), projectId: 'other' },
    { ...contentResponse(), requestId: 'other' },
    { ...contentResponse(), planRevision: 0 },
    { ...contentResponse(), inputDigest: 'invalid' },
    { ...contentResponse(), planDigest: 'invalid' },
    { ...contentResponse(), extra: true },
    { ...contentResponse(), report: { ...contentResponse().report, pageId: 'other' } },
    {
      ...contentResponse(),
      report: {
        ...contentResponse().report,
        checks: { ...contentResponse().report.checks, sources: 'passed' },
      },
    },
    {
      ...contentResponse(),
      report: {
        ...contentResponse().report,
        findings: [{ code: 'claim_text_not_found', claimId: 'unknown' }],
      },
    },
    { ...contentResponse(), padding: 'x'.repeat(65536) },
  ]
  for (const value of values) {
    f.request.mockResolvedValue(new Response(JSON.stringify(value)))
    expect(await f.skill.executeTool(contentCall)).toMatchObject({
      isError: true,
      output: 'presentation_response_invalid',
    })
  }
  expect(f.rememberProject).not.toHaveBeenCalled()
})
it('requires exact page and request for content checks and handles old PC clearly', async () => {
  const f = fixture()
  for (const input of [
    { project_id: 'p', request_id: 'r' },
    { project_id: 'p', page_id: 'one' },
    { ...contentCall.input, extra: 1 },
  ])
    expect(await f.skill.executeTool({ ...contentCall, input })).toMatchObject({
      isError: true,
      output: 'invalid_tool_input',
    })
  expect(f.request).not.toHaveBeenCalled()
  f.request.mockResolvedValue(new Response(JSON.stringify({ error: 'invalid_request' })))
  expect(await f.skill.executeTool(contentCall)).toMatchObject({
    isError: true,
    output: 'presentation_upgrade_required',
  })
})
it('does not publish content prechecks after clear, cancellation or document switch', async () => {
  for (const mode of ['clear', 'abort', 'document']) {
    const f = fixture(),
      ac = new AbortController()
    f.request.mockImplementation(async () => {
      if (mode === 'clear') f.skill.clear()
      if (mode === 'abort') ac.abort()
      if (mode === 'document') f.documentId.mockResolvedValue('other')
      return new Response(JSON.stringify(contentResponse()))
    })
    expect(await f.skill.executeTool(contentCall, ac.signal)).toMatchObject({ isError: true })
    expect(f.rememberProject).not.toHaveBeenCalled()
    expect(f.vfs.list('/home/user')).toEqual([])
  }
})
it('preserves an already prepared artifact when checking a frozen page', async () => {
  const f = prepareFixture()
  expect((await f.skill.executeTool(prepare)).isError).not.toBe(true)
  const original = f.skill.artifact(),
    files = f.vfs.list('/home/user')
  f.rememberProject.mockClear()
  f.request.mockResolvedValue(new Response(JSON.stringify(contentResponse())))
  expect((await f.skill.executeTool(contentCall)).isError).not.toBe(true)
  expect(f.skill.artifact()).toBe(original)
  expect(f.vfs.list('/home/user')).toEqual(files)
  expect(f.rememberProject).not.toHaveBeenCalled()
})

const evidenceCall = {
  id: 'evidence',
  name: 'read_presentation_claim_evidence',
  input: {
    project_id: 'p',
    request_id: 'r',
    page_id: 'one',
    claim_id: 'c',
    source_id: 's',
    offset: 2,
    max_chars: 8,
  },
}
const evidenceResponse = () => ({
  version: 1,
  projectId: 'p',
  requestId: 'r',
  planRevision: 1,
  inputDigest: 'a'.repeat(64),
  planDigest: 'b'.repeat(64),
  pageId: 'one',
  claimId: 'c',
  statement: 'fact',
  source: { id: 's', uri: `attachment:${'c'.repeat(64)}`, excerpt: 'fact', locator: 'section 1' },
  attachment: {
    id: 'c'.repeat(64),
    name: 'source.txt',
    offset: 2,
    totalChars: 10,
    text: 'a fact z',
    offsetUnit: 'utf16_code_unit',
  },
  excerptMatch: { status: 'found', offset: 4 },
  checks: {
    support: 'not_verified',
    sourceAuthority: 'not_verified',
    timeliness: 'not_verified',
    host: 'not_checked',
  },
})
it('reads exact claim evidence without changing the prepared artifact or other state', async () => {
  const f = prepareFixture()
  await f.skill.executeTool(prepare)
  const original = f.skill.artifact()
  f.rememberProject.mockClear()
  f.request.mockResolvedValue(new Response(JSON.stringify(evidenceResponse())))
  expect(f.skill.tools.map((t) => t.name)).toContain(evidenceCall.name)
  const result = await f.skill.executeTool(evidenceCall)
  expect(result.isError, result.output).not.toBe(true)
  expect(JSON.parse(result.output)).toMatchObject(evidenceResponse())
  expect(f.request).toHaveBeenLastCalledWith(
    {
      operation: 'production_claim_evidence',
      documentId: 'doc',
      projectId: 'p',
      requestId: 'r',
      pageId: 'one',
      claimId: 'c',
      sourceId: 's',
      offset: 2,
      maxChars: 8,
    },
    undefined,
  )
  expect(f.skill.artifact()).toBe(original)
  expect(f.rememberProject).not.toHaveBeenCalled()
  expect(f.vfs.list('/home/user')).toEqual([])
})
it('rejects wrong evidence binding, fake match, unsupported checks and excessive window output', async () => {
  const f = fixture()
  for (const value of [
    { ...evidenceResponse(), projectId: 'other' },
    { ...evidenceResponse(), requestId: 'other' },
    { ...evidenceResponse(), pageId: 'other' },
    { ...evidenceResponse(), claimId: 'other' },
    { ...evidenceResponse(), source: { ...evidenceResponse().source, id: 'other' } },
    { ...evidenceResponse(), attachment: { ...evidenceResponse().attachment, offset: 3 } },
    {
      ...evidenceResponse(),
      attachment: { ...evidenceResponse().attachment, totalChars: 11, text: 'a fact zz' },
    },
    { ...evidenceResponse(), excerptMatch: { status: 'found', offset: 3 } },
    { ...evidenceResponse(), checks: { ...evidenceResponse().checks, support: 'verified' } },
    { ...evidenceResponse(), extra: true },
  ]) {
    f.request.mockResolvedValue(new Response(JSON.stringify(value)))
    expect(await f.skill.executeTool(evidenceCall)).toMatchObject({
      isError: true,
      output: 'presentation_response_invalid',
    })
  }
  expect(f.rememberProject).not.toHaveBeenCalled()
})
it('validates evidence window inputs, old PC and unsupported sources', async () => {
  const f = fixture()
  for (const change of [
    { claim_id: undefined },
    { source_id: undefined },
    { offset: -1 },
    { offset: 1.5 },
    { max_chars: 8001 },
    { max_chars: 0 },
    { extra: 1 },
  ])
    expect(
      await f.skill.executeTool({ ...evidenceCall, input: { ...evidenceCall.input, ...change } }),
    ).toMatchObject({ isError: true, output: 'invalid_tool_input' })
  expect(f.request).not.toHaveBeenCalled()
  for (const [error, output] of [
    ['invalid_request', 'presentation_upgrade_required'],
    ['evidence_source_unsupported', 'presentation_evidence_source_unsupported'],
  ]) {
    f.request.mockResolvedValue(new Response(JSON.stringify({ error })))
    expect(await f.skill.executeTool(evidenceCall)).toMatchObject({ isError: true, output })
  }
})
it('rejects evidence results after cancellation, clear or document change', async () => {
  for (const mode of ['abort', 'clear', 'document']) {
    const f = fixture(),
      ac = new AbortController()
    f.request.mockImplementation(async () => {
      if (mode === 'abort') ac.abort()
      if (mode === 'clear') f.skill.clear()
      if (mode === 'document') f.documentId.mockResolvedValue('other')
      return new Response(JSON.stringify(evidenceResponse()))
    })
    expect(await f.skill.executeTool(evidenceCall, ac.signal)).toMatchObject({ isError: true })
    expect(f.rememberProject).not.toHaveBeenCalled()
  }
})

const recordReviewCall = {
  id: 'record-review',
  name: 'record_presentation_claim_review',
  input: {
    ...evidenceCall.input,
    review_id: 'review1',
    outcome: 'supported',
    notes: 'The supplied passage supports this limited claim.',
  },
}
const readReviewCall = {
  id: 'read-review',
  name: 'read_presentation_claim_review',
  input: { project_id: 'p', request_id: 'r', review_id: 'review1' },
}
const claimReviewResponse = (evidenceDigest = 'd'.repeat(64)) => ({
  version: 1,
  projectId: 'p',
  requestId: 'r',
  reviewId: 'review1',
  planRevision: 1,
  inputDigest: 'a'.repeat(64),
  planDigest: 'b'.repeat(64),
  pageId: 'one',
  claimId: 'c',
  sourceId: 's',
  attachmentId: 'c'.repeat(64),
  offset: 2,
  maxChars: 8,
  evidenceDigest,
  outcome: 'supported',
  notes: recordReviewCall.input.notes,
  reviewer: 'agent',
  createdAt: '2026-09-24T00:00:00.000Z',
  checks: {
    support: 'agent_reviewed',
    sourceAuthority: 'not_verified',
    timeliness: 'not_verified',
    host: 'not_checked',
  },
})
it('requires live evidence before recording a claim judgment and binds the write to its internal digest', async () => {
  const f = fixture()
  expect(await f.skill.executeTool(recordReviewCall)).toMatchObject({
    isError: true,
    output: 'presentation_evidence_read_required',
  })
  expect(f.request).not.toHaveBeenCalled()
  f.request.mockResolvedValue(new Response(JSON.stringify(evidenceResponse())))
  expect((await f.skill.executeTool(evidenceCall)).isError).not.toBe(true)
  f.request.mockImplementation(
    async (body) =>
      new Response(
        JSON.stringify(claimReviewResponse((body as { evidenceDigest: string }).evidenceDigest)),
      ),
  )
  const result = await f.skill.executeTool(recordReviewCall)
  expect(result.isError, result.output).not.toBe(true)
  expect(f.request).toHaveBeenLastCalledWith(
    expect.objectContaining({
      operation: 'production_record_claim_review',
      documentId: 'doc',
      reviewId: 'review1',
      evidenceDigest: expect.stringMatching(/^[a-f0-9]{64}$/),
      outcome: 'supported',
      notes: recordReviewCall.input.notes,
    }),
    undefined,
  )
  expect(JSON.parse(result.output)).toMatchObject({
    reviewer: 'agent',
    checks: { support: 'agent_reviewed', host: 'not_checked' },
  })
  expect(f.rememberProject).not.toHaveBeenCalled()
  expect(f.vfs.list('/home/user')).toEqual([])
})
it('reads historical review without granting permission to write another judgment', async () => {
  const f = fixture()
  f.request.mockResolvedValue(new Response(JSON.stringify(claimReviewResponse())))
  expect((await f.skill.executeTool(readReviewCall)).isError).not.toBe(true)
  expect(await f.skill.executeTool(recordReviewCall)).toMatchObject({
    isError: true,
    output: 'presentation_evidence_read_required',
  })
  expect(f.request).toHaveBeenCalledTimes(1)
  expect(f.rememberProject).not.toHaveBeenCalled()
})
it('rejects forged review outcomes, wrong frozen identities and stale live evidence', async () => {
  const f = fixture()
  f.request.mockResolvedValue(new Response(JSON.stringify(evidenceResponse())))
  await f.skill.executeTool(evidenceCall)
  for (const change of [
    { notes: 'altered' },
    { outcome: 'contradicted' },
    { pageId: 'other' },
    { evidenceDigest: 'e'.repeat(64) },
    { planDigest: 'e'.repeat(64) },
    { inputDigest: 'e'.repeat(64) },
    { reviewer: 'human' },
    { reviewId: 'other' },
  ]) {
    f.request.mockImplementation(
      async (body) =>
        new Response(
          JSON.stringify({
            ...claimReviewResponse((body as { evidenceDigest: string }).evidenceDigest),
            ...change,
          }),
        ),
    )
    expect(await f.skill.executeTool(recordReviewCall)).toMatchObject({
      isError: true,
      output: 'presentation_response_invalid',
    })
  }
  f.skill.clear()
  expect(await f.skill.executeTool(recordReviewCall)).toMatchObject({
    isError: true,
    output: 'presentation_evidence_read_required',
  })
})
it('validates claim review inputs and preserves explicit server conflicts', async () => {
  const f = fixture()
  for (const change of [
    { review_id: undefined },
    { outcome: 'verified' },
    { notes: '' },
    { notes: ' ' },
    { notes: 'x'.repeat(2001) },
    { notes: '\u000b' },
    { notes: '\uffff' },
    { notes: '\ud800' },
    { evidence_digest: 'a'.repeat(64) },
  ])
    expect(
      await f.skill.executeTool({
        ...recordReviewCall,
        input: { ...recordReviewCall.input, ...change },
      }),
    ).toMatchObject({ isError: true, output: 'invalid_tool_input' })
  f.request.mockResolvedValue(new Response(JSON.stringify(evidenceResponse())))
  await f.skill.executeTool(evidenceCall)
  for (const error of ['evidence_changed', 'request_conflict', 'quota_exceeded']) {
    f.request.mockResolvedValue(new Response(JSON.stringify({ error })))
    expect(await f.skill.executeTool(recordReviewCall)).toMatchObject({
      isError: true,
      output: `presentation_${error}`,
    })
  }
})
it('does not authorize reviews from evidence reads cancelled before publication', async () => {
  for (const mode of ['abort', 'clear', 'document']) {
    const f = fixture(),
      ac = new AbortController()
    f.request.mockImplementation(async () => {
      if (mode === 'abort') ac.abort()
      if (mode === 'clear') f.skill.clear()
      if (mode === 'document') f.documentId.mockResolvedValue('other')
      return new Response(JSON.stringify(evidenceResponse()))
    })
    expect((await f.skill.executeTool(evidenceCall, ac.signal)).isError).toBe(true)
    f.documentId.mockResolvedValue('doc')
    f.request.mockClear()
    expect(await f.skill.executeTool(recordReviewCall)).toMatchObject({
      isError: true,
      output: 'presentation_evidence_read_required',
    })
    expect(f.request).not.toHaveBeenCalled()
  }
})
it('refuses a late review response after clear, abort or document switch without changing active state', async () => {
  for (const mode of ['abort', 'clear', 'document']) {
    const f = fixture(),
      ac = new AbortController()
    f.request.mockResolvedValue(new Response(JSON.stringify(evidenceResponse())))
    await f.skill.executeTool(evidenceCall)
    f.request.mockImplementation(async (body) => {
      if (mode === 'abort') ac.abort()
      if (mode === 'clear') f.skill.clear()
      if (mode === 'document') f.documentId.mockResolvedValue('other')
      return new Response(
        JSON.stringify(claimReviewResponse((body as { evidenceDigest: string }).evidenceDigest)),
      )
    })
    expect((await f.skill.executeTool(recordReviewCall, ac.signal)).isError).toBe(true)
    expect(f.rememberProject).not.toHaveBeenCalled()
    expect(f.vfs.list('/home/user')).toEqual([])
  }
})

const pageReviewsCall = {
  id: 'page-reviews',
  name: 'read_presentation_page_reviews',
  input: { project_id: 'p', request_id: 'r', page_id: 'one' },
}
const pageReviewsResponse = () => ({
  version: 1,
  projectId: 'p',
  requestId: 'r',
  pageId: 'one',
  planRevision: 1,
  inputDigest: 'a'.repeat(64),
  planDigest: 'b'.repeat(64),
  claims: [
    {
      claimId: 'c',
      status: 'unreviewed',
      sources: [{ sourceId: 's', status: 'unreviewed', reviews: [] }],
    },
  ],
  checks: {
    support: 'historical_agent_reviews',
    sourceAuthority: 'not_verified',
    timeliness: 'not_verified',
    host: 'not_checked',
  },
})
it('reads page review history without altering the prepared artifact or authorizing a new review', async () => {
  const f = prepareFixture()
  await f.skill.executeTool(prepare)
  const original = f.skill.artifact()
  f.rememberProject.mockClear()
  f.request.mockResolvedValue(new Response(JSON.stringify(pageReviewsResponse())))
  expect(f.skill.tools.map((t) => t.name)).toContain(pageReviewsCall.name)
  const result = await f.skill.executeTool(pageReviewsCall)
  expect(result.isError, result.output).not.toBe(true)
  expect(JSON.parse(result.output)).toEqual(pageReviewsResponse())
  expect(f.request).toHaveBeenLastCalledWith(
    {
      operation: 'production_page_reviews',
      documentId: 'doc',
      projectId: 'p',
      requestId: 'r',
      pageId: 'one',
    },
    undefined,
  )
  expect(f.skill.artifact()).toBe(original)
  expect(f.rememberProject).not.toHaveBeenCalled()
  expect(f.vfs.list('/home/user')).toEqual([])
  expect(await f.skill.executeTool(recordReviewCall)).toMatchObject({
    isError: true,
    output: 'presentation_evidence_read_required',
  })
})
it('rejects mismatched page history and fabricated aggregate status', async () => {
  const f = fixture()
  for (const response of [
    { ...pageReviewsResponse(), projectId: 'other' },
    { ...pageReviewsResponse(), requestId: 'other' },
    { ...pageReviewsResponse(), pageId: 'other' },
    { ...pageReviewsResponse(), planRevision: 0 },
    {
      ...pageReviewsResponse(),
      claims: [{ ...pageReviewsResponse().claims[0], status: 'supported' }],
    },
    { ...pageReviewsResponse(), checks: { ...pageReviewsResponse().checks, support: 'verified' } },
    { ...pageReviewsResponse(), extra: true },
    { ...pageReviewsResponse(), padding: 'x'.repeat(65536) },
  ]) {
    f.request.mockResolvedValue(new Response(JSON.stringify(response)))
    expect(await f.skill.executeTool(pageReviewsCall)).toMatchObject({
      isError: true,
      output: 'presentation_response_invalid',
    })
  }
})
it('requires explicit page/request for review summaries and preserves old PC upgrade errors', async () => {
  const f = fixture()
  for (const input of [
    { project_id: 'p', request_id: 'r' },
    { project_id: 'p', page_id: 'one' },
    { ...pageReviewsCall.input, extra: true },
  ])
    expect(await f.skill.executeTool({ ...pageReviewsCall, input })).toMatchObject({
      isError: true,
      output: 'invalid_tool_input',
    })
  expect(f.request).not.toHaveBeenCalled()
  f.request.mockResolvedValue(new Response(JSON.stringify({ error: 'invalid_request' })))
  expect(await f.skill.executeTool(pageReviewsCall)).toMatchObject({
    isError: true,
    output: 'presentation_upgrade_required',
  })
})
it('does not publish page history after clear, cancellation or a document switch', async () => {
  for (const mode of ['clear', 'abort', 'document']) {
    const f = fixture(),
      ac = new AbortController()
    f.request.mockImplementation(async () => {
      if (mode === 'clear') f.skill.clear()
      if (mode === 'abort') ac.abort()
      if (mode === 'document') f.documentId.mockResolvedValue('other')
      return new Response(JSON.stringify(pageReviewsResponse()))
    })
    expect((await f.skill.executeTool(pageReviewsCall, ac.signal)).isError).toBe(true)
    expect(f.rememberProject).not.toHaveBeenCalled()
  }
})

it('explains as-of metadata findings conservatively without changing session state', async () => {
  const f = fixture()
  const value = contentResponse()
  const findings = [
    { code: 'source_as_of_missing', claimId: 'c1', sourceId: 's1' },
    { code: 'source_as_of_differs', claimId: 'c1', sourceId: 's2' },
  ]
  f.request.mockResolvedValue(
    new Response(JSON.stringify({ ...value, report: { ...value.report, findings } })),
  )
  const result = await f.skill.executeTool(contentCall)
  expect(result.isError, result.output).not.toBe(true)
  const output = JSON.parse(result.output)
  expect(output.report.findings).toEqual(findings)
  expect(output.report.checks.timeliness).toBe('not_verified')
  expect(output.recommendations).toEqual([
    { code: 'source_as_of_missing', action: expect.stringContaining('时点') },
    { code: 'source_as_of_differs', action: expect.stringContaining('不代表过期') },
  ])
  expect(f.rememberProject).not.toHaveBeenCalled()
  expect(f.readReceipt).not.toHaveBeenCalled()
  expect(f.vfs.list('/home/user')).toEqual([])
  expect(f.skill.artifact()).toBeUndefined()
})

it('accepts a maximum bounded content report above 64 KiB and rejects transport above 256 KiB', async () => {
  const f = fixture()
  const value = contentResponse()
  const claimIds = Array.from({ length: 32 }, (_, i) => `c${i}`.padEnd(80, 'c'))
  const sources = Array.from({ length: 3 }, (_, i) => `s${i}`.padEnd(80, 's'))
  const findings = claimIds.flatMap((claimId) => [
    { code: 'claim_text_not_found', claimId },
    { code: 'calculation_not_reproduced', claimId },
    ...sources.flatMap((sourceId) =>
      ['source_excerpt_missing', 'source_locator_missing', 'source_as_of_missing'].map((code) => ({
        code,
        claimId,
        sourceId,
      })),
    ),
  ])
  const body = JSON.stringify({ ...value, report: { ...value.report, claimIds, findings } })
  expect(findings).toHaveLength(352)
  expect(new TextEncoder().encode(body).byteLength).toBeGreaterThan(64 * 1024)
  f.request.mockResolvedValue(new Response(body))
  const result = await f.skill.executeTool(contentCall)
  expect(result.isError, result.output).not.toBe(true)
  expect(JSON.parse(result.output).report.findings).toEqual(findings)
  // JSON whitespace keeps the payload structurally valid, testing the byte limit itself.
  f.request.mockResolvedValue(new Response(body + ' '.repeat(256 * 1024)))
  expect(await f.skill.executeTool(contentCall)).toMatchObject({
    isError: true,
    output: 'presentation_response_invalid',
  })
})
