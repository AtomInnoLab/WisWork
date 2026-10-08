import { describe, expect, it, vi } from 'vitest'
import { createPresentationResearchSkill } from '../src/skills/powerpoint/presentation-research.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'
import { createHash } from 'node:crypto'
const canonical = (v: unknown): string =>
  Array.isArray(v)
    ? `[${v.map(canonical).join(',')}]`
    : v && typeof v === 'object'
      ? `{${Object.keys(v)
          .sort()
          .map((k) => `${JSON.stringify(k)}:${canonical((v as Record<string, unknown>)[k])}`)
          .join(',')}}`
      : JSON.stringify(v)
const draft = {
  scope: '科研汇报研究范围与争议',
  sources: [
    {
      id: 'source1',
      title: '用户原始资料',
      uri: `attachment:${'a'.repeat(64)}`,
      excerpt: '原始证据有局限',
    },
  ],
  facts: [
    {
      claimId: 'claim1',
      statement: '据原文结论需谨慎',
      type: 'judgment',
      sourceRefs: ['source1'],
      sourceTier: 'unverified',
      slideRefs: ['proposed1'],
      confidence: 'low',
      reviewStatus: 'needs_review',
      conflictsWith: [],
    },
  ],
}
function setup() {
  const record = {
    version: 1,
    documentId: 'doc',
    projectId: 'project',
    id: 'research1',
    sequence: 1,
    draftDigest: createHash('sha256').update(canonical(draft)).digest('hex'),
    draft,
    state: 'completed',
    startedAt: '2026-09-29T00:00:00.000Z',
    finishedAt: '2026-09-29T00:01:00.000Z',
    sources: [
      {
        sourceId: 'source1',
        attachmentId: 'a'.repeat(64),
        status: 'found',
        offset: 0,
        provenance: 'user_supplied',
        sha256: 'a'.repeat(64),
        parsedTextSha256: 'b'.repeat(64),
      },
    ],
    checks: {
      scope: 'research_draft',
      support: 'not_verified',
      sourceAuthority: 'not_verified',
      timeliness: 'not_verified',
    },
  }
  const history = {
    version: 1,
    documentId: 'doc',
    projectId: 'project',
    revision: 2,
    totalRecords: 1,
    records: [
      {
        id: record.id,
        sequence: 1,
        draftDigest: record.draftDigest,
        state: record.state,
        startedAt: record.startedAt,
        finishedAt: record.finishedAt,
        sourceCount: 1,
        factCount: 1,
        conflictCount: 0,
      },
    ],
  }
  const request = vi.fn(async (body: unknown) => {
    const op = (body as { operation: string }).operation
    return Response.json(
      op === 'research_capabilities'
        ? { version: 1, available: true }
        : op === 'research_build'
          ? { history, record }
          : op === 'research_list'
            ? history
            : op === 'research_latest'
              ? { record }
              : record,
    )
  })
  let document = 'doc'
  const vfs = new InMemoryVfs(),
    rememberProject = vi.fn(async () => undefined),
    onChanged = vi.fn()
  const options = {
    available: () => true,
    request,
    documentId: async () => document,
    vfs,
    rememberProject,
    onChanged,
  }
  const skill = createPresentationResearchSkill(options)
  const call = (
    name = 'build_research_ledger',
    input: Record<string, unknown> = {
      project_id: 'project',
      ledger_id: 'research1',
      expected_revision: 0,
      draft,
    },
    signal?: AbortSignal,
  ) => skill.executeTool({ id: 'x', name, input }, signal)
  return {
    record,
    history,
    request,
    vfs,
    rememberProject,
    onChanged,
    skill,
    call,
    options,
    switchDocument: () => {
      document = 'other'
    },
  }
}
describe('independent pre-plan research ledger tools', () => {
  it('builds and restores research without a saved plan or production, preserving unverified judgments', async () => {
    const f = setup(),
      result = await f.call()
    expect(result.isError, result.output).toBeFalsy()
    expect(JSON.parse(result.output)).toEqual({ history: f.history, record: f.record })
    expect(f.rememberProject).toHaveBeenCalledWith('project')
    expect(f.onChanged).toHaveBeenCalledWith('project')
    expect(
      f.request.mock.calls.every(([body]) =>
        String((body as { operation: string }).operation).startsWith('research_'),
      ),
    ).toBe(true)
    expect(await f.skill.latest('project')).toEqual(f.record)
    expect(
      (await f.call('read_research_ledger', { project_id: 'project', ledger_id: 'research1' }))
        .isError,
    ).toBeFalsy()
  })
  it('exports full JSON and readable source, inference and limitation details atomically', async () => {
    const f = setup(),
      result = await f.call('export_research_ledger', {
        project_id: 'project',
        ledger_id: 'research1',
      })
    expect(result.isError, result.output).toBeFalsy()
    const { paths } = JSON.parse(result.output)
    expect(JSON.parse(f.vfs.readText(paths[0]))).toEqual(f.record)
    const md = f.vfs.readText(paths[1])
    expect(md).toContain('判断')
    expect(md).toContain('原始证据有局限')
    expect(md).toContain('未核验')
  })
  it('rejects invalid fact references and unsupported PC before recording or remembering', async () => {
    const f = setup()
    expect(
      (
        await f.call(undefined, {
          project_id: 'project',
          ledger_id: 'research1',
          expected_revision: 0,
          draft: { ...draft, facts: [{ ...draft.facts[0], sourceRefs: ['unknown'] }] },
        })
      ).output,
    ).toBe('invalid_tool_input')
    expect(f.request).not.toHaveBeenCalled()
    f.request.mockImplementation(async () => Response.json({ error: 'invalid_request' }))
    expect((await f.call()).output).toBe('presentation_upgrade_required')
    expect(f.rememberProject).not.toHaveBeenCalled()
    expect(await f.skill.latest('project')).toBeUndefined()
  })
  it('guards cancellation, clear and document switches while PC saves or answers', async () => {
    for (const cause of ['clear', 'abort', 'doc']) {
      const f = setup(),
        controller = new AbortController(),
        actual = f.request.getMockImplementation()!
      f.request.mockImplementation(async (body) => {
        const response = await actual(body)
        if ((body as { operation: string }).operation === 'research_build') {
          if (cause === 'clear') f.skill.clear()
          else if (cause === 'abort') controller.abort()
          else f.switchDocument()
        }
        return response
      })
      expect((await f.call(undefined, undefined, controller.signal)).output).toBe(
        cause === 'doc' ? 'presentation_document_changed' : 'cancelled',
      )
      expect(f.rememberProject).not.toHaveBeenCalled()
      expect(f.vfs.list('/home/user')).toEqual([])
    }
  })
  it('rejects wrong record identity and machine/human factual verification claims', async () => {
    for (const bad of [
      { ...setup().record, documentId: 'other' },
      {
        ...setup().record,
        checks: {
          scope: 'research_draft',
          support: 'verified',
          sourceAuthority: 'not_verified',
          timeliness: 'not_verified',
        },
      },
    ]) {
      const f = setup(),
        actual = f.request.getMockImplementation()!
      f.request.mockImplementation(async (body) =>
        (body as { operation: string }).operation === 'research_read'
          ? Response.json(bad)
          : actual(body),
      )
      expect(
        (await f.call('export_research_ledger', { project_id: 'project', ledger_id: 'research1' }))
          .isError,
      ).toBe(true)
      expect(f.vfs.list('/home/user')).toEqual([])
    }
  })
  it('keeps research readable on PC when session attachment quota is exhausted', async () => {
    const f = setup(),
      skill = createPresentationResearchSkill({
        ...f.options,
        vfs: new InMemoryVfs({ maxFiles: 1 }),
      })
    expect(
      (
        await skill.executeTool({
          id: 'full',
          name: 'export_research_ledger',
          input: { project_id: 'project', ledger_id: 'research1' },
        })
      ).output,
    ).toBe('presentation_session_storage_full')
    expect(await skill.latest('project')).toEqual(f.record)
  })
})
it('accepts exact archived build retries outside the summary window without new research', async () => {
  const f = setup(),
    base = f.history.records[0]!
  Object.assign(f.history, {
    revision: 66,
    totalRecords: 33,
    records: Array.from({ length: 32 }, (_, i) => ({
      ...base,
      id: `research${i + 2}`,
      sequence: i + 2,
    })),
  })
  const result = await f.call()
  expect(result.isError, result.output).toBeFalsy()
  expect(JSON.parse(result.output).record).toEqual(f.record)
  expect(
    f.request.mock.calls.filter(
      ([body]) => (body as { operation: string }).operation === 'research_build',
    ),
  ).toHaveLength(1)
})
it('rejects a forged draft digest in a read even when identity and shape are valid', async () => {
  const f = setup(),
    actual = f.request.getMockImplementation()!
  f.request.mockImplementation(async (body) =>
    (body as { operation: string }).operation === 'research_read'
      ? Response.json({ ...f.record, draftDigest: 'f'.repeat(64) })
      : actual(body),
  )
  expect(
    (await f.call('read_research_ledger', { project_id: 'project', ledger_id: 'research1' }))
      .output,
  ).toBe('presentation_response_invalid')
})
