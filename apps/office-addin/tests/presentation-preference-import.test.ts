import { expect, it, vi } from 'vitest'
import { createPresentationPlanningSkill } from '../src/skills/powerpoint/presentation-planning'
import { createStructuredProposalController } from '../src/agent/proposal-controller'
import { InMemoryVfs } from '../src/skills/shared/vfs'

const source = { documentId: 'source-doc', projectId: 'source-project', changeId: 'edit-1' }
const input = {
  source_document_id: source.documentId,
  source_project_id: source.projectId,
  source_change_id: source.changeId,
  project_id: 'target-project',
}
const digest = async (text: string) =>
  Array.from(
    new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text))),
    (x) => x.toString(16).padStart(2, '0'),
  ).join('')
function fixture() {
  let sourcePreference: Record<string, unknown> | undefined = {
    projectId: source.projectId,
    changeId: source.changeId,
    text: '保留 原文空格与中文',
  }
  let document = 'target-doc',
    available = true
  const proposals = createStructuredProposalController()
  const request = vi.fn(async (body: unknown) => {
    const call = body as Record<string, unknown>
    if (call.operation === 'preference_get')
      return new Response(JSON.stringify({ preference: sourcePreference }))
    if (call.operation === 'preference_import')
      return new Response(
        JSON.stringify({
          preference: {
            projectId: call.projectId,
            changeId: `reuse_${await digest(JSON.stringify([call.documentId, call.projectId, source.documentId, source.projectId, source.changeId, call.expectedTextDigest]))}`,
            text: sourcePreference!.text,
            reuse: {
              version: 1,
              source,
              sourceTextDigest: call.expectedTextDigest,
              approvedAt: '2026-09-29T00:00:00.000Z',
              approvalId: call.approvalId,
            },
          },
        }),
      )
    return new Response(JSON.stringify({ preferences: [] }))
  })
  const skill = createPresentationPlanningSkill({
    vfs: new InMemoryVfs(),
    proposals,
    request,
    documentId: async () => document,
    available: () => available,
    lastProject: () => 'target-project',
    rememberProject: async () => {},
  })
  const propose = () =>
    skill.executeTool({ id: 'import', name: 'import_presentation_preference', input })
  return {
    skill,
    request,
    proposals,
    propose,
    setSource: (value: Record<string, unknown> | undefined) => {
      sourcePreference = value
    },
    setDocument: (value: string) => {
      document = value
    },
    setAvailable: (value: boolean) => {
      available = value
    },
  }
}
it('offers a visible cross-project copy proposal and imports only after confirmation', async () => {
  const f = fixture()
  expect(f.skill.tools.some((t) => t.name === 'import_presentation_preference')).toBe(true)
  const result = await f.propose()
  expect(result.isError).not.toBe(true)
  expect(
    f.request.mock.calls.every(
      ([body]) => (body as Record<string, unknown>).operation === 'preference_get',
    ),
  ).toBe(true)
  const p = f.proposals.pending()!
  expect(p.preview).toMatchObject({
    source,
    target: { documentId: 'target-doc', projectId: 'target-project' },
  })
  const decision = f.proposals.waitForDecision(p.id)
  await f.proposals.confirm(p.id)
  expect(await decision).toEqual({ status: 'confirmed' })
  expect(f.request).toHaveBeenCalledWith(
    {
      operation: 'preference_import',
      documentId: 'target-doc',
      projectId: 'target-project',
      source,
      expectedTextDigest: await digest('保留 原文空格与中文'),
      approvalId: p.id,
    },
    expect.any(AbortSignal),
  )
})
it.each(['source-changed', 'source-missing', 'document-changed', 'availability-lost', 'cleared'])(
  'rejects confirmation after %s without importing',
  async (scenario) => {
    const f = fixture()
    await f.propose()
    const p = f.proposals.pending()!
    if (scenario === 'source-changed')
      f.setSource({ projectId: source.projectId, changeId: source.changeId, text: 'Changed' })
    if (scenario === 'source-missing') f.setSource(undefined)
    if (scenario === 'document-changed') f.setDocument('other-doc')
    if (scenario === 'availability-lost') f.setAvailable(false)
    if (scenario === 'cleared') f.skill.clear()
    await expect(f.proposals.confirm(p.id)).rejects.toThrow()
    expect(
      f.request.mock.calls.some(
        ([body]) => (body as Record<string, unknown>).operation === 'preference_import',
      ),
    ).toBe(false)
  },
)

it.each(['project', 'change', 'text', 'source', 'digest', 'approval', 'date', 'extra'])(
  'rejects a malformed or substituted %s import receipt',
  async (scenario) => {
    const f = fixture(),
      original = f.request.getMockImplementation()!
    f.request.mockImplementation(async (body) => {
      const response = await original(body)
      if ((body as Record<string, unknown>).operation !== 'preference_import') return response
      const value = await response.json(),
        p = value.preference
      if (scenario === 'project') p.projectId = 'foreign'
      if (scenario === 'change') p.changeId = 'foreign'
      if (scenario === 'text') p.text = 'Changed'
      if (scenario === 'source') p.reuse.source.documentId = 'foreign'
      if (scenario === 'digest') p.reuse.sourceTextDigest = 'f'.repeat(64)
      if (scenario === 'approval') p.reuse.approvalId = 'not-a-uuid'
      if (scenario === 'date') p.reuse.approvedAt = '2026-02-30T00:00:00.000Z'
      if (scenario === 'extra') p.extra = 'hidden'
      return new Response(JSON.stringify(value))
    })
    await f.propose()
    const p = f.proposals.pending()!
    await expect(f.proposals.confirm(p.id)).rejects.toThrow('presentation_response_invalid')
    expect(
      f.request.mock.calls.filter(
        ([body]) => (body as Record<string, unknown>).operation === 'preference_import',
      ),
    ).toHaveLength(1)
  },
)
it('accepts an idempotent existing copy with its valid first approval rather than inventing a new one', async () => {
  const f = fixture(),
    original = f.request.getMockImplementation()!
  f.request.mockImplementation(async (body) => {
    const response = await original(body)
    if ((body as Record<string, unknown>).operation !== 'preference_import') return response
    const value = await response.json()
    value.preference.reuse.approvalId = '11111111-1111-1111-1111-111111111111'
    value.preference.reuse.approvedAt = '2025-01-01T00:00:00.000Z'
    return new Response(JSON.stringify(value))
  })
  await f.propose()
  const p = f.proposals.pending()!,
    decision = f.proposals.waitForDecision(p.id)
  expect(p.preview.note).toContain('首次批准')
  await f.proposals.confirm(p.id)
  expect(await decision).toEqual({ status: 'confirmed' })
})
it('refuses forwarding a previously imported preference before creating a proposal', async () => {
  const f = fixture()
  f.setSource({
    projectId: source.projectId,
    changeId: source.changeId,
    text: 'Copied',
    reuse: {
      version: 1,
      source: { documentId: 'origin', projectId: 'origin', changeId: 'edit' },
      sourceTextDigest: await digest('Copied'),
      approvedAt: '2026-09-29T00:00:00.000Z',
      approvalId: '11111111-1111-1111-1111-111111111111',
    },
  })
  expect(await f.propose()).toMatchObject({
    isError: true,
    output: 'presentation_preference_import_unavailable',
  })
  expect(f.proposals.pending()).toBeUndefined()
  expect(f.request.mock.calls).toHaveLength(1)
})
it('retains imported provenance through listing and exact confirmed deletion while accepting legacy items', async () => {
  const f = fixture(),
    original = f.request.getMockImplementation()!
  const item = {
    projectId: 'target-project',
    changeId: 'reuse_test',
    text: 'Imported literal',
    reuse: {
      version: 1,
      source,
      sourceTextDigest: await digest('Imported literal'),
      approvedAt: '2026-09-29T00:00:00.000Z',
      approvalId: '11111111-1111-1111-1111-111111111111',
    },
  }
  let preferences = [item, { projectId: 'target-project', changeId: 'legacy', text: 'Legacy' }]
  f.request.mockImplementation(async (body) => {
    const call = body as Record<string, unknown>
    if (call.operation === 'preference_list') return new Response(JSON.stringify({ preferences }))
    if (call.operation === 'preference_delete') {
      preferences = preferences.filter((p) => p.changeId !== call.changeId)
      return new Response(JSON.stringify({ deleted: true }))
    }
    return original(body)
  })
  const listed = await f.skill.executeTool({
    id: 'list',
    name: 'list_presentation_preferences',
    input: { project_id: 'target-project' },
  })
  expect(JSON.parse(listed.output)).toEqual({ preferences })
  const proposed = await f.skill.executeTool({
    id: 'delete',
    name: 'delete_presentation_preference',
    input: { project_id: 'target-project', change_id: item.changeId },
  })
  expect(proposed.isError).not.toBe(true)
  expect(preferences).toHaveLength(2)
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(preferences).toEqual([{ projectId: 'target-project', changeId: 'legacy', text: 'Legacy' }])
})
it('drops malformed imported metadata rather than returning a stripped legacy item', async () => {
  const f = fixture()
  f.request.mockResolvedValue(
    new Response(
      JSON.stringify({
        preferences: [
          {
            projectId: 'target-project',
            changeId: 'reuse_test',
            text: 'Copy',
            reuse: { version: 1 },
          },
        ],
      }),
    ),
  )
  expect(
    await f.skill.executeTool({
      id: 'list',
      name: 'list_presentation_preferences',
      input: { project_id: 'target-project' },
    }),
  ).toMatchObject({ isError: true, output: 'presentation_response_invalid' })
})

it.each(['source-request', 'source-body'])(
  'rejects a source read when the document changes during %s',
  async (phase) => {
    const f = fixture(),
      original = f.request.getMockImplementation()!
    f.request.mockImplementation(async (body) => {
      const response = await original(body)
      if (phase === 'source-request') f.setDocument('foreign')
      else {
        const value = await response.json()
        response.json = async () => {
          f.setDocument('foreign')
          return value
        }
      }
      return response
    })
    expect(await f.propose()).toMatchObject({
      isError: true,
      output: 'presentation_document_changed',
    })
    expect(f.proposals.pending()).toBeUndefined()
    expect(f.request.mock.calls).toHaveLength(1)
  },
)

it('rejects a same-document same-project copy before any source request or confirmation', async () => {
  const f = fixture()
  expect(
    await f.skill.executeTool({
      id: 'same',
      name: 'import_presentation_preference',
      input: { ...input, source_document_id: 'target-doc', source_project_id: 'target-project' },
    }),
  ).toMatchObject({ isError: true, output: 'invalid_tool_input' })
  expect(f.request).not.toHaveBeenCalled()
  expect(f.proposals.pending()).toBeUndefined()
})
