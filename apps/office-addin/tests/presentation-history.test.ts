import { expect, it } from 'vitest'
import { createHash } from 'node:crypto'
import { createPresentationHistorySkill } from '../src/skills/powerpoint/presentation-history'
const artifact = {
  documentId: 'doc',
  projectId: 'project',
  requestId: 'request',
  pptxBase64: 'base64',
  slideCount: 1,
  pages: [{ id: 'page', title: 'Page', sourceSlideId: '256#' }],
}
const record = {
  version: 1 as const,
  changeId: 'first',
  documentId: 'doc',
  projectId: 'project',
  requestId: 'request',
  artifactDigest: createHash('sha256').update('base64').digest('hex'),
  pageId: 'page',
  hostSlideId: 'host',
  shapeId: 'shape',
  before: 'secret-before',
  after: 'secret-after',
  state: 'applied' as const,
}
const entry = { id: 'text:first', sequence: 1, legacy: false, kind: 'text' as const, record }
const call = { id: 'history', name: 'list_presentation_changes', input: { project_id: 'project' } }
it('lists scoped single-operation summaries without full saved text or passing QA claims', async () => {
  const skill = createPresentationHistorySkill({
    available: () => true,
    artifact: () => artifact,
    documentId: async () => 'doc',
    listChangeHistory: () => [
      entry,
      {
        ...entry,
        id: 'text:other',
        sequence: 2,
        record: { ...record, changeId: 'other', requestId: 'other' },
      },
    ],
  })
  expect(skill.tools.map((t) => t.name)).toEqual(['list_presentation_changes'])
  const result = await skill.executeTool(call)
  expect(result.isError, result.output).not.toBe(true)
  const out = JSON.parse(result.output)
  expect(out.changes).toHaveLength(1)
  expect(out.changes[0]).toMatchObject({
    kind: 'text',
    change_id: 'first',
    page_id: 'page',
    state: 'applied',
    changeSet: { scope: { slideIds: ['host'], shapeIds: ['shape'] }, risk: 'medium' },
  })
  expect(out.changes[0].changeSet.operations).toHaveLength(1)
  expect(out.changes[0].changeSet.validation).toContain('受影响页面截图复核')
  expect(result.output).not.toContain('secret')
  expect(out.currentHostVerified).toBe(false)
})
it('rejects duplicate/malformed history and invalid input rather than showing an empty history', async () => {
  let entries = [entry, entry]
  const skill = createPresentationHistorySkill({
    available: () => true,
    artifact: () => artifact,
    documentId: async () => 'doc',
    listChangeHistory: () => entries,
  })
  expect((await skill.executeTool(call)).isError).toBe(true)
  entries = [{ ...entry, sequence: -1 }]
  expect((await skill.executeTool(call)).isError).toBe(true)
  entries = [entry]
  expect((await skill.executeTool({ ...call, input: { project_id: 'other' } })).isError).toBe(true)
  expect(
    (await skill.executeTool({ ...call, input: { project_id: 'project', extra: true } })).output,
  ).toBe('invalid_tool_input')
})
it('clear suppresses late history reads', async () => {
  let release!: (id: string) => void
  const skill = createPresentationHistorySkill({
    available: () => true,
    artifact: () => artifact,
    documentId: () =>
      new Promise((r) => {
        release = r
      }),
    listChangeHistory: () => [entry],
  })
  const reading = skill.executeTool(call)
  skill.clear()
  release('doc')
  expect((await reading).isError).toBe(true)
})
it.each([
  { documentId: 'other' },
  { source: 'production' as const },
  { pageId: 'missing' },
  { artifactDigest: 'a'.repeat(64) },
])('hides records outside the current binding (%j)', async (patch) => {
  const skill = createPresentationHistorySkill({
    available: () => true,
    artifact: () => artifact,
    documentId: async () => 'doc',
    listChangeHistory: () => [{ ...entry, record: { ...record, ...patch } }],
  })
  const result = await skill.executeTool(call)
  expect(result.isError, result.output).not.toBe(true)
  expect(JSON.parse(result.output).changes).toEqual([])
})
it('suppresses a history response after switching the active task', async () => {
  let active = artifact,
    release!: (id: string) => void
  const skill = createPresentationHistorySkill({
    available: () => true,
    artifact: () => active,
    documentId: () =>
      new Promise((r) => {
        release = r
      }),
    listChangeHistory: () => [entry],
  })
  const reading = skill.executeTool(call)
  active = { ...artifact, requestId: 'other' }
  release('doc')
  expect((await reading).isError).toBe(true)
})
