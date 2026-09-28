import { deliveryReportFixture } from './presentation-delivery-fixture.js'
import { describe, expect, it, vi } from 'vitest'
import { createPresentationEvidenceDeliverySkill } from '../src/skills/powerpoint/presentation-evidence-delivery.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'

function fixture(value: unknown = {}) {
  const request = vi.fn(async () => new Response(JSON.stringify(value)))
  const documentId = vi.fn(async () => 'd')
  const vfs = new InMemoryVfs()
  const skill = createPresentationEvidenceDeliverySkill({
    request,
    documentId,
    vfs,
    available: () => true,
  })
  const call = (
    name = 'read_presentation_delivery_report',
    input: Record<string, unknown> = { project_id: 'p', request_id: 'r' },
    signal?: AbortSignal,
  ) => skill.executeTool({ id: 'x', name, input }, signal)
  return { request, documentId, vfs, skill, call }
}
describe('evidence delivery boundary', () => {
  it('maps old PC rejection and rejects malformed reports without exporting', async () => {
    expect((await fixture({ error: 'invalid_request' }).call()).output).toBe(
      'presentation_upgrade_required',
    )
    const f = fixture({ checks: { content: 'passed' } })
    expect((await f.call('export_presentation_delivery_report')).isError).toBe(true)
    expect(f.vfs.list('/home/user')).toEqual([])
  })
  it('rejects unknown fields, malformed actions and absent CAS before requesting', async () => {
    const f = fixture()
    for (const input of [
      { project_id: '../p', request_id: 'r' },
      { project_id: 'p', request_id: 'r', hidden: true },
    ])
      expect((await f.call(undefined, input)).isError).toBe(true)
    expect((await f.call('record_presentation_issue_action')).isError).toBe(true)
    expect(f.request).not.toHaveBeenCalled()
  })
  it('invalidates clear and abort while PC response is pending', async () => {
    for (const abort of [false, true]) {
      const f = fixture()
      const controller = new AbortController()
      f.request.mockImplementation(async () => {
        if (abort) controller.abort()
        else f.skill.clear()
        return new Response('{}')
      })
      expect((await f.call(undefined, undefined, controller.signal)).output).toBe('cancelled')
      expect(f.vfs.list('/home/user')).toEqual([])
    }
  })
  it('bounds the entire report in UTF-8 bytes', async () => {
    const f = fixture('中'.repeat(3 * 1024 * 1024))
    expect((await f.call()).output).toBe('presentation_response_invalid')
  })
})

it('reads full evidence, validates identity, and atomically exports complete JSON and Markdown', async () => {
  const report = await deliveryReportFixture()
  const input = { project_id: report.projectId, request_id: report.requestId }
  const f = fixture(report)
  expect(JSON.parse((await f.call(undefined, input)).output)).toEqual(report)
  const exported = await f.call('export_presentation_delivery_report', input)
  expect(exported.isError).toBeFalsy()
  const { paths } = JSON.parse(exported.output)
  expect(JSON.parse(f.vfs.readText(paths[0]))).toEqual(report)
  expect(f.vfs.readText(paths[1])).toContain(report.requestId)
  expect((await f.call()).isError).toBe(true)
  f.documentId.mockResolvedValueOnce('d').mockResolvedValue('other')
  expect((await f.call(undefined, input)).output).toBe('presentation_document_changed')
  const vfs = new InMemoryVfs({ maxFiles: 1 })
  const skill = createPresentationEvidenceDeliverySkill({
    request: async () => new Response(JSON.stringify(report)),
    documentId: async () => 'd',
    available: () => true,
    vfs,
  })
  expect(
    await skill.executeTool({ id: 'x', name: 'export_presentation_delivery_report', input }),
  ).toMatchObject({
    isError: true,
    output: 'presentation_session_storage_full',
    summary: expect.stringContaining('会话附件空间不足'),
  })
  expect(vfs.list('/home/user')).toEqual([])
})
it('sends validated issue action and preserves stale/CAS server errors', async () => {
  const report = await deliveryReportFixture()
  const issue = report.pages.flatMap((page) => page.issues)[0]!
  const action = {
    actionId: 'action1',
    issueId: issue.id,
    issueDigest: issue.digest,
    state: 'deferred',
    note: 'Await source',
  }
  for (const error of ['revision_conflict', 'issue_changed']) {
    const f = fixture({ error })
    expect(
      (
        await f.call('record_presentation_issue_action', {
          project_id: report.projectId,
          request_id: 'r',
          expected_revision: 0,
          action,
        })
      ).output,
    ).toBe(`presentation_${error}`)
    expect(f.request).toHaveBeenCalledWith(
      expect.objectContaining({
        operation: 'production_record_issue_action',
        expectedRevision: 0,
        action,
      }),
      undefined,
    )
  }
})
