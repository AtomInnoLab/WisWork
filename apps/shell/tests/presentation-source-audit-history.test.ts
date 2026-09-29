import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { expect, it } from 'vitest'
import { PresentationStore } from '@wiswork/project-store'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { createPresentationService } from '../src/main/presentation-service.js'
import { createPresentationProjectController } from '../../office-addin/src/skills/powerpoint/presentation-project.js'

it('records real attachment excerpt research, replays exact audit IDs and exposes bounded historical summaries after PC restart', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'source-audit-history-'))
  try {
    let service = createPresentationService({ userDataPath })
    const plan = benchmarkPlan(),
      bytes = Buffer.from('示例数据仅用于测试')
    const attachmentId = createHash('sha256').update(bytes).digest('hex')
    plan.sources[0]!.uri = `attachment:${attachmentId}`
    plan.sources[0]!.locator = '第 1 段'
    const call = async (operation: string, fields: Record<string, unknown> = {}) =>
      JSON.parse(
        Buffer.from(
          await service(
            {
              operation,
              documentId: 'doc',
              ...(operation.startsWith('attachment_') ? {} : { projectId: plan.projectId }),
              ...fields,
            },
            new AbortController().signal,
          ),
        ).toString('utf8'),
      )
    await call('save_plan', { expectedRevision: 0, plan })
    await call('attachment_begin', {
      attachmentId,
      sha256: attachmentId,
      name: 'study.txt',
      sizeBytes: bytes.length,
    })
    await call('attachment_chunk', { attachmentId, offset: 0, base64: bytes.toString('base64') })
    await call('attachment_finish', { attachmentId })
    const result = await call('audit_sources', { auditId: 'audit' })
    expect(result).toMatchObject({
      planRevision: 1,
      sources: [{ sourceId: plan.sources[0]!.id, status: 'found', offset: 0 }],
    })
    expect(result.checks).toEqual({
      support: 'not_verified',
      sourceAuthority: 'not_verified',
      timeliness: 'not_verified',
    })
    const openWorkbench = () =>
      createPresentationProjectController({
        request: async (body, signal) =>
          new Response(Buffer.from(await service(body, signal ?? new AbortController().signal))),
        available: () => true,
        lastProject: () => plan.projectId,
        documentId: async () => 'doc',
        executeTool: async () => ({ output: '', summary: '' }),
        rememberProject: async () => {},
      })
    service = createPresentationService({ userDataPath })
    const workbench = openWorkbench()
    await workbench.refresh()
    expect(workbench.snapshot().sourceAudit).toMatchObject({
      auditId: 'audit',
      planRevision: 1,
      sources: result.sources,
    })
    const restored = workbench.snapshot().sourceAudit
    workbench.clear()
    const reopened = openWorkbench()
    await reopened.refresh()
    expect(reopened.snapshot().sourceAudit).toEqual(restored)
    reopened.clear()
    await call('save_plan', { expectedRevision: 1, plan: { ...plan, title: 'new plan' } })
    service = createPresentationService({ userDataPath })
    expect(await call('audit_sources', { auditId: 'audit' })).toEqual(result)
    const changed = openWorkbench()
    await changed.refresh()
    expect(changed.snapshot().sourceAudit).toBeUndefined()
    expect(changed.snapshot().project?.sourceAuditHistory?.runs[0]?.id).toBe('audit')
    changed.clear()
    const status = await call('status')
    expect(status.sourceAuditHistory).toMatchObject({
      projectId: plan.projectId,
      documentId: 'doc',
      revision: 2,
      runs: [
        {
          id: 'audit',
          planRevision: 1,
          scope: 'source_excerpt_audit',
          state: 'completed',
          sourceCount: 1,
          foundCount: 1,
        },
      ],
    })
    expect(JSON.stringify(status.sourceAuditHistory)).not.toContain('示例数据仅用于测试')
    const read = await call('read_source_audit', { auditId: 'audit' })
    expect(read.audit.sources).toEqual(result.sources)
    expect(read).toMatchObject({ documentId: 'doc', projectId: plan.projectId })
    expect(await call('read_source_audit', { auditId: 'audit', documentId: 'foreign' })).toEqual({
      error: 'document_mismatch',
    })
    // Existing callers need no additional field and produce a new current-plan check.
    expect((await call('audit_sources')).planRevision).toBe(2)
    expect((await call('status')).sourceAuditHistory.runs).toHaveLength(2)
  } finally {
    rmSync(userDataPath, { recursive: true, force: true })
  }
})
it('keeps a started-but-unfinished read explicit and records malformed attachment failures without inventing a completed research result', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'source-audit-interrupted-'))
  try {
    const plan = benchmarkPlan()
    const bytes = Buffer.from('示例数据仅用于测试'),
      attachmentId = createHash('sha256').update(bytes).digest('hex')
    plan.sources[0]!.uri = `attachment:${attachmentId}`
    plan.sources[0]!.locator = '第 1 段'
    const store = new PresentationStore(userDataPath)
    store.savePlan(plan.projectId, 'doc', 0, plan)
    store.beginSourceAudit(plan.projectId, 'doc', 'interrupted')
    const service = createPresentationService({ userDataPath })
    for (const fields of [
      {
        operation: 'attachment_begin',
        attachmentId,
        sha256: attachmentId,
        name: 'study.txt',
        sizeBytes: bytes.length,
      },
      { operation: 'attachment_chunk', attachmentId, offset: 0, base64: bytes.toString('base64') },
      { operation: 'attachment_finish', attachmentId },
    ])
      await service({ documentId: 'doc', ...fields }, new AbortController().signal)
    writeFileSync(
      join(
        userDataPath,
        'presentation-attachments',
        createHash('sha256').update('doc').digest('hex'),
        attachmentId,
        'text.txt',
      ),
      'tampered',
    )
    const result = JSON.parse(
      Buffer.from(
        await service(
          {
            operation: 'audit_sources',
            documentId: 'doc',
            projectId: plan.projectId,
            auditId: 'failed',
          },
          new AbortController().signal,
        ),
      ).toString('utf8'),
    )
    expect(result).toEqual({ error: 'invalid_state' })
    expect(store.sourceAudits(plan.projectId, 'doc').runs).toMatchObject([
      { id: 'interrupted', state: 'running' },
      { id: 'failed', state: 'failed', error: 'invalid_state' },
    ])
  } finally {
    rmSync(userDataPath, { recursive: true, force: true })
  }
})
