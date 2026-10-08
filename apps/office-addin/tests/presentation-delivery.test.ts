import { describe, expect, it, vi } from 'vitest'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import {
  createPresentationDeliverySkill,
  type PresentationImportRecord,
} from '../src/skills/powerpoint/presentation-delivery.js'

function fixture() {
  const receipts = new Map<string, PresentationImportRecord>()
  const before = { slideIds: ['original'], fingerprint: '["original"]' }
  const adapter = {
    available: () => true,
    snapshot: vi.fn(async () => before),
    insert: vi.fn(async () => ({ slideIds: ['new-1'] })),
    verify: vi.fn(async () => true),
  }
  const proposals = createStructuredProposalController()
  const documentId = vi.fn(async () => 'doc-1')
  const writeReceipt = vi.fn(async (key: string, record: PresentationImportRecord | undefined) => {
    if (record) receipts.set(key, record)
    else receipts.delete(key)
  })
  const artifact = {
    documentId: 'doc-1',
    projectId: 'project-1',
    requestId: 'request-1',
    pptxBase64: 'UEsDBAAAAAA=',
    slideCount: 1,
  }
  const skill = createPresentationDeliverySkill({
    adapter,
    proposals,
    available: () => true,
    artifact: () => artifact,
    documentId,
    readReceipt: (key) => receipts.get(key),
    writeReceipt,
  })
  const call = {
    id: 'import',
    name: 'import_generated_presentation',
    input: { project_id: 'project-1' },
  }
  return { receipts, adapter, proposals, documentId, writeReceipt, skill, call, artifact }
}

describe('generated deck delivery', () => {
  it('requires one confirmed append, saves pending before host write, and makes retries idempotent', async () => {
    const f = fixture()
    await f.skill.executeTool(f.call)
    expect(f.adapter.insert).not.toHaveBeenCalled()
    f.adapter.insert.mockImplementation(async () => {
      expect([...f.receipts.values()][0]?.state).toBe('pending')
      expect([...f.receipts.values()][0]?.toolCallId).toBe('import')
      return { slideIds: ['new-1'] }
    })
    await f.proposals.confirm(f.proposals.pending()!.id)
    expect([...f.receipts.values()][0]).toMatchObject({
      state: 'complete',
      toolCallId: 'import',
      slideIds: ['new-1'],
    })
    expect(await f.skill.executeTool(f.call)).toMatchObject({
      mutated: false,
      output: expect.stringContaining('already_imported'),
    })
    expect(f.adapter.insert).toHaveBeenCalledOnce()
  })
  it('blocks replay after an uncertain write, preserving existing content', async () => {
    const f = fixture()
    f.adapter.insert.mockRejectedValue(new Error('office_state_uncertain'))
    await f.skill.executeTool(f.call)
    await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow(
      'office_state_uncertain',
    )
    expect(await f.skill.executeTool(f.call)).toMatchObject({
      isError: true,
      output: 'presentation_import_uncertain',
    })
    expect(f.adapter.insert).toHaveBeenCalledOnce()
  })
  it('does not write if document identity changes before confirmation', async () => {
    const f = fixture()
    await f.skill.executeTool(f.call)
    f.documentId.mockResolvedValue('doc-2')
    await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('proposal_stale')
    expect(f.adapter.insert).not.toHaveBeenCalled()
  })
  it('rechecks durable receipts at confirmation to prevent two pending proposals from importing twice', async () => {
    const f = fixture()
    await f.skill.executeTool(f.call)
    await f.writeReceipt('project-1/request-1', {
      state: 'complete',
      documentId: 'doc-1',
      slideIds: ['new-1'],
    })
    await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow('proposal_stale')
    expect(f.adapter.insert).not.toHaveBeenCalled()
  })
  it('does not report import complete when readback fails', async () => {
    const f = fixture()
    f.adapter.verify.mockResolvedValue(false)
    await f.skill.executeTool(f.call)
    await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow(
      'office_state_uncertain',
    )
    expect([...f.receipts.values()][0]?.state).toBe('pending')
  })
})

describe('import binding review regressions', () => {
  it('accepts long document locations without exceeding proposal metadata bounds', async () => {
    const f = fixture()
    const long = 'x'.repeat(1000)
    f.documentId.mockResolvedValue(long)
    // The fixture's generated artifact belongs to this document.
    f.artifact.documentId = long
    expect(await f.skill.executeTool(f.call)).not.toMatchObject({ isError: true })
    expect(f.proposals.pending()!.fingerprint.length).toBeLessThanOrEqual(512)
  })
  it('does not import an artifact compiled for a different document', async () => {
    const f = fixture()
    f.documentId.mockResolvedValue('doc-after-save-as')
    expect(await f.skill.executeTool(f.call)).toMatchObject({
      isError: true,
      output: 'presentation_document_changed',
    })
    expect(f.proposals.pending()).toBeUndefined()
    expect(f.adapter.insert).not.toHaveBeenCalled()
  })
})
