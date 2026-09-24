import { describe, expect, it, vi } from 'vitest'
import {
  createPresentationAgentRunCheckpoint,
  createPresentationDocumentBinding,
} from '../src/skills/powerpoint/presentation-document.js'

describe('presentation AgentRun checkpoint', () => {
  it('binds a new run to the Save As copy and finishes only that document', async () => {
    const values = new Map<string, string>()
    let location = 'file:///original.pptx'
    const binding = createPresentationDocumentBinding(
      {
        get: (key) => values.get(key),
        set: (key, value) => {
          values.set(key, value)
        },
        save: async () => undefined,
        location: () => location,
      },
      () => 'doc-id',
    )
    const originalId = await binding.documentId()
    const checkpoint = createPresentationAgentRunCheckpoint(binding)
    location = 'file:///copy.pptx'
    const copyId = await binding.documentId()
    await checkpoint.begin('run-1')
    expect(binding.interruptedAgentRun(copyId)).toBe(true)
    expect(binding.interruptedAgentRun(originalId)).toBe(false)
    location = 'file:///original.pptx'
    await expect(checkpoint.finish('run-1')).rejects.toThrow('presentation_document_changed')
    location = 'file:///copy.pptx'
    expect(binding.interruptedAgentRun(copyId)).toBe(true)
    await checkpoint.finish('run-1')
    expect(binding.interruptedAgentRun(copyId)).toBe(false)
  })

  it('detects a foreground run after reopen, but not in a Save As copy', async () => {
    const values = new Map<string, string>()
    let location = 'file:///original.pptx'
    const settings = {
      get: (key: string) => values.get(key),
      set: (key: string, value: string) => {
        values.set(key, value)
      },
      save: vi.fn(async () => undefined),
      location: () => location,
    }
    const create = () => createPresentationDocumentBinding(settings, () => 'doc-id')
    const first = create()
    const originalId = await first.documentId()
    await first.rememberAgentRun(originalId, 'run-1')
    expect(create().interruptedAgentRun(originalId)).toBe(true)
    location = 'file:///copy.pptx'
    expect(create().interruptedAgentRun(await create().documentId())).toBe(false)
    location = 'file:///original.pptx'
    await create().finishAgentRun(originalId, 'other-run')
    expect(create().interruptedAgentRun(originalId)).toBe(true)
    await create().finishAgentRun(originalId, 'run-1')
    expect(create().interruptedAgentRun(originalId)).toBe(false)
  })

  it('rolls back a checkpoint when document settings cannot save', async () => {
    const values = new Map<string, string>()
    const settings = {
      get: (key: string) => values.get(key),
      set: (key: string, value: string) => {
        values.set(key, value)
      },
      save: vi.fn(async () => undefined),
      location: () => 'file:///original.pptx',
    }
    const binding = createPresentationDocumentBinding(settings, () => 'doc-id')
    const id = await binding.documentId()
    settings.save.mockRejectedValueOnce(new Error('save failed'))
    await expect(binding.rememberAgentRun(id, 'run-1')).rejects.toThrow('save failed')
    expect(binding.interruptedAgentRun(id)).toBe(false)
  })
})
