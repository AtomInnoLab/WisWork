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
    location = 'file:///copy.pptx'
    const copyId = await binding.documentId()
    const checkpoint = createPresentationAgentRunCheckpoint(binding, copyId)
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

  it('never starts an old request when the document changes after resume validation', async () => {
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
    const checkpoint = createPresentationAgentRunCheckpoint(binding, originalId)
    location = 'file:///copy.pptx'
    await expect(checkpoint.begin('run-1', 'Old request')).rejects.toThrow(
      'presentation_document_changed',
    )
    expect(binding.agentRunRecovery(await binding.documentId())).toBeUndefined()
  })

  it('keeps recoverable instructions in local storage, isolated by document and run', async () => {
    const values = new Map<string, string>()
    const local = new Map<string, string>()
    const storage = {
      get length() {
        return local.size
      },
      key: (index: number) => [...local.keys()][index] ?? null,
      getItem: (key: string) => local.get(key) ?? null,
      setItem: (key: string, value: string) => {
        local.set(key, value)
      },
      removeItem: (key: string) => {
        local.delete(key)
      },
    }
    const binding = createPresentationDocumentBinding(
      {
        get: (key) => values.get(key),
        set: (key, value) => {
          values.set(key, value)
        },
        save: async () => undefined,
        location: () => 'file:///deck.pptx',
      },
      () => 'doc-id',
    )
    const id = await binding.documentId()
    const checkpoint = createPresentationAgentRunCheckpoint(binding, id, storage)
    await checkpoint.begin('run-1', 'Private request')
    expect(values.get('wiswork.presentation.agent-run.v1')).not.toContain('Private request')
    expect(createPresentationAgentRunCheckpoint(binding, id, storage).recovery()?.instruction).toBe(
      'Private request',
    )
    expect(
      createPresentationAgentRunCheckpoint(binding, 'foreign-document', storage).recovery(),
    ).toBeUndefined()
    const [localKey] = local.keys()
    const stored = JSON.parse(local.get(localKey)!)
    local.set(localKey, JSON.stringify({ ...stored, expiresAt: Date.now() - 1 }))
    expect(createPresentationAgentRunCheckpoint(binding, id, storage).recovery()?.instruction).toBe(
      '',
    )
    expect(local.size).toBe(0)
    await checkpoint.finish('run-1')
    expect(local.size).toBe(0)
  })

  it('sweeps expired orphan prompts when a different deck initializes', async () => {
    const values = new Map<string, string>()
    const local = new Map<string, string>([
      [
        'wiswork.presentation.agent-run.prompt.v1.orphan',
        JSON.stringify({
          documentId: 'old',
          runId: 'orphan',
          instruction: 'Secret',
          expiresAt: Date.now() - 1,
        }),
      ],
      ['unrelated', 'preserve'],
    ])
    const storage = {
      get length() {
        return local.size
      },
      key: (index: number) => [...local.keys()][index] ?? null,
      getItem: (key: string) => local.get(key) ?? null,
      setItem: (key: string, value: string) => {
        local.set(key, value)
      },
      removeItem: (key: string) => {
        local.delete(key)
      },
    }
    const binding = createPresentationDocumentBinding(
      {
        get: (key) => values.get(key),
        set: (key, value) => {
          values.set(key, value)
        },
        save: async () => undefined,
        location: () => 'file:///new.pptx',
      },
      () => 'doc-id',
    )
    createPresentationAgentRunCheckpoint(binding, await binding.documentId(), storage)
    expect(local.has('wiswork.presentation.agent-run.prompt.v1.orphan')).toBe(false)
    expect(local.get('unrelated')).toBe('preserve')
  })

  it('scrubs a legacy embedded prompt on load while retaining the running checkpoint', async () => {
    const values = new Map<string, string>()
    const save = vi.fn(async () => undefined)
    const binding = createPresentationDocumentBinding(
      {
        get: (key) => values.get(key),
        set: (key, value) => {
          values.set(key, value)
        },
        save,
        location: () => 'file:///old.pptx',
      },
      () => 'doc-id',
    )
    const id = await binding.documentId()
    values.set(
      'wiswork.presentation.agent-run.v1',
      JSON.stringify({
        documentId: id,
        runId: 'run-1',
        startedAt: Date.now(),
        instruction: 'Old secret',
        phase: 'tool_pending',
        toolName: 'write_page',
      }),
    )
    await binding.scrubAgentRunPrompt(id)
    expect(values.get('wiswork.presentation.agent-run.v1')).not.toContain('Old secret')
    expect(binding.agentRunRecovery(id)).toMatchObject({
      runId: 'run-1',
      phase: 'tool_pending',
      toolName: 'write_page',
    })
    expect(save).toHaveBeenCalledTimes(2)
  })

  it('does not treat a failed legacy-prompt scrub as successful', async () => {
    const values = new Map<string, string>()
    const save = vi.fn(async () => undefined)
    const binding = createPresentationDocumentBinding(
      {
        get: (key) => values.get(key),
        set: (key, value) => {
          values.set(key, value)
        },
        save,
        location: () => 'file:///old.pptx',
      },
      () => 'doc-id',
    )
    const id = await binding.documentId()
    const raw = JSON.stringify({
      documentId: id,
      runId: 'run-1',
      startedAt: Date.now(),
      instruction: 'Old secret',
      phase: 'running',
    })
    values.set('wiswork.presentation.agent-run.v1', raw)
    save.mockRejectedValueOnce(new Error('save failed'))
    await expect(binding.scrubAgentRunPrompt(id)).rejects.toThrow('save failed')
    expect(values.get('wiswork.presentation.agent-run.v1')).toBe(raw)
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

  it('recovers a bounded request and tool boundary only in the bound document', async () => {
    const values = new Map<string, string>()
    let location = 'file:///original.pptx'
    const create = () =>
      createPresentationDocumentBinding(
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
    const binding = create()
    const id = await binding.documentId()
    await binding.rememberAgentRun(id, 'run-1')
    await binding.updateAgentRun(id, 'run-1', 'tool_pending', 'write_presentation_page')
    expect(create().agentRunRecovery(id)).toMatchObject({
      instruction: '',
      phase: 'tool_pending',
      toolName: 'write_presentation_page',
    })
    location = 'file:///copy.pptx'
    expect(create().agentRunRecovery(await create().documentId())).toBeUndefined()
    location = 'file:///original.pptx'
    await binding.updateAgentRun(id, 'run-1', 'tool_completed', 'write_presentation_page')
    expect(create().agentRunRecovery(id)?.phase).toBe('tool_completed')
  })

  it('rejects malformed or oversized run records', async () => {
    const values = new Map<string, string>()
    const binding = createPresentationDocumentBinding(
      {
        get: (key) => values.get(key),
        set: (key, value) => {
          values.set(key, value)
        },
        save: async () => undefined,
        location: () => 'file:///deck.pptx',
      },
      () => 'doc-id',
    )
    const id = await binding.documentId()
    const key = 'wiswork.presentation.agent-run.v1'
    const base = {
      documentId: id,
      runId: 'run-1',
      startedAt: Date.now(),
      phase: 'running',
    }
    values.set(key, JSON.stringify({ ...base, unexpected: true }))
    expect(binding.agentRunRecovery(id)).toBeUndefined()
    values.set(key, JSON.stringify({ ...base, toolName: '界'.repeat(1500) }))
    expect(binding.agentRunRecovery(id)).toBeUndefined()
    values.set(key, JSON.stringify({ ...base, instruction: 'Old embedded request' }))
    expect(binding.agentRunRecovery(id)).toMatchObject({ instruction: '', phase: 'running' })
    await binding.updateAgentRun(id, 'run-1', 'tool_pending', 'read_page')
    expect(values.get(key)).not.toContain('Old embedded request')
  })
})
