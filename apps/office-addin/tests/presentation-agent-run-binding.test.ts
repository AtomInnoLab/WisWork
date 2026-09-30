import { describe, expect, it, vi } from 'vitest'
import {
  createPresentationAgentRunCheckpoint,
  createPresentationDocumentBinding,
  preparePresentationAgentRunRecovery,
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
    await checkpoint.tool('run-1', 'tool_pending', 'read_presentation_plan')
    expect(createPresentationAgentRunCheckpoint(binding, id, storage).recovery()).toMatchObject({
      instruction: 'Private request',
      phase: 'tool_pending',
      restartSafe: true,
    })
    await checkpoint.tool('run-1', 'tool_completed', 'read_presentation_plan')
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

  it('recovers a longer local brief without embedding it in document settings', async () => {
    const values = new Map<string, string>()
    const local = new Map<string, string>()
    const storage = {
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
    const brief = '研究资料与页面要求。'.repeat(200)
    expect(new TextEncoder().encode(brief).byteLength).toBeGreaterThan(3000)
    await checkpoint.begin('long-run', brief)
    expect(values.get('wiswork.presentation.agent-run.v1')).not.toContain(brief)
    expect(createPresentationAgentRunCheckpoint(binding, id, storage).recovery()?.instruction).toBe(
      brief,
    )
    await checkpoint.finish('long-run')
    expect(local.size).toBe(0)
    await checkpoint.begin('oversized-run', 'A'.repeat(9000))
    expect(createPresentationAgentRunCheckpoint(binding, id, storage).recovery()?.instruction).toBe(
      '',
    )
    await checkpoint.finish('oversized-run')
  })

  it('binds an unfinished tool and its completion to the same model call', async () => {
    const values = new Map<string, string>()
    const local = new Map<string, string>()
    const storage = {
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
    await checkpoint.begin('run-1', 'Read plan')
    await checkpoint.tool('run-1', 'tool_pending', 'read_presentation_plan', false, 'call-1')
    expect(checkpoint.recovery()).toMatchObject({ toolCallId: 'call-1', restartSafe: true })
    await expect(
      checkpoint.tool('run-1', 'tool_completed', 'read_presentation_plan', false, 'call-2'),
    ).rejects.toThrow('presentation_run_checkpoint_unavailable')
    expect(checkpoint.recovery()).toMatchObject({ phase: 'tool_pending', toolCallId: 'call-1' })
    await checkpoint.tool('run-1', 'tool_completed', 'read_presentation_plan', false, 'call-1')
    expect(createPresentationAgentRunCheckpoint(binding, id, storage).recovery()).toMatchObject({
      phase: 'tool_completed',
      toolCallId: 'call-1',
      restartSafe: true,
    })
    await checkpoint.tool('run-1', 'tool_pending', 'list_presentation_attachments', false, 'call-2')
    expect(checkpoint.recovery()).toMatchObject({ phase: 'tool_pending', toolCallId: 'call-2' })
    await checkpoint.tool(
      'run-1',
      'tool_completed',
      'list_presentation_attachments',
      false,
      'call-2',
    )
    const key = 'wiswork.presentation.agent-run.v1'
    values.set(key, JSON.stringify({ ...JSON.parse(values.get(key)!), toolCallId: 'forged' }))
    expect(createPresentationAgentRunCheckpoint(binding, id, storage).recovery()?.instruction).toBe(
      '',
    )
  })

  it('requires the local checkpoint to agree before offering a document-provided safe restart', async () => {
    const values = new Map<string, string>()
    const local = new Map<string, string>()
    const storage = {
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
    await checkpoint.begin('run-1', 'Change this slide')
    const [localKey] = local.keys()
    const initial = local.get(localKey)!
    for (const changed of [
      { ...JSON.parse(initial), phase: 'tool_completed', toolName: 'read_presentation_plan' },
      { ...JSON.parse(initial), restartSafe: false },
    ]) {
      local.set(localKey, JSON.stringify(changed))
      expect(
        createPresentationAgentRunCheckpoint(binding, id, storage).recovery()?.instruction,
      ).toBe('')
    }
    local.set(localKey, 'invalid json')
    expect(createPresentationAgentRunCheckpoint(binding, id, storage).recovery()?.instruction).toBe(
      '',
    )
    local.set(localKey, initial)
    await checkpoint.tool('run-1', 'tool_pending', 'write_presentation_page')
    await checkpoint.tool('run-1', 'tool_completed', 'write_presentation_page', true)
    const key = 'wiswork.presentation.agent-run.v1'
    const actual = values.get(key)!
    values.set(
      key,
      JSON.stringify({
        ...JSON.parse(actual),
        phase: 'tool_completed',
        toolName: 'read_presentation_plan',
        restartSafe: true,
      }),
    )
    expect(binding.agentRunRecovery(id)?.restartSafe).toBe(true)
    expect(createPresentationAgentRunCheckpoint(binding, id, storage).recovery()?.restartSafe).toBe(
      false,
    )
    values.set(key, actual)
    expect(createPresentationAgentRunCheckpoint(binding, id, storage).recovery()?.restartSafe).toBe(
      false,
    )
  })

  it.each(['throws', 'silently drops'] as const)(
    'fails closed when the browser %s the next local tool boundary',
    async (failure) => {
      const values = new Map<string, string>()
      const local = new Map<string, string>()
      let rejectWrites = false
      const storage = {
        getItem: (key: string) => local.get(key) ?? null,
        setItem: (key: string, value: string) => {
          if (rejectWrites) {
            if (failure === 'throws') throw new Error('storage unavailable')
            return
          }
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
      await checkpoint.begin('run-1', 'Read plan')
      rejectWrites = true
      await expect(
        checkpoint.tool('run-1', 'tool_pending', 'read_presentation_plan'),
      ).rejects.toThrow()
      expect(binding.agentRunRecovery(id)?.phase).toBe('tool_pending')
      expect(
        createPresentationAgentRunCheckpoint(binding, id, storage).recovery()?.restartSafe,
      ).toBe(false)
    },
  )

  it('correlates an interrupted import call with its durable receipt', async () => {
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
    const checkpoint = createPresentationAgentRunCheckpoint(binding, id)
    await checkpoint.begin('run-1')
    await checkpoint.tool('run-1', 'tool_pending', 'import_generated_presentation', false, 'call-1')
    expect(checkpoint.recovery()?.importReceipt).toBeUndefined()
    await binding.writeReceipt('project/request', {
      state: 'pending',
      documentId: id,
      toolCallId: 'call-1',
    })
    expect(createPresentationAgentRunCheckpoint(binding, id).recovery()?.importReceipt).toEqual({
      state: 'uncertain',
      completed: 0,
    })
    expect(binding.readReceipt('project/request')?.agentRunId).toBe('run-1')
    expect(binding.agentImportReceipt('another-document', 'run-1', 'call-1')).toBeUndefined()
    expect(binding.agentImportReceipt(id, 'other-run', 'call-1')).toBeUndefined()
    expect(binding.agentImportReceipt(id, 'run-1', 'other-call')).toBeUndefined()
    expect(binding.agentImportReceipt(id, 'run-1', 'call-1')).toEqual({
      state: 'uncertain',
      completed: 0,
    })
    await binding.writeReceipt('project/request', {
      state: 'complete',
      documentId: id,
      toolCallId: 'call-1',
      slideIds: ['slide-1'],
    })
    expect(checkpoint.recovery()?.importReceipt).toEqual({ state: 'complete', completed: 1 })
    expect(binding.readReceipt('project/request')?.agentRunId).toBe('run-1')
    await checkpoint.finish('run-1')
    expect(checkpoint.recovery()).toBeUndefined()
    await checkpoint.begin('run-2')
    await checkpoint.tool('run-2', 'tool_pending', 'import_generated_presentation', false, 'call-1')
    expect(checkpoint.recovery()?.importReceipt).toBeUndefined()
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
    expect(await preparePresentationAgentRunRecovery(binding, id)).toEqual({
      interrupted: true,
      scrubFailed: true,
    })
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

  it('allows restarting only when every checkpointed tool is explicitly read-only', async () => {
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
    await binding.rememberAgentRun(id, 'run-1')
    await expect(
      binding.updateAgentRun(id, 'run-1', 'tool_completed', 'read_presentation_plan'),
    ).rejects.toThrow('presentation_run_checkpoint_unavailable')
    await binding.updateAgentRun(id, 'run-1', 'tool_pending', 'read_presentation_plan')
    expect(binding.agentRunRecovery(id)?.restartSafe).toBe(true)
    await binding.updateAgentRun(id, 'run-1', 'tool_completed', 'read_presentation_plan')
    await binding.updateAgentRun(id, 'run-1', 'tool_pending', 'list_presentation_attachments')
    expect(binding.agentRunRecovery(id)?.restartSafe).toBe(true)
    await binding.updateAgentRun(id, 'run-1', 'tool_completed', 'list_presentation_attachments')
    await binding.updateAgentRun(id, 'run-1', 'tool_pending', 'write_presentation_page')
    expect(binding.agentRunRecovery(id)?.restartSafe).toBe(false)
    const key = 'wiswork.presentation.agent-run.v1'
    const forged = JSON.parse(values.get(key)!) as Record<string, unknown>
    values.set(key, JSON.stringify({ ...forged, restartSafe: true }))
    expect(binding.agentRunRecovery(id)?.restartSafe).toBe(false)
    values.set(key, JSON.stringify(forged))
    await binding.updateAgentRun(id, 'run-1', 'tool_completed', 'write_presentation_page')
    await binding.updateAgentRun(id, 'run-1', 'tool_pending', 'read_presentation_plan')
    expect(binding.agentRunRecovery(id)?.restartSafe).toBe(false)
  })

  it.each([
    'read_presentation_page_reviews',
    'read_presentation_claim_review',
    'read_presentation_claim_evidence',
    'check_presentation_page_content',
    'read_presentation_production',
    'list_presentation_review_comments',
    'check_presentation_baseline',
    'check_presentation_baseline_windows',
    'list_slide_shapes',
    'read_slide_text',
    'verify_slides',
    'read_presentation_image_replacement',
    'read_presentation_preference_candidates',
  ])('keeps %s restart-safe only after an unmutated completion', async (toolName) => {
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
    await binding.rememberAgentRun(id, 'read-run')
    await binding.updateAgentRun(id, 'read-run', 'tool_pending', toolName)
    expect(binding.agentRunRecovery(id)?.restartSafe).toBe(true)
    await binding.updateAgentRun(id, 'read-run', 'tool_completed', toolName)
    expect(binding.agentRunRecovery(id)?.restartSafe).toBe(true)
    await binding.finishAgentRun(id, 'read-run')
    await binding.rememberAgentRun(id, 'mutated-run')
    await binding.updateAgentRun(id, 'mutated-run', 'tool_pending', toolName)
    await binding.updateAgentRun(id, 'mutated-run', 'tool_completed', toolName, true)
    expect(binding.agentRunRecovery(id)?.restartSafe).toBe(false)
  })

  it('never marks a persisted source audit safe to replay after restart', async () => {
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
    await binding.rememberAgentRun(id, 'audit-run')
    await binding.updateAgentRun(id, 'audit-run', 'tool_pending', 'audit_presentation_sources')
    expect(binding.agentRunRecovery(id)?.restartSafe).toBe(false)
    await binding.updateAgentRun(id, 'audit-run', 'tool_completed', 'audit_presentation_sources')
    expect(binding.agentRunRecovery(id)?.restartSafe).toBe(false)
  })

  it('never promotes a legacy tool checkpoint or a mutating read result to restart-safe', async () => {
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
    await binding.rememberAgentRun(id, 'run-1')
    const key = 'wiswork.presentation.agent-run.v1'
    const legacy = JSON.parse(values.get(key)!) as Record<string, unknown>
    delete legacy.restartSafe
    values.set(
      key,
      JSON.stringify({ ...legacy, phase: 'tool_pending', toolName: 'read_presentation_plan' }),
    )
    expect(binding.agentRunRecovery(id)?.restartSafe).not.toBe(true)
    await binding.updateAgentRun(id, 'run-1', 'tool_completed', 'read_presentation_plan')
    expect(binding.agentRunRecovery(id)?.restartSafe).toBe(false)
    await binding.finishAgentRun(id, 'run-1')
    await binding.rememberAgentRun(id, 'run-2')
    await binding.updateAgentRun(id, 'run-2', 'tool_pending', 'read_presentation_plan')
    await binding.updateAgentRun(id, 'run-2', 'tool_completed', 'read_presentation_plan', true)
    expect(binding.agentRunRecovery(id)?.restartSafe).toBe(false)
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
