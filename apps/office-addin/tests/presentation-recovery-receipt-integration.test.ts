import { describe, expect, it, vi } from 'vitest'
import type { AgentStreamCallbacks } from '@wiswork/agent-core'
import { createOfficeAgentSession } from '../src/agent/use-office-agent.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import {
  createPresentationAgentRunCheckpoint,
  createPresentationDocumentBinding,
} from '../src/skills/powerpoint/presentation-document.js'
import type { PresentationTextChange } from '../src/skills/powerpoint/presentation-text-change.js'

describe('write receipt feedback through a real presentation checkpoint', () => {
  it('uses the newly persisted import receipt on failure and reopen without replaying it', async () => {
    const settings = new Map<string, string>()
    const binding = createPresentationDocumentBinding(
      {
        get: (key) => settings.get(key),
        set: (key, value) => void settings.set(key, value),
        save: async () => undefined,
        location: () => 'file:///deck.pptx',
      },
      () => 'document',
    )
    const documentId = await binding.documentId()
    const checkpoint = createPresentationAgentRunCheckpoint(binding, documentId)
    let callbacks!: AgentStreamCallbacks
    const stream = vi.fn((_request, next: AgentStreamCallbacks) => {
      callbacks = next
      return { cancel: vi.fn() }
    })
    const executeTool = vi.fn(async (call) => {
      await binding.writeReceipt('project/request', {
        state: 'complete',
        documentId,
        toolCallId: call.id,
        slideIds: ['slide-1', 'slide-2'],
      })
      return { output: '{}', summary: 'Imported', mutated: true }
    })
    const open = () =>
      createOfficeAgentSession({
        transport: { stream },
        skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool },
        proposals: createStructuredProposalController(),
        runCheckpoint: {
          ...checkpoint,
          interrupted: Boolean(checkpoint.recovery()),
          recovery: checkpoint.recovery(),
          readRecovery: checkpoint.recovery,
          validateDocument: async () => true,
        },
      })
    const session = open()
    session.send('Import this presentation')
    await vi.waitFor(() => expect(stream).toHaveBeenCalledTimes(1))
    callbacks.onToolCall({ id: 'import', name: 'import_generated_presentation', input: {} })
    callbacks.onDone()
    await vi.waitFor(() => expect(stream).toHaveBeenCalledTimes(2))
    callbacks.onError('network_error')
    expect(session.snapshot().errorMessage).toContain('2 页已记录导入')
    expect(session.snapshot().errorMessage).toContain('视觉审查与交付仍需核验')
    expect(session.snapshot().errorMessage).toContain('本次前台运行')
    expect(session.snapshot().retryable).toBe(false)
    session.retry()
    expect(executeTool).toHaveBeenCalledTimes(1)
    expect(stream).toHaveBeenCalledTimes(2)
    session.dispose()
    const reopened = open()
    const notice = reopened.snapshot().timeline[0]
    const text = notice?.kind === 'system' ? notice.text : ''
    expect(text).toContain('2 页已记录导入')
    expect(text).not.toMatch(/import_generated_presentation|tool_completed/)
    expect(reopened.snapshot().recoveryAvailable).toBe(false)
    reopened.dispose()
  })
  it('describes settled and unresolved modifications from the durable history', async () => {
    const settings = new Map<string, string>()
    const binding = createPresentationDocumentBinding(
      {
        get: (key) => settings.get(key),
        set: (key, value) => void settings.set(key, value),
        save: async () => undefined,
        location: () => 'file:///edits.pptx',
      },
      () => 'document',
    )
    const documentId = await binding.documentId()
    const checkpoint = createPresentationAgentRunCheckpoint(binding, documentId)
    let callbacks!: AgentStreamCallbacks
    const stream = vi.fn((_request, next: AgentStreamCallbacks) => {
      callbacks = next
      return { cancel: vi.fn() }
    })
    const executeTool = vi.fn(async () => {
      const pending: PresentationTextChange = {
        version: 1,
        documentId,
        changeId: 'one',
        projectId: 'project',
        requestId: 'request',
        artifactDigest: 'a'.repeat(64),
        pageId: 'page',
        hostSlideId: 'slide',
        shapeId: 'shape',
        before: 'one',
        after: 'two',
        state: 'pending',
      }
      await binding.writeTextChange(pending, undefined)
      const applied = { ...pending, state: 'applied' as const }
      await binding.writeTextChange(applied, pending)
      const undoing = { ...pending, state: 'undo_pending' as const }
      await binding.writeTextChange(undoing, applied)
      const undone = { ...pending, state: 'undone' as const }
      await binding.writeTextChange(undone, undoing)
      await binding.writeTextChange({ ...pending, changeId: 'two' }, undone)
      return { output: '{}', summary: 'Recorded', mutated: true }
    })
    const session = createOfficeAgentSession({
      transport: { stream },
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool },
      proposals: createStructuredProposalController(),
      runCheckpoint: {
        ...checkpoint,
        interrupted: false,
        recovery: checkpoint.recovery(),
        readRecovery: checkpoint.recovery,
        validateDocument: async () => true,
      },
    })
    session.send('Modify this presentation')
    await vi.waitFor(() => expect(stream).toHaveBeenCalledTimes(1))
    callbacks.onToolCall({ id: 'change', name: 'edit_presentation_page_text', input: {} })
    callbacks.onDone()
    await vi.waitFor(() => expect(stream).toHaveBeenCalledTimes(2))
    callbacks.onError('network_error')
    expect(session.snapshot().errorMessage).toContain('对应修改历史 2 项，其中 1 项未结算')
    expect(session.snapshot().errorMessage).toContain('1 项已结算（可含撤销或丢弃）')
    expect(session.snapshot().errorMessage).toContain('变更历史核对保存点与宿主对象')
    expect(session.snapshot().retryable).toBe(false)
    session.retry()
    expect(executeTool).toHaveBeenCalledTimes(1)
    expect(stream).toHaveBeenCalledTimes(2)
    session.dispose()
  })
})
