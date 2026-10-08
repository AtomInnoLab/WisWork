import type { AgentStreamCallbacks, ToolExecution } from '@wiswork/agent-core'
import { describe, expect, it, vi } from 'vitest'
import { createOfficeAgentSession } from '../src/agent/use-office-agent.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import {
  createPresentationAgentRunCheckpoint,
  createPresentationDocumentBinding,
} from '../src/skills/powerpoint/presentation-document.js'

async function fixture(mutated = false) {
  const settings = new Map<string, string>()
  const local = new Map<string, string>()
  let localWriteUnavailable = false
  let location = 'file:///recovery.pptx'
  const binding = createPresentationDocumentBinding(
    {
      get: (key) => settings.get(key),
      set: (key, value) => void settings.set(key, value),
      save: async () => undefined,
      location: () => location,
    },
    () => 'recovery-document',
  )
  const documentId = await binding.documentId()
  const storage = {
    getItem: (key: string) => local.get(key) ?? null,
    setItem: (key: string, value: string) => {
      if (localWriteUnavailable) throw new Error('local storage unavailable')
      local.set(key, value)
    },
    removeItem: (key: string) => void local.delete(key),
  }
  const checkpoint = createPresentationAgentRunCheckpoint(binding, documentId, storage)
  let callbacks!: AgentStreamCallbacks
  const stream = vi.fn((_request: unknown, next: AgentStreamCallbacks) => {
    callbacks = next
    return { cancel: vi.fn() }
  })
  const finish = vi.fn(checkpoint.finish)
  const executeTool = vi.fn(async (): Promise<ToolExecution> => ({
    output: '{}',
    summary: 'Recorded',
    ...(mutated ? { mutated: true } : {}),
  }))
  const session = (validateDocument = async () => (await binding.documentId()) === documentId) =>
    createOfficeAgentSession({
      transport: { stream },
      skill: { id: 'recovery', systemPrompt: 'test', tools: [], executeTool },
      proposals: createStructuredProposalController(),
      runCheckpoint: {
        interrupted: Boolean(checkpoint.recovery()),
        recovery: checkpoint.recovery(),
        readRecovery: checkpoint.recovery,
        validateDocument,
        begin: checkpoint.begin,
        tool: checkpoint.tool,
        finish,
      },
    })
  return {
    session,
    checkpoint,
    settings,
    stream,
    executeTool,
    finish,
    callbacks: () => callbacks,
    setLocalWriteUnavailable: (value: boolean) => {
      localWriteUnavailable = value
    },
    changeDocument: () => {
      location = 'file:///another.pptx'
    },
  }
}

describe('presentation transient failure with a real run checkpoint', () => {
  it('retains a completed read and resumes explicitly after reopening the panel', async () => {
    const f = await fixture()
    const first = f.session()
    first.send('Private request for this deck')
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(1))
    f.callbacks().onToolCall({ id: 'read-1', name: 'read_presentation_plan', input: {} })
    f.callbacks().onDone()
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(2))
    f.callbacks().onError('network_error')
    expect(f.finish).not.toHaveBeenCalled()
    expect(f.checkpoint.recovery()).toMatchObject({
      instruction: 'Private request for this deck',
      phase: 'tool_completed',
      restartSafe: true,
      toolCallId: 'read-1',
    })
    expect(JSON.stringify([...f.settings.values()])).not.toContain('Private request for this deck')
    first.dispose()

    const reopened = f.session()
    expect(reopened.snapshot().recoveryAvailable).toBe(true)
    expect(f.stream).toHaveBeenCalledTimes(2)
    await reopened.resumeInterrupted?.()
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(3))
    expect(f.executeTool).toHaveBeenCalledOnce()
    reopened.dispose()
  })

  it('retains a completed write but never replays the request through retry or reopening', async () => {
    const f = await fixture(true)
    const first = f.session()
    first.send('Modify this deck')
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(1))
    f.callbacks().onToolCall({
      id: 'write-1',
      name: 'replace_existing_presentation_image',
      input: {},
    })
    f.callbacks().onDone()
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(2))
    f.callbacks().onError('provider_unavailable')
    expect(f.finish).not.toHaveBeenCalled()
    expect(f.checkpoint.recovery()).toMatchObject({ phase: 'tool_completed', restartSafe: false })
    expect(first.snapshot()).toMatchObject({ retryable: false, recoveryAvailable: false })
    first.retry()
    await first.resumeInterrupted?.()
    expect(f.stream).toHaveBeenCalledTimes(2)
    first.dispose()

    const reopened = f.session()
    expect(reopened.snapshot().recoveryAvailable).toBe(false)
    await reopened.resumeInterrupted?.()
    expect(f.stream).toHaveBeenCalledTimes(2)
    expect(f.executeTool).toHaveBeenCalledOnce()
    reopened.dispose()
  })

  it('does not resume a preserved pre-tool request in a Save As document', async () => {
    const f = await fixture()
    const first = f.session()
    first.send('Only for the original deck')
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(1))
    f.callbacks().onError('request_timeout')
    expect(f.finish).not.toHaveBeenCalled()
    expect(f.checkpoint.recovery()?.phase).toBe('running')
    first.dispose()

    const reopened = f.session()
    f.changeDocument()
    await reopened.resumeInterrupted?.()
    reopened.retry()
    expect(f.stream).toHaveBeenCalledTimes(1)
    reopened.dispose()
  })

  it('keeps a longer recoverable brief intact after a transient failure', async () => {
    const f = await fixture()
    const session = f.session()
    const brief = '资料、出处与每页制作要求。'.repeat(120)
    session.send(brief)
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(1))
    f.callbacks().onError('provider_unavailable')
    expect(f.checkpoint.recovery()?.instruction).toBe(brief)
    expect(session.snapshot().retryable).toBe(true)
    session.retry()
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(2))
    expect(f.checkpoint.recovery()?.instruction).toBe(brief)
    session.dispose()
  })

  it('rejects retry if a previously safe live checkpoint becomes unreadable', async () => {
    const f = await fixture()
    const session = f.session()
    session.send('Preserve this request')
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(1))
    f.callbacks().onError('network_error')
    expect(session.snapshot().retryable).toBe(true)
    f.settings.set('wiswork.presentation.agent-run.v1', '{damaged')
    session.retry()
    expect(f.stream).toHaveBeenCalledTimes(1)
    expect(f.finish).not.toHaveBeenCalled()
    session.dispose()
  })

  it('retries the initial checkpoint save after local storage becomes available', async () => {
    const f = await fixture()
    f.setLocalWriteUnavailable(true)
    const session = f.session()
    session.send('Create a new deck')
    await vi.waitFor(() =>
      expect(session.snapshot()).toMatchObject({
        error: 'presentation_run_checkpoint_unavailable',
        retryable: true,
        busy: false,
      }),
    )
    expect(f.stream).not.toHaveBeenCalled()
    expect(f.checkpoint.recovery()).toBeUndefined()
    f.setLocalWriteUnavailable(false)
    session.retry()
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledOnce())
    expect(f.checkpoint.recovery()?.instruction).toBe('Create a new deck')
    session.dispose()
  })

  it('does not replace another safe run created during document validation', async () => {
    const f = await fixture()
    await f.checkpoint.begin('run-A', 'Request A')
    let release!: (value: boolean) => void
    const validate = vi.fn(
      () =>
        new Promise<boolean>((resolve) => {
          release = resolve
        }),
    )
    const session = f.session(validate)
    const pending = session.resumeInterrupted?.()
    expect(validate).toHaveBeenCalledOnce()
    await f.checkpoint.begin('run-B', 'Request B')
    release(true)
    await pending
    expect(f.stream).not.toHaveBeenCalled()
    expect(f.checkpoint.recovery()).toMatchObject({ runId: 'run-B', instruction: 'Request B' })
    session.dispose()
  })

  it('does not overwrite a new run while retrying an initial checkpoint save', async () => {
    const f = await fixture()
    f.setLocalWriteUnavailable(true)
    let release!: (value: boolean) => void
    const validate = vi.fn(
      () =>
        new Promise<boolean>((resolve) => {
          release = resolve
        }),
    )
    const session = f.session(validate)
    session.send('Initial request')
    await vi.waitFor(() =>
      expect(session.snapshot().error).toBe('presentation_run_checkpoint_unavailable'),
    )
    f.setLocalWriteUnavailable(false)
    session.retry()
    expect(validate).toHaveBeenCalledOnce()
    await f.checkpoint.begin('new-run', 'Another panel request')
    release(true)
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(f.stream).not.toHaveBeenCalled()
    expect(f.checkpoint.recovery()).toMatchObject({
      runId: 'new-run',
      instruction: 'Another panel request',
    })
    session.dispose()
  })

  it('can retry again when saving the restarted safe run fails once', async () => {
    const f = await fixture()
    const session = f.session()
    session.send('Read this deck')
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledOnce())
    f.callbacks().onError('network_error')
    f.setLocalWriteUnavailable(true)
    session.retry()
    await vi.waitFor(() =>
      expect(session.snapshot().error).toBe('presentation_run_checkpoint_unavailable'),
    )
    expect(f.stream).toHaveBeenCalledOnce()
    expect(f.checkpoint.recovery()).toBeUndefined()
    f.setLocalWriteUnavailable(false)
    session.retry()
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(2))
    expect(f.checkpoint.recovery()?.instruction).toBe('Read this deck')
    session.dispose()
  })
})
