import { describe, expect, it, vi } from 'vitest'
import type { AgentMessage, AgentStreamCallbacks } from '@wiswork/agent-core'
import { createOfficeAgentSession } from '../src/agent/use-office-agent.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import {
  createPresentationAgentRunCheckpoint,
  createPresentationDocumentBinding,
} from '../src/skills/powerpoint/presentation-document.js'

async function fixture() {
  const settings = new Map<string, string>()
  const local = new Map<string, string>()
  let failed = false
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
  const storage = {
    getItem: (key: string) => local.get(key) ?? null,
    setItem: (key: string, value: string) => {
      if (failed) throw new Error('full')
      local.set(key, value)
    },
    removeItem: (key: string) => void local.delete(key),
  }
  const checkpoint = () => createPresentationAgentRunCheckpoint(binding, documentId, storage)
  let callbacks!: AgentStreamCallbacks
  const stream = vi.fn((_request, next: AgentStreamCallbacks) => {
    callbacks = next
    return { cancel: vi.fn() }
  })
  const executeTool = vi.fn(async () => ({
    output: 'Private durable evidence',
    summary: 'read',
    mutated: false,
  }))
  const session = (
    options: {
      failConversation?: boolean
      validateDocument?: () => Promise<boolean>
      conversation?: (runId: string, messages: readonly AgentMessage[]) => Promise<void>
    } = {},
  ) => {
    const record = checkpoint()
    return createOfficeAgentSession({
      transport: { stream },
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool },
      proposals: createStructuredProposalController(),
      runCheckpoint: {
        ...record,
        ...(options.failConversation
          ? {
              conversation: async () => {
                throw new Error('presentation_run_checkpoint_unavailable')
              },
            }
          : {}),
        ...(options.conversation ? { conversation: options.conversation } : {}),
        interrupted: Boolean(record.recovery()),
        recovery: record.recovery(),
        readRecovery: record.recovery,
        validateDocument: options.validateDocument ?? (async () => true),
      },
    })
  }
  return {
    checkpoint,
    session,
    stream,
    executeTool,
    callbacks: () => callbacks,
    settings,
    local,
    failStorage: () => {
      failed = true
    },
  }
}
const messages: AgentMessage[] = [
  { role: 'user', text: 'Read the deck' },
  {
    role: 'assistant',
    text: '',
    toolCalls: [{ id: 'read-1', name: 'read_presentation_plan', input: {} }],
  },
  {
    role: 'tool',
    results: [{ id: 'read-1', name: 'read_presentation_plan', output: 'Private durable evidence' }],
  },
]
describe('presentation completed-read conversation checkpoints', () => {
  it('restores an audited baseline read without replaying its completed tool', async () => {
    const f = await fixture()
    const checkpoint = f.checkpoint()
    const toolName = 'check_presentation_baseline'
    await checkpoint.begin('run', 'Read the deck')
    await checkpoint.tool('run', 'tool_pending', toolName, false, 'baseline-1')
    await checkpoint.tool('run', 'tool_completed', toolName, false, 'baseline-1')
    const saved: AgentMessage[] = [
      { role: 'user', text: 'Read the deck' },
      { role: 'assistant', text: '', toolCalls: [{ id: 'baseline-1', name: toolName, input: {} }] },
      { role: 'tool', results: [{ id: 'baseline-1', name: toolName, output: 'unchanged' }] },
    ]
    await checkpoint.conversation('run', saved)
    const reopened = f.session()
    expect(reopened.snapshot().recoveryAvailable).toBe(true)
    await reopened.resumeInterrupted!()
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledOnce())
    expect(f.stream.mock.calls[0]![0].messages).toEqual(saved)
    expect(f.executeTool).not.toHaveBeenCalled()
    reopened.dispose()
  })

  it('retries a complete multi-read batch once and removes local results after completion', async () => {
    const f = await fixture()
    const session = f.session()
    session.send('Read the deck')
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(1))
    f.callbacks().onToolCall({ id: 'read-1', name: 'read_presentation_plan', input: {} })
    f.callbacks().onToolCall({ id: 'read-2', name: 'read_presentation_production', input: {} })
    f.callbacks().onDone()
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(2))
    f.callbacks().onError('network_error')
    const runId = f.checkpoint().recovery()!.runId
    session.retry()
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(3))
    expect(f.checkpoint().recovery()?.runId).toBe(runId)
    const restored = f.stream.mock.calls[2]![0].messages as AgentMessage[]
    expect(restored.filter((message) => message.role === 'user')).toHaveLength(1)
    const last = restored.at(-1)
    expect(last?.role === 'tool' ? last.results.map((result) => result.id) : []).toEqual([
      'read-1',
      'read-2',
    ])
    expect(f.stream.mock.calls[2]![0].system).toContain('historical observations')
    expect(f.executeTool).toHaveBeenCalledTimes(2)
    f.callbacks().onDone()
    await vi.waitFor(() => expect(f.local.size).toBe(0))
    expect(f.checkpoint().recovery()).toBeUndefined()
    session.dispose()
  })
  it('rejects forged, oversized, unknown-tool and expired conversation records', async () => {
    const f = await fixture()
    const checkpoint = f.checkpoint()
    await checkpoint.begin('run', 'Read the deck')
    await checkpoint.tool('run', 'tool_pending', 'read_presentation_plan', false, 'read-1')
    await checkpoint.tool('run', 'tool_completed', 'read_presentation_plan', false, 'read-1')
    for (const invalid of [
      [{ role: 'user', text: 'Another instruction' }, ...messages.slice(1)],
      [
        messages[0],
        { role: 'assistant', text: '', toolCalls: [{ id: 'read-1', name: 'unknown', input: {} }] },
        { role: 'tool', results: [{ id: 'read-1', name: 'unknown', output: 'data' }] },
      ],
      [
        ...messages.slice(0, 2),
        {
          role: 'tool',
          results: [
            { id: 'read-1', name: 'read_presentation_plan', output: 'x'.repeat(128 * 1024) },
          ],
        },
      ],
    ]) {
      await checkpoint.conversation('run', invalid as AgentMessage[])
      expect(f.checkpoint().recovery()?.messages).toBeUndefined()
    }
    await checkpoint.conversation('run', messages)
    const [key, raw] = [...f.local.entries()][0]!
    f.local.set(key, JSON.stringify({ ...JSON.parse(raw), expiresAt: 0 }))
    expect(f.checkpoint().recovery()?.instruction).toBe('')
    expect(f.checkpoint().recovery()?.messages).toBeUndefined()
  })
  it('does not save conversations for writes and surfaces local storage failure', async () => {
    const f = await fixture()
    const checkpoint = f.checkpoint()
    await checkpoint.begin('run', 'Read the deck')
    await checkpoint.tool('run', 'tool_pending', 'read_presentation_plan', false, 'read-1')
    await checkpoint.tool('run', 'tool_completed', 'read_presentation_plan', false, 'read-1')
    f.failStorage()
    await expect(checkpoint.conversation('run', messages)).rejects.toThrow(
      'presentation_run_checkpoint_unavailable',
    )
    const other = await fixture()
    const write = other.checkpoint()
    await write.begin('write', 'Read the deck')
    await write.tool('write', 'tool_pending', 'set_text', false, 'read-1')
    await write.tool('write', 'tool_completed', 'set_text', true, 'read-1')
    await write.conversation('write', messages)
    expect(other.checkpoint().recovery()?.messages).toBeUndefined()
    expect(JSON.stringify([...other.local.values()])).not.toContain('Private durable evidence')
  })
  it('keeps results local, binds them to the completed call and rejects a changed call', async () => {
    const f = await fixture()
    const checkpoint = f.checkpoint()
    await checkpoint.begin('run', 'Read the deck')
    await checkpoint.tool('run', 'tool_pending', 'read_presentation_plan', false, 'read-1')
    await checkpoint.tool('run', 'tool_completed', 'read_presentation_plan', false, 'read-1')
    await checkpoint.conversation('run', messages)
    expect(f.checkpoint().recovery()?.messages).toEqual(messages)
    expect(JSON.stringify([...f.settings.values()])).not.toContain('Private durable evidence')
    const withPreviousCompletedTurn: AgentMessage[] = [
      { role: 'user', text: 'Earlier edit' },
      {
        role: 'assistant',
        text: '',
        toolCalls: [{ id: 'old-write', name: 'set_text', input: {} }],
      },
      {
        role: 'tool',
        results: [{ id: 'old-write', name: 'set_text', output: 'Earlier edit completed' }],
      },
      ...messages,
    ]
    await checkpoint.conversation('run', withPreviousCompletedTurn)
    expect(f.checkpoint().recovery()?.messages).toEqual(withPreviousCompletedTurn)
    await checkpoint.tool('run', 'tool_pending', 'read_presentation_plan', false, 'read-1')
    expect(f.checkpoint().recovery()?.messages).toBeUndefined()
    await checkpoint.tool('run', 'tool_completed', 'read_presentation_plan', false, 'read-1')
    expect(f.checkpoint().recovery()?.messages).toBeUndefined()
  })
  it('resumes the original result batch after closing the panel without rerunning the read', async () => {
    const f = await fixture()
    const first = f.session()
    first.send('Read the deck')
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(1))
    f.callbacks().onToolCall({ id: 'read-1', name: 'read_presentation_plan', input: {} })
    f.callbacks().onDone()
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(2))
    f.callbacks().onError('network_error')
    first.dispose()
    const reopened = f.session()
    await reopened.resumeInterrupted!()
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(3))
    const restored = f.stream.mock.calls[2]![0].messages as AgentMessage[]
    expect(restored.filter((message) => message.role === 'user')).toHaveLength(1)
    expect(restored.at(-1)).toEqual(messages.at(-1))
    expect(f.executeTool).toHaveBeenCalledTimes(1)
    reopened.dispose()
  })
  it('blocks the next model request when saving the completed batch fails', async () => {
    const f = await fixture()
    const session = f.session({ failConversation: true })
    session.send('Read the deck')
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(1))
    f.callbacks().onToolCall({ id: 'read-1', name: 'read_presentation_plan', input: {} })
    f.callbacks().onDone()
    await vi.waitFor(() =>
      expect(session.snapshot().error).toBe('presentation_run_checkpoint_unavailable'),
    )
    expect(f.stream).toHaveBeenCalledTimes(1)
    expect(f.checkpoint().recovery()?.phase).toBe('tool_completed')
    expect(session.snapshot().errorMessage).toContain('读取结果未能保存')
    expect(session.snapshot().errorMessage).not.toContain('未自动重放写入')
    session.dispose()
  })
  it('rejects a changed result snapshot while document validation is waiting', async () => {
    const f = await fixture()
    const saved = f.checkpoint()
    await saved.begin('run', 'Read the deck')
    await saved.tool('run', 'tool_pending', 'read_presentation_plan', false, 'read-1')
    await saved.tool('run', 'tool_completed', 'read_presentation_plan', false, 'read-1')
    await saved.conversation('run', messages)
    let release!: () => void
    const reopened = f.session({
      validateDocument: async () => {
        await new Promise<void>((resolve) => {
          release = resolve
        })
        return true
      },
    })
    const resuming = reopened.resumeInterrupted!()
    await vi.waitFor(() => expect(release).toBeTypeOf('function'))
    const [key, raw] = [...f.local.entries()][0]!
    const record = JSON.parse(raw)
    record.messages[2].results[0].output = 'Different evidence'
    f.local.set(key, JSON.stringify(record))
    release()
    await resuming
    expect(f.stream).not.toHaveBeenCalled()
    expect(f.checkpoint().recovery()?.runId).toBe('run')
    reopened.dispose()
  })
  it('stops while a batch is being saved without issuing the next request', async () => {
    const f = await fixture()
    let release!: () => void
    const session = f.session({
      conversation: async () => {
        await new Promise<void>((resolve) => {
          release = resolve
        })
      },
    })
    session.send('Read the deck')
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(1))
    f.callbacks().onToolCall({ id: 'read-1', name: 'read_presentation_plan', input: {} })
    f.callbacks().onDone()
    await vi.waitFor(() => expect(release).toBeTypeOf('function'))
    session.stop()
    release()
    await vi.waitFor(() => expect(session.snapshot().busy).toBe(false))
    expect(f.stream).toHaveBeenCalledTimes(1)
    expect(session.snapshot().status).toBe('cancelled')
    session.dispose()
  })
})
