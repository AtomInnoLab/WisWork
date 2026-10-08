import { describe, expect, it, vi } from 'vitest'
import { AgentLoop, type AgentMessage, type AgentStreamCallbacks } from '../src/index.js'

const history: AgentMessage[] = [
  { role: 'user', text: 'Investigate this deck' },
  { role: 'assistant', text: '', toolCalls: [{ id: 'read-1', name: 'read', input: {} }] },
  { role: 'tool', results: [{ id: 'read-1', name: 'read', output: 'saved evidence' }] },
]
function fixture(onTurnEnd?: () => Promise<void>) {
  let callbacks!: AgentStreamCallbacks
  const stream = vi.fn((_request, next: AgentStreamCallbacks) => {
    callbacks = next
    return { cancel: vi.fn() }
  })
  const executeTool = vi.fn(async () => ({ output: 'new evidence', summary: 'read' }))
  const onDone = vi.fn()
  const onError = vi.fn()
  const loop = new AgentLoop({
    transport: { stream },
    skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool },
    events: { onTurnEnd, onDone, onError },
    maxTurns: 1,
  })
  return { loop, stream, executeTool, onDone, onError, callbacks: () => callbacks }
}
describe('completed tool conversation resume', () => {
  it('keeps a previous completed batch when a stopped turn has no closing assistant message', () => {
    const f = fixture()
    const later: AgentMessage[] = [
      ...history,
      { role: 'user', text: 'Follow up after stopping' },
      { role: 'assistant', text: '', toolCalls: [{ id: 'read-2', name: 'read', input: {} }] },
      { role: 'tool', results: [{ id: 'read-2', name: 'read', output: 'latest evidence' }] },
    ]
    expect(f.loop.resume(later)).toBe(true)
    expect(f.stream.mock.calls[0]![0].messages.slice(0, 6)).toEqual(later)
    expect(f.executeTool).not.toHaveBeenCalled()
  })
  it('continues from saved results without duplicating the instruction or executing old tools', () => {
    const f = fixture()
    expect(f.loop.resume(history)).toBe(true)
    expect(f.stream.mock.calls[0]![0].messages.slice(0, 3)).toEqual(history)
    expect(f.stream.mock.calls[0]![0].tools).toEqual([])
    expect(f.executeTool).not.toHaveBeenCalled()
    expect(f.loop.resume(history)).toBe(false)
    f.callbacks().onDelta('Conclusion from evidence')
    f.callbacks().onDone()
    expect(f.onDone).toHaveBeenCalledWith({
      text: 'Conclusion from evidence',
      cancelled: false,
      turnLimit: true,
    })
    expect(history).toHaveLength(3)
  })
  it.each(
    [
      history.slice(0, 2),
      [{ role: 'tool', results: [] }],
      [
        ...history.slice(0, 2),
        { role: 'tool', results: [{ id: 'other', name: 'read', output: 'bad' }] },
      ],
      [
        ...history.slice(0, 2),
        { role: 'tool', results: [{ id: 'read-1', name: 'write', output: 'bad' }] },
      ],
      [...history, { role: 'user', text: 'unanswered' }],
      [
        history[0],
        {
          role: 'assistant',
          text: '',
          toolCalls: [
            {
              id: 'read-1',
              name: 'read',
              input: Array.from({ length: 80 }).reduce<Record<string, unknown>>(
                (input) => ({ nested: input }),
                {},
              ),
            },
          ],
        },
        history[2],
      ],
    ].map((messages) => ({ messages })),
  )('rejects incomplete protocol without launching', ({ messages }) => {
    const f = fixture()
    expect(f.loop.resume(messages as AgentMessage[])).toBe(false)
    expect(f.stream).not.toHaveBeenCalled()
    expect(f.loop.busy).toBe(false)
  })
  it.each(['cancel', 'reset'] as const)(
    'waits for the batch checkpoint and respects %s',
    async (action) => {
      let release!: () => void
      const f = fixture(
        () =>
          new Promise<void>((resolve) => {
            release = resolve
          }),
      )
      f.loop.run('first')
      await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(1))
      f.callbacks().onToolCall({ id: 'new', name: 'read', input: {} })
      f.callbacks().onDone()
      await vi.waitFor(() => expect(release).toBeTypeOf('function'))
      expect(f.stream).toHaveBeenCalledTimes(1)
      f.loop[action]()
      release()
      await vi.waitFor(() => expect(f.loop.busy).toBe(false))
      expect(f.stream).toHaveBeenCalledTimes(1)
    },
  )
  it('does not issue another request after a checkpoint failure', async () => {
    const f = fixture(async () => {
      throw new Error('checkpoint_failed')
    })
    f.loop.run('first')
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(1))
    f.callbacks().onToolCall({ id: 'new', name: 'read', input: {} })
    f.callbacks().onDone()
    await vi.waitFor(() => expect(f.onError).toHaveBeenCalledWith('checkpoint_failed'))
    expect(f.stream).toHaveBeenCalledTimes(1)
    expect(f.loop.busy).toBe(false)
  })
  it('settles cancellation when the in-flight checkpoint rejects', async () => {
    let reject!: (error: Error) => void
    const f = fixture(
      () =>
        new Promise<void>((_resolve, fail) => {
          reject = fail
        }),
    )
    f.loop.run('first')
    await vi.waitFor(() => expect(f.stream).toHaveBeenCalledTimes(1))
    f.callbacks().onToolCall({ id: 'new', name: 'read', input: {} })
    f.callbacks().onDone()
    await vi.waitFor(() => expect(reject).toBeTypeOf('function'))
    f.loop.cancel()
    reject(new Error('checkpoint_failed'))
    await vi.waitFor(() => expect(f.loop.busy).toBe(false))
    expect(f.onError).not.toHaveBeenCalled()
    expect(f.onDone).toHaveBeenCalledWith({ text: '', cancelled: true, turnLimit: false })
    expect(f.stream).toHaveBeenCalledTimes(1)
  })
})
