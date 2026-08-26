import { describe, expect, it, vi } from 'vitest'
import type {
  AgentMessage,
  AgentSkill,
  AgentStreamCallbacks,
  AgentToolCall,
  AgentTransport,
  ToolExecution,
} from '@wiswork/agent-core'
import {
  LEGACY_ERROR_MAX_BYTES,
  LegacyAgentRuntime,
  selectAgentRuntime,
  type AgentEvent,
} from '../src'

interface ScriptedTransport extends AgentTransport {
  callbacks: AgentStreamCallbacks[]
  cancels: number
  requests: AgentMessage[][]
}

function scriptedTransport(
  script: Array<(callbacks: AgentStreamCallbacks) => void>,
): ScriptedTransport {
  let turn = 0
  const transport: ScriptedTransport = {
    callbacks: [],
    cancels: 0,
    requests: [],
    stream(request, callbacks) {
      transport.callbacks.push(callbacks)
      transport.requests.push(request.messages)
      const step = script[turn++]
      if (step) queueMicrotask(() => step(callbacks))
      return {
        cancel: () => {
          transport.cancels++
          queueMicrotask(() => callbacks.onDone())
        },
      }
    },
  }
  return transport
}

function makeSkill(executeTool?: (call: AgentToolCall) => ToolExecution): AgentSkill {
  return {
    id: 'test',
    systemPrompt: 'system',
    tools: [{ name: 'edit', description: 'edit', inputSchema: { type: 'object' } }],
    executeTool: executeTool ?? (() => ({ output: 'ok', summary: 'edited', mutated: true })),
  }
}

const flush = () => new Promise((resolve) => setTimeout(resolve, 0))

describe('selectAgentRuntime', () => {
  it('maps the product-safe enhanced setting and migrates the one-release internal value', () => {
    expect(selectAgentRuntime('enhanced')).toBe('codex')
    expect(selectAgentRuntime('codex')).toBe('codex')
    for (const value of [
      'Enhanced',
      ' enhanced',
      'enhanced ',
      'standard',
      '',
      true,
      1,
      null,
      undefined,
    ]) {
      expect(selectAgentRuntime(value)).toBe('legacy')
    }
  })
})

describe('LegacyAgentRuntime', () => {
  it('opens a document session and rejects a duplicate until the first session closes', async () => {
    const runtime = new LegacyAgentRuntime()
    const options = {
      documentId: 'doc-1',
      transport: scriptedTransport([]),
      skill: makeSkill(),
    }
    const first = runtime.openSession(options)

    expect(first.documentId).toBe('doc-1')
    expect(first.runtimeKind).toBe('legacy')
    expect(first.busy).toBe(false)
    expect(() => runtime.openSession(options)).toThrow(/doc-1/)

    await Promise.all([first.close(), first.close()])
    const replacement = runtime.openSession(options)
    expect(replacement).not.toBe(first)
    await runtime.close()
  })

  it('normalizes cumulative text, tool lifecycle, the first mutation snapshot, turn end, and done', async () => {
    const transport = scriptedTransport([
      (callbacks) => {
        callbacks.onDelta('Working')
        callbacks.onDelta('...')
        callbacks.onToolCall({ id: 'call-1', name: 'edit', input: { value: 1 } })
        callbacks.onDone()
      },
      (callbacks) => {
        callbacks.onDelta('Finished')
        callbacks.onDone()
      },
    ])
    const runtime = new LegacyAgentRuntime<string>()
    const session = runtime.openSession({
      documentId: 'doc-events',
      transport,
      skill: makeSkill(),
      captureSnapshot: () => 'snapshot-before',
    })
    const events: AgentEvent<string>[] = []
    session.subscribe((event) => events.push(event))

    session.startTurn({ text: 'edit this', images: [{ base64: 'AAAA', mime: 'image/png' }] })
    expect(session.busy).toBe(true)
    await flush()
    await flush()

    expect(events).toEqual([
      { type: 'text', text: 'Working' },
      { type: 'text', text: 'Working...' },
      { type: 'tool-start', call: { id: 'call-1', name: 'edit', input: { value: 1 } } },
      {
        type: 'tool-executed',
        call: { id: 'call-1', name: 'edit', input: { value: 1 } },
        execution: { output: 'ok', summary: 'edited', mutated: true },
        snapshotBefore: 'snapshot-before',
      },
      { type: 'turn-end' },
      { type: 'text', text: 'Finished' },
      {
        type: 'done',
        result: { text: 'Finished', cancelled: false, turnLimit: false },
      },
    ])
    expect(transport.requests[0]![0]).toEqual({
      role: 'user',
      text: 'edit this',
      images: [{ base64: 'AAAA', mime: 'image/png' }],
    })
    expect(session.busy).toBe(false)
  })

  it('keeps empty and busy starts silent like AgentLoop', async () => {
    const transport = scriptedTransport([
      () => {
        // Remain busy until the test cancels.
      },
    ])
    const session = new LegacyAgentRuntime().openSession({
      documentId: 'doc-silent',
      transport,
      skill: makeSkill(),
    })
    const listener = vi.fn()
    session.subscribe(listener)

    session.startTurn({ text: '' })
    expect(transport.requests).toHaveLength(0)
    session.startTurn({ text: 'first' })
    session.startTurn({ text: 'ignored while busy' })
    await flush()
    expect(transport.requests).toHaveLength(1)
    expect(listener).not.toHaveBeenCalled()

    session.cancel()
    await flush()
    expect(transport.cancels).toBe(1)
    expect(listener).toHaveBeenCalledWith({
      type: 'done',
      result: { text: '', cancelled: true, turnLimit: false },
    })
  })

  it('isolates listeners and supports unsubscribe', async () => {
    const transport = scriptedTransport([
      (callbacks) => {
        callbacks.onDelta('safe')
        callbacks.onDone()
      },
    ])
    const session = new LegacyAgentRuntime().openSession({
      documentId: 'doc-listeners',
      transport,
      skill: makeSkill(),
    })
    const removed = vi.fn()
    const unsubscribe = session.subscribe(removed)
    unsubscribe()
    session.subscribe(() => {
      throw new Error('listener bug')
    })
    const healthy = vi.fn()
    session.subscribe(healthy)

    expect(() => session.startTurn({ text: 'go' })).not.toThrow()
    await flush()

    expect(removed).not.toHaveBeenCalled()
    expect(healthy).toHaveBeenNthCalledWith(1, { type: 'text', text: 'safe' })
    expect(healthy).toHaveBeenNthCalledWith(2, {
      type: 'done',
      result: { text: 'safe', cancelled: false, turnLimit: false },
    })
  })

  it('bounds legacy failures in UTF-8 bytes, emits a stable code, and lets AgentLoop own recovery', async () => {
    const secretPrefix = 'DO_NOT_LEAK_PROMPT_OR_TOKEN'
    const longError = `${secretPrefix}${'界'.repeat(LEGACY_ERROR_MAX_BYTES)}`
    const transport = scriptedTransport([
      (callbacks) => callbacks.onError(longError),
      (callbacks) => {
        callbacks.onDelta('recovered')
        callbacks.onDone()
      },
    ])
    const session = new LegacyAgentRuntime().openSession({
      documentId: 'doc-errors',
      transport,
      skill: makeSkill(),
    })
    const events: AgentEvent[] = []
    session.subscribe((event) => events.push(event))

    session.startTurn({ text: 'failed prompt' })
    await flush()
    const error = events[0]
    expect(error).toMatchObject({ type: 'error', code: 'legacy_turn_failed' })
    expect(
      Buffer.byteLength(error!.type === 'error' ? error.message : '', 'utf8'),
    ).toBeLessThanOrEqual(LEGACY_ERROR_MAX_BYTES)
    expect(error!.type === 'error' ? error.message : '').toBe('Legacy turn failed.')
    expect(error!.type === 'error' ? error.message : '').not.toContain(secretPrefix)
    expect(events.some((event) => event.type === 'done')).toBe(false)

    session.startTurn({ text: 'next prompt' })
    await flush()
    expect(transport.requests[1]).toEqual([{ role: 'user', text: 'next prompt' }])
    expect(events.at(-1)).toEqual({
      type: 'done',
      result: { text: 'recovered', cancelled: false, turnLimit: false },
    })
  })

  it('cancels and resets exactly once when close calls race', async () => {
    const transport = scriptedTransport([
      () => {
        // Remain busy until close resets the loop.
      },
    ])
    const runtime = new LegacyAgentRuntime()
    const session = runtime.openSession({
      documentId: 'doc-close',
      transport,
      skill: makeSkill(),
    })
    const listener = vi.fn()
    session.subscribe(listener)
    session.startTurn({ text: 'pending' })
    await flush()

    const first = session.close()
    const second = session.close()
    expect(second).toBe(first)
    await Promise.all([first, second, runtime.close(), runtime.close()])

    expect(transport.cancels).toBe(1)
    expect(session.busy).toBe(false)
    expect(listener).not.toHaveBeenCalled()
    session.startTurn({ text: 'ignored after close' })
    expect(transport.requests).toHaveLength(1)
  })
})
