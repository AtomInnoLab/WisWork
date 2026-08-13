import { describe, expect, it } from 'vitest'
import { AgentPanelSession } from '../src/renderer/ai/agent-panel-session.js'

function deferred<T>() {
  let resolve!: (value: T) => void
  const promise = new Promise<T>((done) => {
    resolve = done
  })
  return { promise, resolve }
}

describe('AI panel project and run isolation', () => {
  it('invalidates a cancelled run while keeping its completion eligible to settle busy UI', () => {
    const session = new AgentPanelSession('project-a')
    const loop = {}
    session.attachLoop(loop, 'project-a')
    const run = session.beginRun(loop)
    expect(session.acceptsRun(loop, run)).toBe(true)

    session.cancelRun(loop, run)
    expect(session.acceptsRun(loop, run)).toBe(false)
    expect(session.acceptsRunResult(loop, run)).toBe(false)
    expect(session.acceptsCompletion(loop, run)).toBe(true)
  })

  it('accepts a deferred proposal after normal completion until a newer run starts', async () => {
    const session = new AgentPanelSession('project-a')
    const loop = {}
    session.attachLoop(loop, 'project-a')
    const run = session.beginRun(loop)
    const proposal = deferred<string>()
    const committed: string[] = []
    const completion = proposal.promise.then((value) => {
      if (session.acceptsRunResult(loop, run)) committed.push(value)
    })

    session.finishRun(loop, run)
    proposal.resolve('proposal-1')
    await completion
    expect(committed).toEqual(['proposal-1'])

    session.beginRun(loop)
    expect(session.acceptsRunResult(loop, run)).toBe(false)
  })

  it('rejects every late A event after B starts, including deferred proposal and done', async () => {
    const session = new AgentPanelSession('project-a')
    const loopA = {}
    const scopeA = session.attachLoop(loopA, 'project-a')
    const runA = session.beginRun(loopA)
    const late = deferred<string>()
    const state = { text: 'B text', tools: ['B tool'], proposal: 'B proposal', done: false }
    const completion = late.promise.then((value) => {
      if (session.acceptsRunResult(loopA, runA)) state.proposal = value
    })

    const loopB = {}
    session.attachLoop(loopB, 'project-b')
    const runB = session.beginRun(loopB)
    if (session.acceptsRun(loopA, runA)) state.text = 'A late text'
    if (session.acceptsRun(loopA, runA)) state.tools.push('A late tool')
    if (session.acceptsCompletion(loopA, runA)) state.done = true
    late.resolve('A late proposal')
    await completion
    expect(state).toEqual({
      text: 'B text',
      tools: ['B tool'],
      proposal: 'B proposal',
      done: false,
    })
    expect(session.acceptsRun(loopB, runB)).toBe(true)
    expect(session.acceptsLoopProject(loopA, scopeA)).toBe(false)
    expect(session.timelineId(runA, 'tool-1')).not.toBe(session.timelineId(runB, 'tool-1'))
  })

  it('allows chat restore only for the current loop before its first run', async () => {
    const session = new AgentPanelSession('project-a')
    const loopA = {}
    const scopeA = session.attachLoop(loopA, 'project-a')
    const load = deferred<string[]>()
    const restored: string[] = []
    const completion = load.promise.then((messages) => {
      if (session.canRestoreChat(loopA, scopeA)) restored.push(...messages)
    })

    expect(session.canSend(loopA, 'loading')).toBe(false)
    expect(session.canSend(loopA, 'error')).toBe(true)
    session.beginRun(loopA)
    load.resolve(['late history'])
    await completion
    expect(restored).toEqual([])
  })
})
